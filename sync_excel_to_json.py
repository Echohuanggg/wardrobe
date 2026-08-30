# -*- coding: utf-8 -*-
"""
衣物清单同步脚本：以 Excel 为权威源，自动更新 wardrobe.json 与 thumbnails 缩略图。
用法：python sync_excel_to_json.py
查找顺序：1) 脚本同目录下的 衣物识别清单.xlsx  2) 上级目录下的 衣物识别清单.xlsx
"""
import html
import json
import os
import re
import shutil
import sys
import zipfile
from datetime import date

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_NAMES = ["衣物识别清单.xlsx"]
OUT_JSON = os.path.join(BASE_DIR, "wardrobe.json")
THUMB_DIR = os.path.join(BASE_DIR, "thumbnails")
BAK_JSON = os.path.join(BASE_DIR, "wardrobe.json.bak")
OUTFITS_JSON = os.path.join(BASE_DIR, "outfits.json")
OUTFITS_DIR = os.path.join(BASE_DIR, "outfits")
BAK_OUTFITS = os.path.join(BASE_DIR, "outfits.json.bak")
OUTFIT_SHEET = "穿搭样板"


def find_excel():
    """依次在脚本目录、上级目录查找 Excel"""
    for base in (BASE_DIR, os.path.dirname(BASE_DIR)):
        for name in EXCEL_NAMES:
            p = os.path.join(base, name)
            if os.path.exists(p):
                return p
    return None


def read_excel_rows(excel_path):
    """用 zipfile 方式解析 xlsx：读取共享字符串 + 单元格文本，避免依赖 openpyxl"""
    try:
        from openpyxl import load_workbook
        return read_with_openpyxl(excel_path)
    except Exception:
        return read_with_zip(excel_path)


def read_with_openpyxl(excel_path):
    from openpyxl import load_workbook
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows())  # Cell 对象，保留物理行号
    wb.close()
    if not rows:
        return []
    header = [str(c.value).strip() if c.value is not None else "" for c in rows[0]]
    col = {name: i for i, name in enumerate(header)}
    items = []
    for r in rows[1:]:
        if r[0].value is None:
            continue
        try:
            rid = int(r[0].value)
        except (ValueError, TypeError):
            continue
        vals = [c.value for c in r]
        excel_row = r[0].row
        items.append({
            "id": rid,
            "excel_row": excel_row,
            "category": str(vals[col["类别"]] or "").strip() if "类别" in col else "",
            "color": str(vals[col["颜色"]] or "").strip() if "颜色" in col else "",
            "style": str(vals[col["款式"]] or "").strip() if "款式" in col else "",
            "brand": str(vals[col["品牌"]] or "").strip() if "品牌" in col else "",
            "material": str(vals[col["材质"]] or "").strip() if "材质" in col else "",
            "purchase_date": str(vals[col["购买日期"]] or "").strip() if "购买日期" in col else "",
            "seasons_raw": str(vals[col["季节"]] or "").strip() if "季节" in col else "",
            "note": str(vals[col["备注"]] or "").strip() if "备注" in col else "",
        })
    return items


def read_with_zip(excel_path):
    """降级方案：直接解析 sheet XML + sharedStrings，不依赖第三方库"""
    items = []
    try:
        with zipfile.ZipFile(excel_path) as z:
            shared = []
            if "xl/sharedStrings.xml" in z.namelist():
                sxml = z.read("xl/sharedStrings.xml").decode("utf-8", errors="ignore")
                # 逐 si 提取全部 <t> 文本拼接（兼容富文本与 <phoneticPr/> 干扰）
                for si in re.findall(r"<si>(.*?)</si>", sxml, re.S):
                    texts = re.findall(r"<t[^>]*>([^<]*)</t>", si)
                    shared.append("".join(texts) if texts else "")
            sheet_name = [n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml", n)][0]
            xml = z.read(sheet_name).decode("utf-8", errors="ignore")
            rows_data = []  # [(excel_row, {col: val})]
            # 按行解析
            for row_m in re.finditer(r"<row[^>]*r=\"(\d+)\"[^>]*>(.*?)</row>", xml, re.S):
                excel_row = int(row_m.group(1))
                cells = {}
                # 整体捕获属性串，再从中提取 t 类型（避免非贪婪+可选组导致的 t 丢失）
                for c in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*)>(?:<v>([^<]*)</v>)?(?:<is><t[^>]*>([^<]*)</t></is>)?</c>', row_m.group(2)):
                    col_letter, attrs, v, inline = c.group(1), c.group(2), c.group(3), c.group(4)
                    val = ""
                    if inline is not None:
                        val = inline
                    else:
                        tm = re.search(r'\bt="([^"]+)"', attrs)
                        if tm and tm.group(1) == "s" and v is not None:
                            val = shared[int(v)]
                        elif v is not None:
                            val = v
                    cells[col_letter] = val.strip()
                rows_data.append((excel_row, cells))
            if not rows_data:
                return items
            # 表头行：列名 -> 列字母（增列后仍能正确对应）
            header_cells = dict(rows_data[0][1])
            col = {}
            for letter, name in header_cells.items():
                col[str(name).strip()] = letter
            for excel_row, cells in rows_data[1:]:
                try:
                    rid = int(cells.get(col.get("编号", "A"), ""))
                except (ValueError, TypeError):
                    continue
                items.append({
                    "id": rid,
                    "excel_row": excel_row,
                    "category": cells.get(col.get("类别", "B"), ""),
                    "color": cells.get(col.get("颜色", "C"), ""),
                    "style": cells.get(col.get("款式", "D"), ""),
                    "brand": cells.get(col.get("品牌", "E"), ""),
                    "material": cells.get(col.get("材质", "F"), ""),
                    "purchase_date": cells.get(col.get("购买日期", "G"), ""),
                    "seasons_raw": cells.get(col.get("季节", "H"), ""),
                    "note": cells.get(col.get("备注", "I"), ""),
                })
    except Exception as e:
        print(f"[错误] 解析 Excel 失败: {e}")
        sys.exit(1)
    return items


def _load_shared_strings(z):
    """读取 xlsx 共享字符串（兼容富文本与 <phoneticPr/>）"""
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        sxml = z.read("xl/sharedStrings.xml").decode("utf-8", errors="ignore")
        for si in re.findall(r"<si>(.*?)</si>", sxml, re.S):
            texts = re.findall(r"<t[^>]*>([^<]*)</t>", si)
            shared.append(html.unescape("".join(texts)) if texts else "")
    return shared


def _parse_sheet_xml_rows(xml, shared):
    """通用解析 sheet XML -> [(excel_row, {col: val})]"""
    rows_data = []
    for row_m in re.finditer(r"<row[^>]*r=\"(\d+)\"[^>]*>(.*?)</row>", xml, re.S):
        excel_row = int(row_m.group(1))
        cells = {}
        for c in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*)>(?:<v>([^<]*)</v>)?(?:<is><t[^>]*>([^<]*)</t></is>)?</c>', row_m.group(2)):
            col_letter, attrs, v, inline = c.group(1), c.group(2), c.group(3), c.group(4)
            val = ""
            if inline is not None:
                val = inline
            else:
                tm = re.search(r'\bt="([^"]+)"', attrs)
                if tm and tm.group(1) == "s" and v is not None:
                    val = shared[int(v)]
                elif v is not None:
                    val = v
            cells[col_letter] = html.unescape(val).strip()
        rows_data.append((excel_row, cells))
    return rows_data


def _norm_xl_path(target):
    """把 rels 的 Target 规范化为以 xl/ 开头的包内路径（兼容相对/绝对/../）"""
    t = target.replace("\\", "/").lstrip("/")
    if t.startswith("xl/"):
        return t
    if t.startswith("../"):
        return "xl/" + t.replace("../", "", 1)
    return "xl/" + t


def _zip_find_sheet(z, sheet_name):
    """返回指定名称 sheet 的 xml 文件名，如 xl/worksheets/sheet2.xml"""
    names = z.namelist()
    if "xl/workbook.xml" not in names:
        return None
    wxml = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
    sheet_rid = None
    for m in re.finditer(r"<sheet[^>]*name=\"([^\"]+)\"[^>]*r:id=\"([^\"]+)\"[^>]*/>", wxml):
        if m.group(1) == sheet_name:
            sheet_rid = m.group(2)
            break
    if not sheet_rid:
        return None
    rels_path = "xl/_rels/workbook.xml.rels"
    if rels_path not in names:
        return None
    rxml = z.read(rels_path).decode("utf-8", errors="ignore")
    for rm in re.finditer(r"<Relationship[^>]*/>", rxml):
        if f'Id="{sheet_rid}"' in rm.group(0):
            tgt_m = re.search(r'Target="([^"]+)"', rm.group(0))
            if tgt_m:
                return _norm_xl_path(tgt_m.group(1))
    return None


def _zip_find_drawing(z, sheet_xml_name):
    """解析 sheet xml 中的 drawing r:id -> drawing xml 文件名"""
    names = z.namelist()
    xml = z.read(sheet_xml_name).decode("utf-8", errors="ignore")
    m = re.search(r"<drawing[^>]*r:id=\"([^\"]+)\"", xml)
    if not m:
        return None
    rid = m.group(1)
    rels_path = sheet_xml_name.replace("worksheets/", "worksheets/_rels/") + ".rels"
    if rels_path not in names:
        return None
    rxml = z.read(rels_path).decode("utf-8", errors="ignore")
    for rm in re.finditer(r"<Relationship[^>]*/>", rxml):
        if f'Id="{rid}"' in rm.group(0):
            tgt_m = re.search(r'Target="([^"]+)"', rm.group(0))
            if tgt_m:
                return _norm_xl_path(tgt_m.group(1))
    return None


def read_outfits_openpyxl(excel_path):
    """openpyxl 路径：读取'穿搭样板' sheet 行 + 内嵌图对象"""
    from openpyxl import load_workbook
    wb = load_workbook(excel_path, data_only=True)
    if OUTFIT_SHEET not in wb.sheetnames:
        wb.close()
        return []
    ws = wb[OUTFIT_SHEET]
    img_map = {}
    for img in ws._images:
        img_map[img.anchor._from.row + 1] = img  # 0-based row -> Excel 行号
    items = []
    for row in ws.iter_rows(min_row=1, max_col=7):
        if row[0].row == 1:
            header = [str(c.value).strip() if c.value is not None else "" for c in row]
            col = {name: i for i, name in enumerate(header)}
            continue
        r0 = row[0].value
        if r0 is None:
            continue
        try:
            rid = int(r0)
        except (ValueError, TypeError):
            continue
        excel_row = row[0].row
        items.append({
            "id": rid,
            "category": str(row[col.get("类别", 1)].value or "").strip() if col.get("类别", 1) < 7 else "",
            "title": str(row[col.get("标题", 2)].value or "").strip() if col.get("标题", 2) < 7 else "",
            "season": str(row[col.get("季节", 3)].value or "").strip() if col.get("季节", 3) < 7 else "",
            "scene": str(row[col.get("适用场景", 4)].value or "").strip() if col.get("适用场景", 4) < 7 else "",
            "note": str(row[col.get("备注", 6)].value or "").strip() if col.get("备注", 6) < 7 else "",
            "image_obj": img_map.get(excel_row),
        })
    wb.close()
    return items


def read_outfits_zip(excel_path):
    """zip 降级路径：解析'穿搭样板' sheet 行 + 内嵌图 media 文件"""
    items = []
    with zipfile.ZipFile(excel_path) as z:
        names = z.namelist()
        shared = _load_shared_strings(z)
        sheet_xml = _zip_find_sheet(z, OUTFIT_SHEET)
        if not sheet_xml:
            return []
        xml = z.read(sheet_xml).decode("utf-8", errors="ignore")
        rows_data = _parse_sheet_xml_rows(xml, shared)
        if not rows_data:
            return items
        header_cells = dict(rows_data[0][1])
        col = {}
        for letter, name in header_cells.items():
            col[str(name).strip()] = letter
        img_map = {}
        drawing_xml = _zip_find_drawing(z, sheet_xml)
        if drawing_xml and drawing_xml in names:
            dxml = z.read(drawing_xml).decode("utf-8", errors="ignore")
            rel_path = drawing_xml.replace("drawings/", "drawings/_rels/") + ".rels"
            rels = {}
            if rel_path in names:
                rxml = z.read(rel_path).decode("utf-8", errors="ignore")
                for rm in re.finditer(r"<Relationship[^>]*/>", rxml):
                    rid_m = re.search(r'Id="([^"]+)"', rm.group(0))
                    tgt_m = re.search(r'Target="([^"]+)"', rm.group(0))
                    if rid_m and tgt_m:
                        rels[rid_m.group(1)] = tgt_m.group(1)
            anchors = re.findall(r"<(?:xdr:)?(?:oneCellAnchor|twoCellAnchor)[^>]*>(.*?)</(?:xdr:)?(?:oneCellAnchor|twoCellAnchor)>", dxml, re.S)
            for body in anchors:
                row_m = re.search(r"<(?:xdr:)?from>.*?<(?:xdr:)?row>(\d+)</(?:xdr:)?row>", body, re.S)
                embed_m = re.search(r'r:embed="([^"]+)"', body)
                if not row_m or not embed_m:
                    continue
                excel_row = int(row_m.group(1)) + 1
                target = rels.get(embed_m.group(1), "")
                media_file = _norm_xl_path(target)
                if media_file in names:
                    img_map[excel_row] = media_file
        for excel_row, cells in rows_data[1:]:
            try:
                rid = int(cells.get(col.get("编号", "A"), ""))
            except (ValueError, TypeError):
                continue
            items.append({
                "id": rid,
                "category": cells.get(col.get("类别", "B"), ""),
                "title": cells.get(col.get("标题", "C"), ""),
                "season": cells.get(col.get("季节", "D"), ""),
                "scene": cells.get(col.get("适用场景", "E"), ""),
                "note": cells.get(col.get("备注", "G"), ""),
                "image_media": img_map.get(excel_row, ""),
            })
    return items


def sync_outfits(excel_path):
    """同步穿搭样板：Excel '穿搭样板' sheet -> outfits/ 图片 + outfits.json，旧数据自动备份"""
    use_zip = False
    try:
        from openpyxl import load_workbook  # noqa
        outfits = read_outfits_openpyxl(excel_path)
    except Exception:
        use_zip = True
        outfits = read_outfits_zip(excel_path)
    if not outfits:
        return 0
    if os.path.exists(OUTFITS_JSON):
        try:
            with open(OUTFITS_JSON, encoding="utf-8") as f:
                json.load(f)
            shutil.copy2(OUTFITS_JSON, BAK_OUTFITS)
        except Exception:
            pass
    os.makedirs(OUTFITS_DIR, exist_ok=True)
    saved = 0
    new_items = []
    for o in outfits:
        img_path = ""
        if use_zip:
            media = o.get("image_media", "")
            if media:
                with zipfile.ZipFile(excel_path) as z:
                    data = z.read(media)
                ext = os.path.splitext(media)[1] or ".jpg"
                img_path = f"outfits/{o['id']:03d}{ext}"
                with open(os.path.join(BASE_DIR, img_path), "wb") as f:
                    f.write(data)
                saved += 1
        else:
            img = o.get("image_obj")
            if img is not None:
                ext = "." + (img.format or "jpg")
                img_path = f"outfits/{o['id']:03d}{ext}"
                with open(os.path.join(BASE_DIR, img_path), "wb") as f:
                    f.write(img._data())
                saved += 1
        new_items.append({
            "id": o["id"],
            "category": o.get("category", ""),
            "title": o.get("title", ""),
            "season": o.get("season", ""),
            "scene": o.get("scene", ""),
            "note": o.get("note", ""),
            "image": img_path,
        })
    with open(OUTFITS_JSON, "w", encoding="utf-8") as f:
        json.dump(new_items, f, ensure_ascii=False, indent=2)
    print(f"[穿搭] 穿搭样板同步完成：{len(new_items)} 套，导出图片 {saved} 张")
    return len(new_items)


def extract_images(excel_path):
    """从 Excel 内嵌图导出缩略图：解析 drawing1.xml 锚点 -> Excel行号 -> [media文件]"""
    exported = {}
    try:
        with zipfile.ZipFile(excel_path) as z:
            names = z.namelist()
            drawing = [n for n in names if re.match(r"xl/drawings/drawing\d+\.xml", n)]
            if not drawing:
                return exported
            xml = z.read(drawing[0]).decode("utf-8", errors="ignore")
            # rels 映射 rId -> media 文件（属性顺序不固定，逐块提取 Id/Target）
            rel_path = drawing[0].replace("drawings/", "drawings/_rels/") + ".rels"
            rels = {}
            if rel_path in names:
                rxml = z.read(rel_path).decode("utf-8", errors="ignore")
                for rm in re.finditer(r"<Relationship[^>]*/>", rxml):
                    rid_m = re.search(r'Id="([^"]+)"', rm.group(0))
                    tgt_m = re.search(r'Target="([^"]+)"', rm.group(0))
                    if rid_m and tgt_m:
                        rels[rid_m.group(1)] = tgt_m.group(1)
            # oneCellAnchor / twoCellAnchor 锚点（兼容 xdr: 前缀与默认命名空间、带属性标签）
            anchors = re.findall(r"<(?:xdr:)?(?:oneCellAnchor|twoCellAnchor)[^>]*>(.*?)</(?:xdr:)?(?:oneCellAnchor|twoCellAnchor)>", xml, re.S)
            for body in anchors:
                row_m = re.search(r"<(?:xdr:)?from>.*?<(?:xdr:)?row>(\d+)</(?:xdr:)?row>", body, re.S)
                embed_m = re.search(r'r:embed="([^"]+)"', body)
                if not row_m or not embed_m:
                    continue
                excel_row = int(row_m.group(1)) + 1  # row 是 0-based，+1 = Excel 行号
                rid = embed_m.group(1)
                target = rels.get(rid, "")
                media_file = target.lstrip("/").replace("../", "xl/")
                if media_file in names:
                    exported.setdefault(excel_row, []).append(media_file)
    except Exception as e:
        print(f"[警告] 提取内嵌图失败: {e}")
    return exported


def parse_seasons(raw):
    """把季节文本转数组：'春夏秋'->['春','夏','秋']，'四季'->全季"""
    if not raw or raw == "无":
        return []
    if "四季" in raw or raw == "全年":
        return ["春", "夏", "秋", "冬"]
    result = []
    for s in "春夏秋冬":
        if s in raw:
            result.append(s)
    return result


def parse_purchase_date(raw):
    """购买日期：数字 202608 -> '202608'，空 -> ''"""
    if not raw:
        return ""
    if raw in ("无", "未知", "N/A"):
        return ""
    return str(raw).strip()


def main():
    excel_path = find_excel()
    if not excel_path:
        print("[错误] 未找到 衣物识别清单.xlsx（脚本目录或上级目录）")
        sys.exit(1)
    print(f"[1/4] 读取 Excel: {excel_path}")

    rows = read_excel_rows(excel_path)
    if not rows:
        print("[错误] Excel 无数据行")
        sys.exit(1)
    print(f"[2/4] Excel 共 {len(rows)} 件衣物（编号 {min(r['id'] for r in rows)}~{max(r['id'] for r in rows)}）")

    # 读取旧 JSON（保留 name/occasion/created_at 等字段）
    old_map = {}
    if os.path.exists(OUT_JSON):
        try:
            with open(OUT_JSON, encoding="utf-8") as f:
                old_map = {int(item["id"]): item for item in json.load(f)}
        except Exception:
            old_map = {}
    if old_map:
        shutil.copy2(OUT_JSON, BAK_JSON)
        print(f"    旧 JSON 已备份: {BAK_JSON}")

    # 提取内嵌图
    image_map = extract_images(excel_path)  # Excel行号 -> [media文件]
    # 关键修复：锚点行号与编号存在偏移（如整行删除后 row=id-1），
    # 必须先把行号映射到该行单元格里的编号，再按编号取图，禁止直接用编号查行号键。
    row_to_id = {r["excel_row"]: r["id"] for r in rows if r.get("excel_row")}
    # 多图行人工指定取第几张（0-based）：历史编辑残留图需视觉确认后选择
    multi_pick = {9: 0, 37: 1, 42: 0, 71: 1}
    image_by_id = {}
    for excel_row, medias in image_map.items():
        rid = row_to_id.get(excel_row)
        if rid is None:
            continue
        pick = multi_pick.get(rid, 0)
        if pick >= len(medias):
            pick = 0
        image_by_id[rid] = medias[pick]
    print(f"[3/4] Excel 内嵌缩略图 {len(image_map)} 张（按编号映射 {len(image_by_id)} 条）")

    os.makedirs(THUMB_DIR, exist_ok=True)
    new_items = []
    exported_count = 0
    for r in sorted(rows, key=lambda x: x["id"]):
        rid = r["id"]
        old = old_map.get(rid, {})
        name = old.get("name") or f"{r['color']}{r['style']}".strip() or f"衣物{rid:03d}"
        # 品牌归一化：以 Excel 权威源为准，'无' 或误填数字 19 一律归一化为空串
        brand = str(r["brand"] or "").strip()
        if brand in ("无", "19"):
            brand = ""
        seasons = parse_seasons(r["seasons_raw"])
        purchase_date = parse_purchase_date(r["purchase_date"])
        # 缩略图：优先从 Excel 内嵌图导出（按编号取图）
        img_file = f"thumbnails/{rid:03d}.png"
        media = image_by_id.get(rid)
        if media:
            try:
                with zipfile.ZipFile(excel_path) as z:
                    data = z.read(media)
                _write_thumbnail(os.path.join(BASE_DIR, img_file), data)
                exported_count += 1
            except Exception as e:
                print(f"    [警告] 导出 #{rid} 图片失败: {e}")
        elif not os.path.exists(os.path.join(BASE_DIR, img_file)):
            print(f"    [提示] #{rid} 无内嵌图且无现有缩略图，生成白底占位图")
            _make_placeholder(os.path.join(BASE_DIR, img_file))

        item = {
            "id": rid,
            "name": name,
            "category": r["category"],
            "color": r["color"],
            "style": r["style"],
            "brand": brand,
            "material": r["material"],
            "seasons": seasons,
            "occasion": old.get("occasion", ""),
            "purchase_date": purchase_date,
            "note": r["note"],
            "created_at": old.get("created_at", date.today().isoformat()),
            "image": img_file,
        }
        new_items.append(item)

    # 写 JSON
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(new_items, f, ensure_ascii=False, indent=2)

    # 验证
    with open(OUT_JSON, encoding="utf-8") as f:
        check = json.load(f)
    missing_img = [i["id"] for i in check if not os.path.exists(os.path.join(BASE_DIR, i["image"]))]
    print(f"[4/4] 同步完成！wardrobe.json 共 {len(check)} 条，导出/刷新缩略图 {exported_count} 张")
    if missing_img:
        print(f"    [警告] 以下编号缺缩略图文件: {missing_img}")
    else:
        print("    所有缩略图文件均可访问")
    # 穿搭样板同步（新增 worksheet 解析）
    try:
        sync_outfits(excel_path)
    except Exception as e:
        print(f"[警告] 穿搭样板同步失败: {e}")
    print("数据已就绪，可以启动页面。")


def _write_thumbnail(path, data):
    """把内嵌图字节写为 160x160 白底居中 PNG 缩略图；PIL 不可用时直接写原图"""
    try:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(data)).convert("RGBA")
        img.thumbnail((160, 160), Image.LANCZOS)
        canvas = Image.new("RGB", (160, 160), (255, 255, 255))
        x = (160 - img.width) // 2
        y = (160 - img.height) // 2
        canvas.paste(img, (x, y), img)
        canvas.save(path, "PNG")
    except Exception:
        with open(path, "wb") as f:
            f.write(data)


def _make_placeholder(path):
    """生成 160x160 白底 PNG 占位图（纯 Python，无 PIL 依赖）"""
    try:
        from PIL import Image
        img = Image.new("RGB", (160, 160), (255, 255, 255))
        img.save(path)
    except Exception:
        # 无 PIL 时写最小 PNG
        import struct, zlib
        w = h = 160
        raw = b"".join(b"\x00" + b"\xff\xff\xff" * w for _ in range(h))
        def chunk(t, d):
            c = t + d
            return struct.pack(">I", len(d)) + c + struct.pack(">I", zlib.crc32(c))
        png = b"\x89PNG\r\n\x1a\n"
        png += chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        png += chunk(b"IDAT", zlib.compress(raw))
        png += chunk(b"IEND", b"")
        with open(path, "wb") as f:
            f.write(png)


if __name__ == "__main__":
    main()
