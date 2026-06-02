import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "ノートPC比較表"

# ---- カラーパレット ----
COLOR_HEADER_BG   = "1F4E79"   # ダークブルー
COLOR_HEADER_FONT = "FFFFFF"
COLOR_MUST_BG     = "FCE4D6"   # 薄オレンジ（マスト条件行）
COLOR_OPT_BG      = "EBF3FB"   # 薄ブルー（任意条件行）
COLOR_GOOD        = "C6EFCE"   # 緑（条件クリア）
COLOR_BAD         = "FFCCCC"   # 赤（条件未達）
COLOR_NEUTRAL     = "FFFACD"   # 薄黄（要確認）
COLOR_SECTION_BG  = "D6E4F0"   # セクション見出し

# ---- ボーダー定義 ----
thin  = Side(style="thin",   color="999999")
thick = Side(style="medium", color="1F4E79")
thin_border  = Border(left=thin,  right=thin,  top=thin,  bottom=thin)
thick_border = Border(left=thick, right=thick, top=thick, bottom=thick)

def header_style():
    return Font(name="Meiryo UI", bold=True, color=COLOR_HEADER_FONT, size=11)

def body_style(bold=False, size=10):
    return Font(name="Meiryo UI", bold=bold, size=size)

def fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

# ---- モデルデータ ----
models = [
    {
        "brand":    "VAIO",
        "model":    "SX14-R\n(2026年5月)",
        "price":    "159,800円〜\n（特別価格）",
        "price_ok": "○",
        "cpu":      "Intel Core Ultra 3\n(Copilot+ PC)",
        "cpu_ok":   "○",
        "mem":      "16GB",
        "mem_ok":   "○",
        "ssd":      "256GB〜",
        "ssd_ok":   "○",
        "hello":    "○\n（9.2MP IRカメラ）",
        "hello_ok": "○",
        "usb_pd":   "○\n（USB-C PD対応）",
        "usb_ok":   "○",
        "display":  "14.0型",
        "weight":   "約958g",
        "origin":   "日本（VAIO）",
        "note":     "超軽量・Copilot+ PC\nセール価格あり確認推奨",
    },
    {
        "brand":    "富士通",
        "model":    "FMV Note A\n(2026年1月)",
        "price":    "169,800円〜",
        "price_ok": "△",
        "cpu":      "Intel Core\nUltra 5 / AMD Ryzen",
        "cpu_ok":   "○",
        "mem":      "16GB",
        "mem_ok":   "○",
        "ssd":      "256GB〜",
        "ssd_ok":   "○",
        "hello":    "○\n（IRカメラ搭載）",
        "hello_ok": "○",
        "usb_pd":   "○\n（65W USB-C PD）",
        "usb_ok":   "○",
        "display":  "16.0型",
        "weight":   "約1.95kg",
        "origin":   "日本（富士通）",
        "note":     "3年保証無料\nAI機能搭載\n169,800円はギリギリ条件内",
    },
    {
        "brand":    "NEC",
        "model":    "LAVIE Direct N15\n(2026年春)",
        "price":    "要確認\n（公式直販サイト）",
        "price_ok": "△",
        "cpu":      "AMD Ryzen 5 7535HS\n（i5相当）",
        "cpu_ok":   "○",
        "mem":      "16GB",
        "mem_ok":   "○",
        "ssd":      "256GB / 512GB",
        "ssd_ok":   "○",
        "hello":    "○\n（IRカメラ）",
        "hello_ok": "○",
        "usb_pd":   "○\n（USB-C PD対応）",
        "usb_ok":   "○",
        "display":  "15.6型",
        "weight":   "約1.65kg",
        "origin":   "日本（NEC）",
        "note":     "国産PC\n価格は直販サイトで要確認",
    },
    {
        "brand":    "dynabook",
        "model":    "Cシリーズ\n(2026年春)",
        "price":    "要確認\n（公式直販サイト）",
        "price_ok": "△",
        "cpu":      "AMD Ryzen 5\n（i5相当）",
        "cpu_ok":   "○",
        "mem":      "16GB\n（DDR4-3200）",
        "mem_ok":   "○",
        "ssd":      "256GB",
        "ssd_ok":   "○",
        "hello":    "○\n（IRカメラ）",
        "hello_ok": "○",
        "usb_pd":   "○\n（USB 3.2 Gen2 Type-C）",
        "usb_ok":   "○",
        "display":  "15.6型",
        "weight":   "約1.9kg",
        "origin":   "日本（dynabook）",
        "note":     "旧東芝PC部門\n価格は直販サイトで要確認",
    },
    {
        "brand":    "HP",
        "model":    "EliteBook 640 G11\n(2026)",
        "price":    "140,000〜160,000円\n（推定）",
        "price_ok": "○",
        "cpu":      "Intel Core Ultra 5\n125U",
        "cpu_ok":   "○",
        "mem":      "16GB",
        "mem_ok":   "○",
        "ssd":      "256GB",
        "ssd_ok":   "○",
        "hello":    "○\n（IRカメラ搭載）",
        "hello_ok": "○",
        "usb_pd":   "○\n（USB4対応）",
        "usb_ok":   "○",
        "display":  "14.0型",
        "weight":   "約1.35kg",
        "origin":   "米国（HP）",
        "note":     "ビジネス向けモデル\nAMD/Intelが選択可能",
    },
]

# ---- シート構成 ----
# Row 1: タイトル
# Row 2: 凡例
# Row 3: ヘッダー（項目名）
# Row 4〜: データ

# タイトル行
ws.merge_cells("A1:H1")
title_cell = ws["A1"]
title_cell.value = "ノートPC比較表（2026年版）  ※Windows Hello対応カメラ搭載・税込17万円以下"
title_cell.font = Font(name="Meiryo UI", bold=True, size=14, color="1F4E79")
title_cell.alignment = Alignment(horizontal="center", vertical="center")
title_cell.fill = fill("DEEAF1")
ws.row_dimensions[1].height = 30

# 凡例行
ws.merge_cells("A2:H2")
legend_cell = ws["A2"]
legend_cell.value = "凡例: ○=条件クリア　△=要確認・ギリギリ　×=条件未達　　マスト条件=背景オレンジ　任意条件=背景ブルー"
legend_cell.font = Font(name="Meiryo UI", size=9, italic=True, color="444444")
legend_cell.alignment = Alignment(horizontal="left", vertical="center")
legend_cell.fill = fill("F2F2F2")
ws.row_dimensions[2].height = 18

# ---- ヘッダー行（行3）----
headers = ["比較項目", "区分"] + [f"{m['brand']}\n{m['model']}" for m in models]
for col_idx, h in enumerate(headers, start=1):
    cell = ws.cell(row=3, column=col_idx, value=h)
    cell.font = header_style()
    cell.fill = fill(COLOR_HEADER_BG)
    cell.alignment = center()
    cell.border = thin_border
ws.row_dimensions[3].height = 45

# ---- データ行定義 ----
# (項目名, 区分, data_key, value_key, is_must)
rows_def = [
    # マスト条件
    ("税込価格",           "マスト",  "price",   "price_ok", True),
    ("17万円以下",         "マスト",  None,      "price_ok", True),
    ("Windows Hello顔認証","マスト",  "hello",   "hello_ok", True),
    # 任意条件
    ("CPU",               "任意",    "cpu",     "cpu_ok",   False),
    ("i3相当以上",         "任意",    None,      "cpu_ok",   False),
    ("メモリ",             "任意",    "mem",     "mem_ok",   False),
    ("SSD",               "任意",    "ssd",     "ssd_ok",   False),
    ("USB給電（USB-C PD）","任意",    "usb_pd",  "usb_ok",   False),
    # 参考情報
    ("ディスプレイサイズ", "参考",    "display", None,       False),
    ("重量",              "参考",    "weight",  None,       False),
    ("製造国（本社）",     "参考",    "origin",  None,       False),
    ("備考",              "参考",    "note",    None,       False),
]

def ok_fill(val):
    if val == "○":
        return fill(COLOR_GOOD)
    elif val == "×":
        return fill(COLOR_BAD)
    elif val == "△":
        return fill(COLOR_NEUTRAL)
    return None

start_row = 4
for row_offset, (label, section, data_key, ok_key, is_must) in enumerate(rows_def):
    row_num = start_row + row_offset
    bg = COLOR_MUST_BG if is_must else (COLOR_OPT_BG if section == "任意" else "FFFFFF")

    # 項目名
    c = ws.cell(row=row_num, column=1, value=label)
    c.font = body_style(bold=True)
    c.fill = fill(bg)
    c.alignment = center()
    c.border = thin_border

    # 区分
    c2 = ws.cell(row=row_num, column=2, value=section)
    c2.font = body_style(size=9)
    c2.fill = fill(bg)
    c2.alignment = center()
    c2.border = thin_border
    if section == "マスト":
        c2.font = Font(name="Meiryo UI", bold=True, size=9, color="C00000")
    elif section == "任意":
        c2.font = Font(name="Meiryo UI", bold=True, size=9, color="1F4E79")
    else:
        c2.font = Font(name="Meiryo UI", size=9, color="595959")

    # 各モデルのデータ
    for col_offset, m in enumerate(models):
        col_num = 3 + col_offset
        if data_key:
            value = m[data_key]
        elif ok_key:
            value = m[ok_key]
        else:
            value = ""

        cell = ws.cell(row=row_num, column=col_num, value=value)
        cell.font = body_style()
        cell.alignment = center()
        cell.border = thin_border

        # 条件達成フィルは ok_key がある行のみ
        if ok_key and data_key is None:
            f = ok_fill(m[ok_key])
            if f:
                cell.fill = f
            else:
                cell.fill = fill(bg)
        else:
            cell.fill = fill(bg)

    ws.row_dimensions[row_num].height = 42

# ---- 列幅設定 ----
ws.column_dimensions["A"].width = 20
ws.column_dimensions["B"].width = 8
for i in range(len(models)):
    ws.column_dimensions[get_column_letter(3 + i)].width = 22

# ---- 注釈シート ----
ws2 = wb.create_sheet("調査メモ")
notes = [
    ["項目", "内容"],
    ["調査日", "2026年6月3日"],
    ["除外条件", "中国メーカー（Lenovo等）は比較対象外"],
    ["マスト条件①", "Windows Hello対応IRカメラ搭載（顔認証）"],
    ["マスト条件②", "税込17万円以下"],
    ["任意条件", "CPU i3相当以上、メモリ16GB、SSD 256GB、USB-C給電（Power Delivery）"],
    ["", ""],
    ["VAIO SX14-R", "通常価格は269,500円〜。特別価格159,800円の確認が必要。公式ストア: store.vaio.com"],
    ["富士通 FMV Note A", "169,800円〜は16GBモデルの最低価格。256GBモデルは要カスタム確認。公式: fmv.com"],
    ["NEC LAVIE N15", "公式直販サイト(nec-lavie.jp)で価格要確認。標準SSDは512GBが多い。"],
    ["dynabook Cシリーズ", "公式直販(dynabook.com/direct)で価格要確認。"],
    ["HP EliteBook 640 G11", "価格は推定。米国HPブランド。量販店・Amazon等で確認推奨。"],
    ["", ""],
    ["注意事項", "価格・スペックは変動することがあります。購入前に公式サイトで最新情報をご確認ください。"],
    ["△表示について", "「要確認」は公式サイトで価格やスペックの詳細確認が必要なことを示します。"],
]
for r, row_data in enumerate(notes, start=1):
    for c, val in enumerate(row_data, start=1):
        cell = ws2.cell(row=r, column=c, value=val)
        if r == 1:
            cell.font = Font(name="Meiryo UI", bold=True, size=11, color=COLOR_HEADER_FONT)
            cell.fill = fill(COLOR_HEADER_BG)
        else:
            cell.font = Font(name="Meiryo UI", size=10)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = thin_border
ws2.column_dimensions["A"].width = 25
ws2.column_dimensions["B"].width = 70
for r in range(1, len(notes) + 1):
    ws2.row_dimensions[r].height = 20

output_path = "/Users/naoto1/claude_work/github_work/ノートPC比較表.xlsx"
wb.save(output_path)
print(f"Saved: {output_path}")
