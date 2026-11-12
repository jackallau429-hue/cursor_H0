#!/usr/bin/env python3
"""
JIS-style 履歴書 (rirekisho) for Herman Mok — single A4 page.
Blank fields (DOB, full address, photo) left for the candidate to fill.
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import black, HexColor
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import shutil
import os

FONT = "/workspace/NotoSansJP-Regular.ttf"
pdfmetrics.registerFont(TTFont("JP", FONT))

PDF_OUT = "/workspace/Herman_Mok_Rirekisho.pdf"
DOCX_OUT = "/workspace/Herman_Mok_Rirekisho.docx"
LINE = HexColor("#333333")
LIGHT = HexColor("#666666")
HDR = HexColor("#F0F0F0")


def P(text, size=7.5, align=TA_LEFT, leading=None):
    return Paragraph(
        text,
        ParagraphStyle(
            "t",
            fontName="JP",
            fontSize=size,
            leading=leading or (size + 1.5),
            alignment=align,
        ),
    )


def draw_rirekisho():
    c = canvas.Canvas(PDF_OUT, pagesize=A4)
    W, H = A4
    ml, mr = 11 * mm, 11 * mm
    usable = W - ml - mr
    top = H - 8 * mm

    c.setFont("JP", 15)
    c.drawCentredString(W / 2, top, "履　歴　書")
    c.setFont("JP", 7.5)
    c.drawRightString(W - mr, top, "2026年11月12日 現在")

    y = top - 6 * mm

    # --- Header: personal + photo ---
    photo_w, photo_h = 28 * mm, 36 * mm
    info_w = usable - photo_w - 2 * mm

    info_data = [
        [P("ふりがな", 6, TA_CENTER), P("もく　ほいまん　／　Mok Hoi Man", 8)],
        [P("氏名", 7, TA_CENTER), P("<b>莫　海文（Herman Mok）</b>", 10)],
        [P("生年月日", 6.5, TA_CENTER), P("　　　年　　月　　日生（満　　歳）※ご記入ください", 7.5)],
        [P("性別", 6.5, TA_CENTER), P("　　　（任意）", 7.5)],
        [
            P("現住所", 6.5, TA_CENTER),
            P(
                "〒　　　－　　　　東京都江東区　　　　　※番地までご記入ください<br/>"
                "Tel: +81-80-9078-5842　Email: hoimanmok@gmail.com<br/>"
                "LinkedIn: linkedin.com/in/hoimanmok-herman　国籍・地域: 中国（香港）",
                7,
                leading=9.5,
            ),
        ],
    ]
    info = Table(
        info_data,
        colWidths=[15 * mm, info_w - 15 * mm],
        rowHeights=[6 * mm, 8.5 * mm, 6.5 * mm, 5.5 * mm, 18 * mm],
    )
    info.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "JP"),
                ("GRID", (0, 0), (-1, -1), 0.55, LINE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                ("BACKGROUND", (0, 0), (0, -1), HDR),
            ]
        )
    )
    iw, ih = info.wrap(info_w, 50 * mm)
    info.drawOn(c, ml, y - ih)

    px = ml + info_w + 2 * mm
    py = y - photo_h
    c.setStrokeColor(LINE)
    c.setLineWidth(0.7)
    c.rect(px, py, photo_w, photo_h)
    c.setFillColor(LIGHT)
    c.setFont("JP", 6.5)
    c.drawCentredString(px + photo_w / 2, py + photo_h / 2 + 3.5 * mm, "写真を貼る位置")
    c.drawCentredString(px + photo_w / 2, py + photo_h / 2 - 0.5 * mm, "縦40mm×横30mm")
    c.drawCentredString(px + photo_w / 2, py + photo_h / 2 - 4 * mm, "（任意）")
    c.setFillColor(black)

    y = min(y - ih, py) - 3.5 * mm

    # --- 学歴・職歴 ---
    c.setFont("JP", 9)
    c.drawString(ml, y, "学歴・職歴")
    y -= 1.5 * mm

    # Compact: one line per event, short role note on same/next as needed
    hist = [
        [P("<b>年</b>", 6.5, TA_CENTER), P("<b>月</b>", 6.5, TA_CENTER), P("<b>学歴・職歴</b>", 6.5, TA_CENTER)],
        [P("", 7), P("", 7), P("<b>【学歴】</b>", 7.5)],
        [P("2011", 7.5, TA_CENTER), P("9", 7.5, TA_CENTER), P("香港理工大学（The Hong Kong Polytechnic University） 入学", 7.5)],
        [P("2015", 7.5, TA_CENTER), P("7", 7.5, TA_CENTER), P("同大学 計算機学 優等学士課程（副専攻：会計学） 卒業", 7.5)],
        [P("", 7), P("", 7), P("以上", 7.5, TA_CENTER)],
        [P("", 7), P("", 7), P("<b>【職歴】</b>", 7.5)],
        [P("2016", 7.5, TA_CENTER), P("11", 7.5, TA_CENTER), P("Hong Kong Life Insurance Limited 入社（プログラマー）", 7.5)],
        [P("2018", 7.5, TA_CENTER), P("5", 7.5, TA_CENTER), P("一身上の都合により退社", 7.5)],
        [P("2018", 7.5, TA_CENTER), P("6", 7.5, TA_CENTER), P("KBQuest Hong Kong Ltd. 入社（ソフトウェアエンジニア／BI）", 7.5)],
        [P("2020", 7.5, TA_CENTER), P("10", 7.5, TA_CENTER), P("一身上の都合により退社", 7.5)],
        [
            P("2020", 7.5, TA_CENTER),
            P("10", 7.5, TA_CENTER),
            P("Hospital Authority（病院管理局）へ Manpower Services 経由で就業", 7),
        ],
        [P("", 7), P("", 7), P("契約アナリストプログラマー（AI Systems Team）", 7)],
        [P("2021", 7.5, TA_CENTER), P("10", 7.5, TA_CENTER), P("契約期間満了につき退職", 7.5)],
        [
            P("2021", 7.5, TA_CENTER),
            P("10", 7.5, TA_CENTER),
            P("Securities and Futures Commission（香港証券先物委員会）へ", 7),
        ],
        [P("", 7), P("", 7), P("NewRecs Consultancy Ltd 経由で就業（システムアナリスト／BI）", 7)],
        [P("2026", 7.5, TA_CENTER), P("6", 7.5, TA_CENTER), P("一身上の都合により退社", 7.5)],
        [P("2026", 7.5, TA_CENTER), P("7", 7.5, TA_CENTER), P("日本国 東京都へ転居　現在に至る（求職活動中）", 7.5)],
        [P("", 7), P("", 7), P("以上", 7.5, TA_CENTER)],
    ]
    col_w = [14 * mm, 9 * mm, usable - 23 * mm]
    hist_t = Table(hist, colWidths=col_w)
    hist_t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "JP"),
                ("GRID", (0, 0), (-1, -1), 0.45, LINE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 1.2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1.2),
                ("BACKGROUND", (0, 0), (-1, 0), HDR),
            ]
        )
    )
    tw, th = hist_t.wrap(usable, 140 * mm)
    hist_t.drawOn(c, ml, y - th)
    y = y - th - 3 * mm

    # --- 免許・資格 ---
    c.setFont("JP", 9)
    c.drawString(ml, y, "免許・資格")
    y -= 1.5 * mm

    cert = [
        [P("<b>年</b>", 6.5, TA_CENTER), P("<b>月</b>", 6.5, TA_CENTER), P("<b>免許・資格</b>", 6.5, TA_CENTER)],
        [P("2019", 7.5, TA_CENTER), P("", 7), P("Microsoft Data Platform & Analytics Foundational（MPN14354）", 7)],
        [P("2019", 7.5, TA_CENTER), P("", 7), P("Microsoft SQL Server 2016 – Mission Critical and Hybrid（MPN14355）", 7)],
        [P("2019", 7.5, TA_CENTER), P("", 7), P("Splunk Enterprise Certified Admin", 7.5)],
        [P("2024", 7.5, TA_CENTER), P("12", 7.5, TA_CENTER), P("日本語能力試験（JLPT）N3 合格", 7.5)],
        [P("", 7), P("", 7), P("以上", 7.5, TA_CENTER)],
    ]
    cert_t = Table(cert, colWidths=col_w)
    cert_t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "JP"),
                ("GRID", (0, 0), (-1, -1), 0.45, LINE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 1.2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1.2),
                ("BACKGROUND", (0, 0), (-1, 0), HDR),
            ]
        )
    )
    cw, ch = cert_t.wrap(usable, 40 * mm)
    cert_t.drawOn(c, ml, y - ch)
    y = y - ch - 3 * mm

    # --- 志望動機 ---
    c.setFont("JP", 8)
    c.drawString(ml, y, "志望の動機、特技、アピールポイントなど")
    y -= 1.2 * mm
    pr_text = (
        "香港にて約8年、Data / BI Engineer・システムアナリストとして、ETL、Power BI／SSAS、SQL Server、"
        "Python（FastAPI）、生成AI連携などに従事。SFCでは規制向けBI／データ基盤、Hospital Authorityでは"
        "病棟ベッド空き予測ダッシュボード等を担当。2026年7月より東京・江東区在住。"
        "英語メインのData / BI職を希望。日本語はJLPT N3（読解対応可・会話学習中）。"
        "在留資格は家族滞在（配偶者は高度専門職1号）。内定後は「技術・人文知識・国際業務」への変更が可能。"
    )
    pr_t = Table([[P(pr_text, 7, leading=9.5)]], colWidths=[usable], rowHeights=[28 * mm])
    pr_t.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, LINE),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    pw, ph = pr_t.wrap(usable, 30 * mm)
    pr_t.drawOn(c, ml, y - ph)
    y = y - ph - 3 * mm

    # --- 本人希望 ---
    c.setFont("JP", 8)
    c.drawString(ml, y, "本人希望記入欄（給料・職種・勤務時間・勤務地・その他）")
    y -= 1.2 * mm
    wish = (
        "職種：Data Engineer / BI Engineer / Analytics Engineer　｜　勤務地：東京都内（英語業務可）　｜　"
        "勤務形態：正社員<br/>"
        "その他：在留資格変更（技人国）のスポンサーシップが可能な企業を希望。ビザ申請費用は自己負担可。"
        "詳細な業務内容は職務経歴書をご参照ください。"
    )
    wish_t = Table([[P(wish, 7, leading=9.5)]], colWidths=[usable], rowHeights=[16 * mm])
    wish_t.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.45, LINE),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    ww, wh = wish_t.wrap(usable, 18 * mm)
    # Guard: if somehow overflow, still draw on this page by shrinking y
    if y - wh < 6 * mm:
        y = wh + 6 * mm
    wish_t.drawOn(c, ml, y - wh)

    c.setFont("JP", 6)
    c.setFillColor(LIGHT)
    c.drawString(
        ml,
        5 * mm,
        "※JIS履歴書の一般的な構成に準拠。生年月日・住所詳細・写真は本人記入。職務詳細は職務経歴書を添付。",
    )
    c.setFillColor(black)
    c.save()
    print("Wrote", PDF_OUT, "bottom_y", round(y - wh, 1))


def build_docx():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(1.2)
        s.bottom_margin = Cm(1.2)
        s.left_margin = Cm(1.5)
        s.right_margin = Cm(1.5)

    def font(run, size=9, bold=False):
        run.font.size = Pt(size)
        run.font.name = "Yu Gothic"
        run.bold = bold
        r = run._element.get_or_add_rPr()
        rFonts = r.get_or_add_rFonts()
        rFonts.set(qn("w:eastAsia"), "Yu Gothic")
        rFonts.set(qn("w:ascii"), "Yu Gothic")
        rFonts.set(qn("w:hAnsi"), "Yu Gothic")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("履　歴　書")
    font(r, 16, True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("2026年11月12日 現在")
    font(r, 9)

    lines = [
        "ふりがな：もく　ほいまん　／　Mok Hoi Man",
        "氏名：莫　海文（Herman Mok）",
        "生年月日：　　年　　月　　日生（満　　歳）※ご記入ください",
        "性別：（任意）",
        "現住所：〒　　－　　　 東京都江東区　　　※番地までご記入ください",
        "Tel: +81-80-9078-5842　Email: hoimanmok@gmail.com",
        "LinkedIn: linkedin.com/in/hoimanmok-herman　国籍・地域: 中国（香港）",
        "写真：別途貼付（縦40mm×横30mm、任意）",
        "",
        "【学歴】",
        "2011年9月　香港理工大学 入学",
        "2015年7月　同大学 計算機学 優等学士課程（副専攻：会計学） 卒業",
        "　　　　　以上",
        "",
        "【職歴】",
        "2016年11月　Hong Kong Life Insurance Limited 入社（プログラマー）",
        "2018年5月　一身上の都合により退社",
        "2018年6月　KBQuest Hong Kong Ltd. 入社（ソフトウェアエンジニア／BI）",
        "2020年10月　一身上の都合により退社",
        "2020年10月　Hospital Authority（病院管理局）へ Manpower Services 経由で就業",
        "　　　　　契約アナリストプログラマー（AI Systems Team）",
        "2021年10月　契約期間満了につき退職",
        "2021年10月　Securities and Futures Commission（香港証券先物委員会）へ",
        "　　　　　NewRecs Consultancy Ltd 経由で就業（システムアナリスト／BI）",
        "2026年6月　一身上の都合により退社",
        "2026年7月　日本国 東京都へ転居　現在に至る（求職活動中）",
        "　　　　　以上",
        "",
        "【免許・資格】",
        "2019年　Microsoft Data Platform & Analytics Foundational（MPN14354）",
        "2019年　Microsoft SQL Server 2016 – Mission Critical and Hybrid（MPN14355）",
        "2019年　Splunk Enterprise Certified Admin",
        "2024年12月　日本語能力試験（JLPT）N3 合格",
        "　　　　　以上",
        "",
        "【志望の動機・アピールポイント】",
        "香港にて約8年、Data / BI Engineer・システムアナリストとして、ETL、Power BI／SSAS、SQL Server、Python（FastAPI）、生成AI連携などに従事。SFCでは規制向けBI／データ基盤、Hospital Authorityでは病棟ベッド空き予測ダッシュボード等を担当。2026年7月より東京・江東区在住。英語メインのData / BI職を希望。日本語はJLPT N3。在留資格は家族滞在（配偶者は高度専門職1号）。内定後は技術・人文知識・国際業務への変更が可能。",
        "",
        "【本人希望記入欄】",
        "職種：Data Engineer / BI Engineer / Analytics Engineer｜勤務地：東京都内（英語業務可）｜勤務形態：正社員",
        "その他：在留資格変更（技人国）のスポンサーシップが可能な企業を希望。ビザ申請費用は自己負担可。",
        "詳細な業務内容は職務経歴書をご参照ください。",
    ]
    for line in lines:
        p = doc.add_paragraph()
        r = p.add_run(line)
        font(r, 9)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(0)

    doc.save(DOCX_OUT)
    print("Wrote", DOCX_OUT)


if __name__ == "__main__":
    draw_rirekisho()
    build_docx()
    for src in (PDF_OUT, DOCX_OUT):
        dst = f"/opt/cursor/artifacts/{os.path.basename(src)}"
        shutil.copy(src, dst)
        print("Copied", dst)
