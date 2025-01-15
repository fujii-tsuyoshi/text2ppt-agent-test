from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE

# テンプレートのプレゼンテーションを読み込む
prs = Presentation('/workspace/input/template.pptx')

# スライド 1: タイトルスライド
slide_layout = prs.slide_layouts[2]
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[10]
subtitle = slide.placeholders[11]
title.text = "大谷翔平: 二刀流の軌跡と未来"
subtitle.text = "野球界の新たな伝説"

# タイトルとサブタイトルを囲む枠を追加
left = Inches(0.5)
top = Inches(1.5)
width = Inches(8.5)
height = Inches(2.0)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)

# スライド 2: はじめに
slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[0]
content = slide.placeholders[1]
title.text = "はじめに"
content.text = "大谷翔平は、現代野球における「二刀流」の象徴であり、彼の活躍は野球界に新たな基準をもたらしました。"

# 投手と打者のアイコンを追加
left = Inches(1)
top = Inches(2)
width = Inches(1)
height = Inches(1)
shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)

# スライド 3: 初期の軌跡
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[0]
content = slide.placeholders[1]
title.text = "初期の軌跡"
content.text = "岩手県奥州市での少年時代から、160 km/hの速球を投げる高校生へ。"

# 時系列を示す矢印を追加
left = Inches(1)
top = Inches(2)
width = Inches(5)
height = Inches(0.5)
shape = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, left, top, width, height)

# スライド 4: 日本プロ野球での活躍
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[0]
content = slide.placeholders[1]
title.text = "日本プロ野球での活躍"
content.text = "北海道日本ハムファイターズでの「二刀流」デビューと成功。"

# 日本ハムのロゴを追加
left = Inches(1)
top = Inches(2)
width = Inches(1)
height = Inches(1)
shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)

# スライド 5: メジャーリーグへの挑戦
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[0]
content = slide.placeholders[1]
title.text = "メジャーリーグへの挑戦"
content.text = "ロサンゼルス・エンゼルスでの挑戦と、MLBでの歴史的な記録達成。"

# エンゼルスのロゴを追加
left = Inches(1)
top = Inches(2)
width = Inches(1)
height = Inches(1)
shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)

# スライド 6: 2023年の偉業
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[0]
content = slide.placeholders[1]
title.text = "2023年の偉業"
content.text = "史上初の50本塁打50盗塁達成と、ドジャース移籍による新たな挑戦。"

# ドジャースのロゴを追加
left = Inches(1)
top = Inches(2)
width = Inches(1)
height = Inches(1)
shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)

# スライド 7: 世界的な影響力
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[0]
content = slide.placeholders[1]
title.text = "世界的な影響力"
content.text = "タイム誌「世界で最も影響力のある100人」に選出されるなど、スポーツを超えた影響力。"

# タイム誌のロゴを追加
left = Inches(1)
top = Inches(2)
width = Inches(1)
height = Inches(1)
shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)

# スライド 8: 経済的成功
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[0]
content = slide.placeholders[1]
title.text = "経済的成功"
content.text = "スポーツ史上最高額の契約と、フォーブスのスポーツ選手長者番付での上位ランクイン。"

# フォーブスのロゴを追加
left = Inches(1)
top = Inches(2)
width = Inches(1)
height = Inches(1)
shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)

# スライド 9: 未来への展望
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[0]
content = slide.placeholders[1]
title.text = "未来への展望"
content.text = "大谷翔平の未来は、さらなる記録更新と野球界の発展に貢献することが期待されます。"

# 未来を示す矢印を追加
left = Inches(1)
top = Inches(2)
width = Inches(5)
height = Inches(0.5)
shape = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, left, top, width, height)

# スライド 10: 結論
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[0]
content = slide.placeholders[1]
title.text = "結論"
content.text = "大谷翔平は、野球界における新たな伝説であり、彼の挑戦は続きます。"

# サインボールのイラストを追加
left = Inches(1)
top = Inches(2)
width = Inches(1)
height = Inches(1)
shape = slide.shapes.add_shape(MSO_SHAPE.BALLOON, left, top, width, height)

# スライド 11: 質疑応答
slide = prs.slides.add_slide(slide_layout)
title = slide.placeholders[0]
content = slide.placeholders[1]
title.text = "質疑応答"
content.text = "ご質問があればどうぞ。"

# 質問マークのアイコンを追加
left = Inches(1)
top = Inches(2)
width = Inches(1)
height = Inches(1)
shape = slide.shapes.add_shape(MSO_SHAPE.OVAL_CALLOUT, left, top, width, height)

# プレゼンテーションを保存
prs.save('/workspace/output/otani_presentation.pptx')
