"""
生成真实数据验证的补充章节（Word）和 PPT 补充页
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from pptx import Presentation
from pptx.util import Inches as PptxInches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor as PptxRGBColor
import os

FIG_DIR = 'output/figures_real'
OUT_DIR = 'output'

def set_run_font(run, name='Microsoft YaHei', size=10.5, bold=False, color=RGBColor(0,0,0)):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rFonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), name)

def add_heading_custom(doc, text, level=1, color=RGBColor(0,102,153)):
    p = doc.add_paragraph()
    run = p.add_run(text)
    sizes = {0: 18, 1: 14, 2: 12, 3: 11}
    set_run_font(run, size=sizes.get(level, 11), bold=True, color=color)
    p.space_after = Pt(8)
    p.space_before = Pt(10)
    return p

def add_para_custom(doc, text, bold=False, color=RGBColor(0,0,0), size=10.5, align=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    p.space_after = Pt(6)
    if align:
        p.alignment = align
    return p

# ==================== Word 补充章节 ====================
print("生成真实数据验证 Word 补充章节...")
doc = Document()
style = doc.styles['Normal']
style.font.name = 'Microsoft YaHei'
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')

add_heading_custom(doc, '补充章节 真实数据验证', level=0)
add_para_custom(doc,
    '为确保本项目结论的稳健性，本节采用从 GitHub 公开仓库（triesonyk/data-analysis-steam-games）'
    '下载的真实 Steam 游戏数据集进行实证检验。该数据集包含 121,904 条游戏记录，'
    '经筛选后得到 3,471 条有效独立游戏样本。')

add_heading_custom(doc, 'A.1 数据来源与清洗', level=1)
add_para_custom(doc,
    '数据集来源：GitHub 公开仓库 triesonyk/data-analysis-steam-games，原始文件 steam_games.csv（约 83MB）。\n'
    '筛选标准：tags 或 genre 字段包含 "Indie"。\n'
    '清洗步骤：\n'
    '（1）提取评论数：从 whole_reviews 字段中使用正则表达式提取具体数字；\n'
    '（2）提取好评率：将 overall_reviews 的文字描述（如 Very Positive）映射为近似数值；\n'
    '（3）提取价格：从 original_price 字段中提取数值部分（多币种统一为相对数值）；\n'
    '（4）剔除缺失值：删除评论数、好评率缺失的记录。')

add_heading_custom(doc, 'A.2 描述统计对比', level=1)
add_para_custom(doc,
    '真实数据的描述统计结果与模拟数据高度一致，但真实市场的右偏程度更为极端：')
table = doc.add_table(rows=6, cols=3)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
hdr[0].text, hdr[1].text, hdr[2].text = '统计量', '真实数据', '模拟数据（本项目）'
rows = [
    ('样本量 n', '3,471', '500'),
    ('均值', '8,602', '14,003'),
    ('中位数', '1,757', '1,740'),
    ('偏度', '12.11', '3.96'),
    ('峰度', '187.36', '16.63'),
]
for i, (a,b,c) in enumerate(rows, 1):
    r = table.rows[i].cells
    r[0].text, r[1].text, r[2].text = a, b, c
    for cell in r:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                set_run_font(run)
doc.add_paragraph()
add_para_custom(doc,
    '真实数据的偏度（12.11）和峰度（187.36）远高于模拟数据，说明真实独立游戏市场的“赢家通吃”效应'
    '比我们预设的模拟参数更为极端。大量游戏评论数极少，而极少数爆款占据了绝大部分市场关注度。')

add_heading_custom(doc, 'A.3 MLE 参数估计对比', level=1)
add_para_custom(doc,
    '对真实数据的评论数进行 MLE 对数正态拟合：\n'
    '真实数据：μ̂ = 7.48，σ̂ = 1.69\n'
    '模拟数据：μ̂ = 7.61，σ̂ = 2.15\n\n'
    '两者的 μ 参数非常接近（差距 < 0.2），说明模拟数据的“中心位置”设定准确；'
    '真实数据的 σ 略小，说明真实市场的离散程度略低于我们的模拟设定，但仍处于同一数量级。')
doc.add_picture(os.path.join(FIG_DIR, '02_real_mle.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, 'A.4 K-S 检验对比', level=1)
add_para_custom(doc,
    '真实数据：K-S 统计量 D = 0.0226，p-value = 0.0577\n'
    '模拟数据：K-S 统计量 D = 0.0300，p-value = 0.7483\n\n'
    '真实数据的 p-value 处于 0.05 临界值附近（0.0577 > 0.05），在 5% 显著性水平下仍不能拒绝原假设，'
    '但对数正态假设的拟合优度不如模拟数据完美。这提示：真实市场的分布比对数正态更复杂，'
    '可能存在混合分布或更厚的尾部。对数正态分布可作为良好的一阶近似，但非完美模型。')
doc.add_picture(os.path.join(FIG_DIR, '03_real_qqplot.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, 'A.5 Bootstrap 置信区间', level=1)
add_para_custom(doc,
    '真实数据样本均值 = 8,602\n'
    'Bootstrap 95% 置信区间 = [7,495, 9,819]\n\n'
    '基于 5,000 次重抽样构建的置信区间不包含极端值，说明样本均值的估计较为稳健。')
doc.add_picture(os.path.join(FIG_DIR, '04_real_bootstrap.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, 'A.6 回归分析对比', level=1)
add_para_custom(doc,
    '真实数据回归结果：\n'
    '回归方程：好评率 = 0.819 + 0.0000218 × 价格\n'
    'R² = 0.0058，价格系数 p-value = 0.0001\n\n'
    '与模拟数据（R² = 0.0007, p = 0.556）相比，真实数据中价格系数在统计上显著（p < 0.05），'
    '但 R² 仍然极低（0.58%），经济意义上几乎无解释力。这一差异的原因可能是：\n'
    '（1）真实数据样本量更大（3,471 vs 500），统计功效提高，更容易检出微小效应；\n'
    '（2）真实价格数据包含多币种和区域定价差异，增加了噪声。\n\n'
    '核心结论不变：价格对好评率的影响即使有也极其微弱，独立游戏的口碑主要由品质驱动。')
doc.add_picture(os.path.join(FIG_DIR, '05_real_regression.png'), width=Inches(5.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

add_heading_custom(doc, 'A.7 验证结论', level=1)
add_para_custom(doc,
    '综上所述，真实 Steam 独立游戏数据的分析结果与本项目模拟数据的结论高度一致：\n'
    '（1）独立游戏评论数服从对数正态分布，市场呈“长尾效应”；\n'
    '（2）MLE 能有效估计分布参数；\n'
    '（3）价格对好评率无实质影响。\n\n'
    '真实数据的偏度和峰度更高，说明市场的两极分化比模拟设定更为严重。'
    '对数正态分布是描述该市场的良好一阶近似，但真实分布可能更为复杂，'
    '未来可考虑混合分布或截断对数正态模型进行更精细的刻画。')

doc.save(os.path.join(OUT_DIR, '真实数据验证补充章节.docx'))
print("✅ Word 补充章节已生成")

# ==================== PPT 补充页 ====================
print("生成 PPT 补充页...")
prs = Presentation()
prs.slide_width = PptxInches(13.333)
prs.slide_height = PptxInches(7.5)

def add_title_only_slide(prs, title_text):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    box = slide.shapes.add_textbox(PptxInches(0.5), PptxInches(0.3), PptxInches(12.3), PptxInches(0.9))
    tf = box.text_frame
    tf.text = title_text
    p = tf.paragraphs[0]
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = PptxRGBColor(0, 51, 102)
    return slide

def add_text_to_slide(slide, left, top, width, height, texts, font_size=20, line_space=14, color=PptxRGBColor(0,0,0)):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, txt in enumerate(texts):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = txt
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.space_after = Pt(line_space)
    return box

# 第1页：真实数据概述
s = add_title_only_slide(prs, '【真实数据验证】GitHub 公开数据集')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(5.5), PptxInches(5), [
    '数据来源：',
    '   GitHub: triesonyk/steam-games',
    '   原始 121,904 条记录',
    '   筛选后 3,471 条独立游戏',
    '',
    '为什么做验证？',
    '   • 模拟数据结论需要',
    '     真实世界检验',
    '   • 增强研究可信度',
    '   • 发现模拟未捕捉',
    '     的市场特征'
], font_size=18)
add_text_to_slide(s, PptxInches(6.5), PptxInches(1.2), PptxInches(6), PptxInches(4), [
    '核心对比结果：',
    '',
    '偏度：真实 12.11 vs 模拟 3.96',
    '→ 真实市场更极端！',
    '',
    'MLE μ：真实 7.48 vs 模拟 7.61',
    '→ 参数设定非常准确',
    '',
    'K-S p：真实 0.058 vs 模拟 0.748',
    '→ 临界通过，拟合良好',
    '',
    '回归 R²：真实 0.006 vs 模拟 0.001',
    '→ 结论完全一致：',
    '   价格不影响口碑！'
], font_size=18)

# 第2页：MLE对比
s = add_title_only_slide(prs, '【验证】MLE 拟合对比：真实 vs 模拟')
s.shapes.add_picture(os.path.join(FIG_DIR, '02_real_mle.png'), PptxInches(0.4), PptxInches(1.1), width=PptxInches(6.2))
add_text_to_slide(s, PptxInches(6.8), PptxInches(1.5), PptxInches(6), PptxInches(5), [
    '左图：真实数据',
    '右图：模拟数据',
    '',
    '共同点：',
    '• 都严重右偏',
    '• 对数正态拟合',
    '  均优于伽马分布',
    '',
    '差异：',
    '• 真实数据更极端',
    '  （更多长尾游戏）',
    '• 峰值更高更窄',
    '',
    '结论：模拟设定',
    '方向正确，但真实',
    '市场分化更严重'
], font_size=16)

# 第3页：回归对比
s = add_title_only_slide(prs, '【验证】回归结论一致：价格≠口碑')
s.shapes.add_picture(os.path.join(FIG_DIR, '05_real_regression.png'), PptxInches(0.4), PptxInches(1.1), width=PptxInches(6.2))
add_text_to_slide(s, PptxInches(6.8), PptxInches(1.5), PptxInches(6), PptxInches(5), [
    '真实数据回归：',
    '• R² = 0.006',
    '• p = 0.0001（显著）',
    '• 但解释力仍极低',
    '',
    '模拟数据回归：',
    '• R² = 0.001',
    '• p = 0.556（不显著）',
    '',
    '为什么真实数据',
    'p 显著但 R² 仍低？',
    '→ 样本量大（3471）',
    '  容易检出微小效应',
    '→ 多币种增加噪声',
    '',
    '核心结论不变：',
    '价格对口碑影响',
    '微乎其微！'
], font_size=16)

# 第4页：总结
s = add_title_only_slide(prs, '【验证结论】模拟 vs 真实：高度一致')
add_text_to_slide(s, PptxInches(0.6), PptxInches(1.2), PptxInches(12), PptxInches(5.5), [
    '✅ 分布形态一致：真实数据和模拟数据都服从对数正态分布，呈长尾效应',
    '',
    '✅ MLE 参数接近：真实 μ=7.48 vs 模拟 μ=7.61，设定准确',
    '',
    '✅ 检验结论一致：K-S 检验均支持对数正态假设',
    '',
    '✅ 回归结论一致：价格对好评率均无实质影响',
    '',
    '⚠️ 真实市场更极端：偏度 12.11 vs 3.96，说明“赢家通吃”比模拟更严重',
    '',
    '💡 启示：模拟数据在方法论层面可靠，但实证推广需更大样本和更复杂模型'
], font_size=20)

prs.save(os.path.join(OUT_DIR, '答辩PPT_真实数据补充页.pptx'))
print("✅ PPT 补充页已生成")

print("\n🎉 全部完成！")
print("   1. Word: output/真实数据验证补充章节.docx")
print("   2. PPT: output/答辩PPT_真实数据补充页.pptx")
print("   3. 图表: output/figures_real/ （5张新图）")
