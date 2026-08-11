# -*- coding: utf-8 -*-
"""
文言文实词抽题程序 v1.1
功能：加载 Excel 资源文件，通过 GUI 生成四种输出：
  1) 输出 docx 题目（带答案）
  2) 输出 pptx 题目（带答案，答案红色+下划线）
  3) 输出 docx 知识手册（宋体/楷体/仿宋 中文字体正确）
  4) 【新增】能力测试模式（docx，全部例句乱序输出）
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import random
import os
from collections import OrderedDict

# ============ Word ============
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ============ PPT ============
from pptx import Presentation
from pptx.util import Pt as PptPt, Cm as PptCm
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn as ppt_qn


# ============================================================
# 工具函数
# ============================================================

def set_cn_font(run, font_name, size_pt):
    """同时设置西文字体和东亚中文字体，确保中文显示正确"""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    rPr = run._element.get_or_add_rPr()
    # 移除旧的 rFonts（如果有），再新建
    for old in rPr.findall(qn('w:rFonts')):
        rPr.remove(old)
    rFonts = run._element.makeelement(qn('w:rFonts'), {
        qn('w:eastAsia'): font_name,
        qn('w:ascii'): font_name,
        qn('w:hAnsi'): font_name,
    })
    rPr.insert(0, rFonts)


def set_pptx_run_font(run, font_name, size_pt, color_rgb=None, underline=False):
    """统一设置 PPTX run 的字体、字号、颜色、下划线"""
    run.font.name = font_name
    run.font.size = PptPt(size_pt)
    if color_rgb is not None:
        run.font.color.rgb = color_rgb
    if underline:
        run.font.underline = True

    # 直接操作 XML 的 rPr，写入 latin + ea 子元素
    r = run._r
    rPr = r.get_or_add_rPr()

    # 清掉旧的 latin / ea 元素，避免重复
    for tag in ['a:latin', 'a:ea', 'a:cs']:
        for old in rPr.findall(ppt_qn(tag)):
            rPr.remove(old)

    nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    latin = rPr.makeelement(ppt_qn('a:latin'), {"typeface": font_name})
    ea = rPr.makeelement(ppt_qn('a:ea'), {"typeface": font_name})
    cs = rPr.makeelement(ppt_qn('a:cs'), {"typeface": font_name})

    # 插入到 rPr 的子元素之前（保持 XML 顺序合法）
    rPr.insert(0, cs)
    rPr.insert(0, ea)
    rPr.insert(0, latin)

    # 如果有颜色，确保存在 srgbClr 子元素
    if color_rgb is not None:
        for old in rPr.findall(ppt_qn('a:solidFill')):
            rPr.remove(old)
        solid = rPr.makeelement(ppt_qn('a:solidFill'), {})
        srgb = rPr.makeelement(ppt_qn('a:srgbClr'), {"val": color_rgb.__str__()})
        solid.append(srgb)
        rPr.append(solid)

    # 下划线
    if underline:
        for old in rPr.findall(ppt_qn('a:uLnTx')):
            rPr.remove(old)
        for old in rPr.findall(ppt_qn('a:uLn')):
            rPr.remove(old)
        uLn = rPr.makeelement(ppt_qn('a:uLn'), {})
        rPr.append(uLn)


# ============================================================
# 加载 Excel
# ============================================================

def load_excel(path):
    """加载 Excel 资源文件，跳过标题行和空行，返回数据列表"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        entry = {
            "id": int(row[0]),
            "word": str(row[1]).strip() if row[1] else "",
            "meaning": str(row[2]).strip() if row[2] else "",
            "difficulty": int(row[3]) if row[3] else 1,
            "pronunciation": str(row[4]).strip() if row[4] else "",
            "examples": []
        }
        # 例句列：I(8)~W(22)，每3列一组
        for i in range(8, 23, 3):
            sentence = str(row[i]).strip() if i < len(row) and row[i] else ""
            source = str(row[i+1]).strip() if i+1 < len(row) and row[i+1] else ""
            translation = str(row[i+2]).strip() if i+2 < len(row) and row[i+2] else ""
            if sentence:
                entry["examples"].append({
                    "sentence": sentence,
                    "source": source,
                    "translation": translation
                })
        rows.append(entry)
    return rows


# ============================================================
# 难度分配抽样
# ============================================================

def stratified_sample(data_in_range, weights, count, order_mode):
    """
    按难度权重分层抽样。
    weights: 长度5，索引0对应难度1，值>=0，总和=1。
    返回: list of (sentence, word, meaning, difficulty, entry_id)
    """
    by_diff = {1: [], 2: [], 3: [], 4: [], 5: []}
    for d in data_in_range:
        diff = d["difficulty"]
        if diff not in by_diff:
            by_diff[diff] = []
        for ex in d["examples"]:
            by_diff[diff].append((ex["sentence"], d["word"], d["meaning"], diff, d["id"]))

    # 按比例计算各难度应抽数量
    raw = [count * weights[i] for i in range(5)]
    ints = [int(x) for x in raw]
    rem = count - sum(ints)
    fracs = sorted(enumerate(raw), key=lambda x: x[1] - int(x[1]), reverse=True)
    for idx, _ in fracs[:rem]:
        ints[idx] += 1

    selected = []
    for i in range(5):
        diff = i + 1
        need = ints[i]
        pool = by_diff[diff]
        if need <= 0:
            continue
        if len(pool) <= need:
            selected.extend(pool)
        else:
            selected.extend(random.sample(pool, need))

    # 不足时从其他池子补齐
    if len(selected) < count:
        have = set(id(x) for x in selected)
        extra = [x for d in range(1,6) for x in by_diff[d] if id(x) not in have]
        random.shuffle(extra)
        selected.extend(extra[:count - len(selected)])

    selected = selected[:count]

    if order_mode == "order":
        selected.sort(key=lambda x: (x[1], x[0]))
    else:
        random.shuffle(selected)

    return selected


def simple_sample(data_in_range, count, order_mode):
    """不使用难度分配时的简单随机抽样"""
    all_items = []
    for d in data_in_range:
        for ex in d["examples"]:
            all_items.append((ex["sentence"], d["word"], d["meaning"], d["difficulty"], d["id"]))
    if count > len(all_items):
        count = len(all_items)
    selected = random.sample(all_items, count)
    if order_mode == "order":
        selected.sort(key=lambda x: (x[1], x[0]))
    else:
        random.shuffle(selected)
    return selected


# ============================================================
# 难度权重校验
# ============================================================

def validate_weights(widgets_list):
    """校验5个难度输入框。返回 list of 5 floats 或 None"""
    vals = []
    has = False
    for i, w in enumerate(widgets_list):
        txt = w.get().strip()
        if txt == "":
            vals.append(0.0)
        else:
            try:
                v = float(txt)
                if v < 0 or v > 1:
                    raise ValueError
                vals.append(v)
                if v > 0:
                    has = True
            except ValueError:
                messagebox.showerror("错误", f"难度{i+1}请输入 0~1 之间的小数")
                return None
    if not has:
        messagebox.showerror("错误", "请至少输入一个难度的权重")
        return None
    s = sum(vals)
    if abs(s - 1.0) > 0.001:
        messagebox.showerror("错误", f"五个难度权重之和必须为 1，当前为 {s:.4f}")
        return None
    return vals


# ============================================================
# 设置 docx 页边距
# ============================================================

def set_doc_margins(doc, cm=1.27):
    for sec in doc.sections:
        sec.top_margin = Cm(cm)
        sec.bottom_margin = Cm(cm)
        sec.left_margin = Cm(cm)
        sec.right_margin = Cm(cm)


# ============================================================
# 模式一：输出 docx 题目
# ============================================================

def generate_docx_questions(data_rows, sub_cfg):
    try:
        start_id = int(sub_cfg["start"].get())
        end_id = int(sub_cfg["end"].get())
        num_papers = int(sub_cfg["papers"].get())
        per_paper = int(sub_cfg["per_paper"].get())
    except ValueError:
        messagebox.showerror("错误", "请填写有效的整数")
        return

    use_diff = sub_cfg["diff_var"].get()
    weights = None
    if use_diff:
        weights = validate_weights(sub_cfg["diff_entries"])
        if weights is None:
            return

    order_mode = sub_cfg["order_var"].get()

    data_range = [r for r in data_rows if start_id <= r["id"] <= end_id]
    if not data_range:
        messagebox.showerror("错误", "所选序号范围内没有数据")
        return

    doc = Document()
    set_doc_margins(doc)

    for paper_idx in range(1, num_papers + 1):
        if use_diff:
            selected = stratified_sample(data_range, weights, per_paper, order_mode)
        else:
            selected = simple_sample(data_range, per_paper, order_mode)

        # 标题
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_title.add_run(f"第{paper_idx}天实词义项默写")
        set_cn_font(run, "黑体", 16)  # 三号≈16pt

        # 题目
        for idx, (sent, word, meaning, *_) in enumerate(selected, 1):
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            text = f"{idx}. {sent}    {word}:{'_' * 16}"
            run = p.add_run(text)
            set_cn_font(run, "宋体", 10.5)  # 五号=10.5pt

        doc.add_page_break()

        # 答案标题
        p_ans_title = doc.add_paragraph()
        p_ans_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_ans_title.add_run(f"第{paper_idx}天实词义项默写答案")
        set_cn_font(run, "黑体", 16)

        # 答案
        for idx, (sent, word, meaning, *_) in enumerate(selected, 1):
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.0
            text = f"{idx}. {sent}    {word}:{meaning}"
            run = p.add_run(text)
            set_cn_font(run, "宋体", 10.5)

        if paper_idx < num_papers:
            doc.add_page_break()

    save_path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word 文档", "*.docx")],
        title="保存 docx 题目"
    )
    if save_path:
        doc.save(save_path)
        messagebox.showinfo("完成", f"已保存：{save_path}")


# ============================================================
# 模式二：输出 pptx 题目
# ============================================================

def generate_pptx_questions(data_rows, sub_cfg):
    try:
        start_id = int(sub_cfg["start"].get())
        end_id = int(sub_cfg["end"].get())
        num_papers = int(sub_cfg["papers"].get())
        per_paper = int(sub_cfg["per_paper"].get())
    except ValueError:
        messagebox.showerror("错误", "请填写有效的整数")
        return

    use_diff = sub_cfg["diff_var"].get()
    weights = None
    if use_diff:
        weights = validate_weights(sub_cfg["diff_entries"])
        if weights is None:
            return

    order_mode = sub_cfg["order_var"].get()

    data_range = [r for r in data_rows if start_id <= r["id"] <= end_id]
    if not data_range:
        messagebox.showerror("错误", "所选序号范围内没有数据")
        return

    saved_files = []
    for paper_idx in range(1, num_papers + 1):
        if use_diff:
            selected = stratified_sample(data_range, weights, per_paper, order_mode)
        else:
            selected = simple_sample(data_range, per_paper, order_mode)

        prs = Presentation()
        prs.slide_width = Cm(31.5)
        prs.slide_height = Cm(13.5)

        # ---- 题目页（每页10道） ----
        q_chunks = [selected[i:i+10] for i in range(0, len(selected), 10)]
        for chunk in q_chunks:
            slide_q = prs.slides.add_slide(prs.slide_layouts[6])
            tb_q = slide_q.shapes.add_textbox(Cm(0), Cm(2), Cm(31.5), Cm(11.5))
            tf_q = tb_q.text_frame
            tf_q.word_wrap = True

            for idx, (sent, word, meaning, *_) in enumerate(chunk, 1):
                line = f"{idx}. {sent}    {word}:{'_' * 4}"
                p = tf_q.paragraphs[0] if idx == 1 else tf_q.add_paragraph()
                run = p.add_run()
                run.text = line
                set_pptx_run_font(run, "华文中宋", 24)

        # ---- 答案页（每页10条） ----
        ans_chunks = [selected[i:i+10] for i in range(0, len(selected), 10)]
        for chunk in ans_chunks:
            slide_a = prs.slides.add_slide(prs.slide_layouts[6])
            tb_a = slide_a.shapes.add_textbox(Cm(0), Cm(2), Cm(31.5), Cm(11.5))
            tf_a = tb_a.text_frame
            tf_a.word_wrap = True

            for j, (sent, word, meaning, *_) in enumerate(chunk, 1):
                p = tf_a.paragraphs[0] if j == 1 else tf_a.add_paragraph()

                # 例句部分（黑色）
                run1 = p.add_run()
                run1.text = f"{j}. {sent}    {word}:"
                set_pptx_run_font(run1, "华文中宋", 24, color_rgb=None, underline=False)

                # 义项部分（红色+下划线）
                run2 = p.add_run()
                run2.text = meaning
                set_pptx_run_font(run2, "华文中宋", 24,
                                  color_rgb=RGBColor(255, 0, 0), underline=True)

        # 保存单篇
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PPT 演示文稿", "*.pptx")],
            title=f"保存第 {paper_idx} 篇 pptx 题目",
            initialfile=f"第{paper_idx}天实词义项默写.pptx"
        )
        if save_path:
            prs.save(save_path)
            saved_files.append(save_path)

    if saved_files:
        messagebox.showinfo("完成", f"已生成 {len(saved_files)} 个 pptx 文件")


# ============================================================
# 模式三：输出 docx 知识手册
# ============================================================

def generate_docx_handbook(data_rows, sub_cfg):
    try:
        start_id = int(sub_cfg["start"].get())
        end_id = int(sub_cfg["end"].get())
    except ValueError:
        messagebox.showerror("错误", "请填写有效的整数")
        return

    data_range = [r for r in data_rows if start_id <= r["id"] <= end_id]
    if not data_range:
        messagebox.showerror("错误", "所选范围内没有数据")
        return

    # 按实词分组
    word_groups = OrderedDict()
    for d in data_range:
        w = d["word"]
        if w not in word_groups:
            word_groups[w] = []
        word_groups[w].append(d)

    doc = Document()
    set_doc_margins(doc)

    global_seq = 0
    for word, entries in word_groups.items():
        global_seq += 1

        # 按拼音分组
        pron_groups = OrderedDict()
        for e in entries:
            pron = e["pronunciation"] if e["pronunciation"] else ""
            if pron not in pron_groups:
                pron_groups[pron] = []
            pron_groups[pron].append(e)

        # 序号.实词 —— 宋体小四(12pt)
        p_word = doc.add_paragraph()
        p_word.paragraph_format.line_spacing = 1.0
        run = p_word.add_run(f"{global_seq}.{word}")
        set_cn_font(run, "宋体", 12)

        meaning_seq = 0
        for pron, pron_entries in pron_groups.items():
            # 拼音行 —— 宋体小四
            if pron:
                p_pron = doc.add_paragraph()
                p_pron.paragraph_format.line_spacing = 1.0
                run = p_pron.add_run(pron)
                set_cn_font(run, "宋体", 12)

            for e in pron_entries:
                meaning_seq += 1
                # 义项 —— 宋体小四
                p_mean = doc.add_paragraph()
                p_mean.paragraph_format.line_spacing = 1.0
                run = p_mean.add_run(f"（{meaning_seq}）{e['meaning']}")
                set_cn_font(run, "宋体", 12)

                # 例句 —— 楷体小四
                for ex in e["examples"]:
                    if ex["sentence"]:
                        p_ex = doc.add_paragraph()
                        p_ex.paragraph_format.line_spacing = 1.0
                        line = ex["sentence"]
                        if ex["source"]:
                            line += f"——{ex['source']}"
                        run = p_ex.add_run(line)
                        set_cn_font(run, "楷体", 12)

                    # 参考翻译 —— 仿宋小四
                    if ex["translation"]:
                        p_trans = doc.add_paragraph()
                        p_trans.paragraph_format.line_spacing = 1.0
                        run = p_trans.add_run(f"参考翻译：{ex['translation']}")
                        set_cn_font(run, "仿宋", 12)

        # 实词之间空一行
        doc.add_paragraph()

    save_path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word 文档", "*.docx")],
        title="保存 docx 知识手册"
    )
    if save_path:
        doc.save(save_path)
        messagebox.showinfo("完成", f"已保存：{save_path}")


# ============================================================
# 模式四（新增）：能力测试模式
# ============================================================

def generate_ability_test(data_rows, sub_cfg):
    """
    能力测试模式：
    1. 按用户指定序号范围确定数据
    2. 不随机抽取，将指定范围内所有例句打乱顺序全部输出
    3. 输出格式：[序号].[例句]（[A列总序数]）[4字符空格][实词或词组]：[4字符下划线]
    4. 无标题，宋体五号，单倍行距，页边距1.27cm
    5. 自动增加序号
    """
    try:
        start_id = int(sub_cfg["start"].get())
        end_id = int(sub_cfg["end"].get())
    except ValueError:
        messagebox.showerror("错误", "请填写有效的整数")
        return

    # 确定数据范围
    data_range = [r for r in data_rows if start_id <= r["id"] <= end_id]
    if not data_range:
        messagebox.showerror("错误", "所选序号范围内没有数据")
        return

    # 收集所有例句，附带 A 列总序数
    all_items = []
    for d in data_range:
        for ex in d["examples"]:
            all_items.append({
                "sentence": ex["sentence"],
                "word": d["word"],
                "row_id": d["id"]
            })

    if not all_items:
        messagebox.showerror("错误", "所选范围内没有例句数据")
        return

    # 打乱顺序（不随机抽取，全部输出但顺序打乱）
    random.shuffle(all_items)

    # 生成 docx
    doc = Document()
    set_doc_margins(doc)

    for idx, item in enumerate(all_items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.0
        text = f"{idx}.{item['sentence']}（{item['row_id']}）    {item['word']}:{'_' * 16}"
        run = p.add_run(text)
        set_cn_font(run, "宋体", 10.5)  # 五号 = 10.5pt

    save_path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word 文档", "*.docx")],
        title="保存能力测试文档"
    )
    if save_path:
        doc.save(save_path)
        messagebox.showinfo("完成", f"已保存：{save_path}，共 {len(all_items)} 道题目")


# ============================================================
# GUI 主界面
# ============================================================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("文言文实词抽题系统 v1.1")
        self.root.geometry("900x720")

        self.data_rows = []
        self.file_path = ""

        # ---- 文件选择 ----
        frm_top = tk.Frame(root)
        frm_top.pack(fill="x", padx=10, pady=8)

        tk.Button(frm_top, text="📂 选择 Excel 资源文件", command=self.select_file,
                  font=("宋体", 11, "bold"), bg="#607D8B", fg="white",
                  padx=12, pady=4).pack(side="left")

        self.stats_var = tk.StringVar(value="实词数：0    义项数：0    例句数：0")
        tk.Label(frm_top, textvariable=self.stats_var, font=("宋体", 10),
                 fg="gray30").pack(side="right")

        # ---- 模式选择 ----
        self.mode_var = tk.StringVar(value="")
        self.mode_var.trace_add("write", self.on_mode_change)

        frm_mode = tk.LabelFrame(root, text="选择模式（四选一）", font=("宋体", 10, "bold"))
        frm_mode.pack(fill="x", padx=10, pady=5)

        tk.Radiobutton(frm_mode, text="📄 输出 docx 题目", variable=self.mode_var,
                       value="docx_q", font=("宋体", 10)).grid(row=0, column=0, sticky="w", padx=10)
        tk.Radiobutton(frm_mode, text="📽 输出 pptx 题目", variable=self.mode_var,
                       value="pptx_q", font=("宋体", 10)).grid(row=0, column=1, sticky="w", padx=10)
        tk.Radiobutton(frm_mode, text="📚 输出 docx 知识手册", variable=self.mode_var,
                       value="docx_h", font=("宋体", 10)).grid(row=1, column=0, sticky="w", padx=10, pady=4)
        # 【新增】能力测试模式单选按钮
        tk.Radiobutton(frm_mode, text="🧠 能力测试模式", variable=self.mode_var,
                       value="ability", font=("宋体", 10)).grid(row=1, column=1, sticky="w", padx=10, pady=4)

        # ---- 子配置容器 ----
        self.frm_docx_q = tk.LabelFrame(root, text="docx 题目配置", font=("宋体", 10))
        self.frm_pptx_q = tk.LabelFrame(root, text="pptx 题目配置", font=("宋体", 10))
        self.frm_docx_h = tk.LabelFrame(root, text="知识手册配置", font=("宋体", 10))
        # 【新增】能力测试子配置容器
        self.frm_ability = tk.LabelFrame(root, text="能力测试配置", font=("宋体", 10))

        self.cfg_docx_q = self._build_question_subconfig(self.frm_docx_q, btn_text="▶ 生成 docx 题目", btn_color="#4CAF50")
        self.cfg_pptx_q = self._build_question_subconfig(self.frm_pptx_q, btn_text="▶ 生成 pptx 题目", btn_color="#2196F3")
        self.cfg_docx_h = self._build_handbook_subconfig(self.frm_docx_h)
        # 【新增】构建能力测试子配置
        self.cfg_ability = self._build_ability_subconfig(self.frm_ability)

    # ---------- 文件选择 ----------
    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if not path:
            return
        try:
            self.data_rows = load_excel(path)
            self.file_path = path
            self._update_stats()
            messagebox.showinfo("成功",
                f"已加载 {len(self.data_rows)} 条义项数据\n来自：{os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败：{e}")

    def _update_stats(self):
        words = set()
        examples = 0
        for r in self.data_rows:
            words.add(r["word"])
            examples += len(r["examples"])
        self.stats_var.set(f"实词数：{len(words)}    义项数：{len(self.data_rows)}    例句数：{examples}")

    # ---------- 模式切换 ----------
    def on_mode_change(self, *args):
        for f in [self.frm_docx_q, self.frm_pptx_q, self.frm_docx_h, self.frm_ability]:
            f.pack_forget()
        mode = self.mode_var.get()
        if mode == "docx_q":
            self.frm_docx_q.pack(fill="both", expand=True, padx=10, pady=5)
        elif mode == "pptx_q":
            self.frm_pptx_q.pack(fill="both", expand=True, padx=10, pady=5)
        elif mode == "docx_h":
            self.frm_docx_h.pack(fill="both", expand=True, padx=10, pady=5)
        # 【新增】能力测试模式显示
        elif mode == "ability":
            self.frm_ability.pack(fill="both", expand=True, padx=10, pady=5)

    # ---------- 构建题目类子配置（docx / pptx 共用） ----------
    def _build_question_subconfig(self, parent, btn_text, btn_color):
        cfg = {}

        # 序号范围
        tk.Label(parent, text="请输入序号范围：", font=("宋体", 10)).grid(row=0, column=0, sticky="w", pady=2)
        e_start = tk.Entry(parent, width=8, font=("宋体", 10)); e_start.insert(0, "1")
        e_start.grid(row=0, column=1, padx=2)
        tk.Label(parent, text="-", font=("宋体", 10)).grid(row=0, column=2)
        e_end = tk.Entry(parent, width=8, font=("宋体", 10))
        e_end.grid(row=0, column=3, padx=2)
        cfg["start"] = e_start; cfg["end"] = e_end

        # 生成篇数
        tk.Label(parent, text="生成篇数：", font=("宋体", 10)).grid(row=1, column=0, sticky="w", pady=2)
        e_papers = tk.Entry(parent, width=8, font=("宋体", 10)); e_papers.insert(0, "1")
        e_papers.grid(row=1, column=1, padx=2)
        cfg["papers"] = e_papers

        # 每篇题目数
        tk.Label(parent, text="每篇题目数：", font=("宋体", 10)).grid(row=2, column=0, sticky="w", pady=2)
        e_per = tk.Entry(parent, width=8, font=("宋体", 10))
        e_per.grid(row=2, column=1, padx=2)
        cfg["per_paper"] = e_per

        # 难度分配模式
        diff_var = tk.BooleanVar(value=False)
        tk.Checkbutton(parent, text="难度分配模式", variable=diff_var, font=("宋体", 10),
                       command=lambda: self._toggle_diff(cfg)).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=2)
        cfg["diff_var"] = diff_var

        diff_frame = tk.Frame(parent)
        diff_frame.grid(row=4, column=0, columnspan=6, sticky="w", padx=20)
        cfg["diff_frame"] = diff_frame
        diff_entries = []
        for i in range(5):
            tk.Label(diff_frame, text=f"难度{i+1}", font=("宋体", 9)).grid(row=0, column=i*2, padx=(0,2))
            e = tk.Entry(diff_frame, width=6, font=("宋体", 9), state="disabled")
            e.grid(row=0, column=i*2+1, padx=(0,10))
            diff_entries.append(e)
        cfg["diff_entries"] = diff_entries

        # 顺序 / 乱序
        order_var = tk.StringVar(value="order")
        tk.Radiobutton(parent, text="顺序模式", variable=order_var, value="order",
                       font=("宋体", 10)).grid(row=5, column=0, sticky="w")
        tk.Radiobutton(parent, text="乱序模式", variable=order_var, value="random",
                       font=("宋体", 10)).grid(row=5, column=1, sticky="w")
        cfg["order_var"] = order_var

        # 生成按钮
        tk.Button(parent, text=btn_text,
                  command=lambda c=cfg, b=btn_color: self._dispatch(c, b),
                  bg=btn_color, fg="white", font=("宋体", 10, "bold"),
                  padx=10, pady=4).grid(row=100, column=0, columnspan=6, pady=10, sticky="w")

        return cfg

    def _toggle_diff(self, cfg):
        state = "normal" if cfg["diff_var"].get() else "disabled"
        for e in cfg["diff_entries"]:
            e.config(state=state)

    def _dispatch(self, cfg, btn_color):
        # 根据调用者决定调用哪个生成函数
        if cfg is self.cfg_docx_q:
            generate_docx_questions(self.data_rows, cfg)
        elif cfg is self.cfg_pptx_q:
            generate_pptx_questions(self.data_rows, cfg)
        elif cfg is self.cfg_docx_h:
            generate_docx_handbook(self.data_rows, cfg)
        # 【新增】能力测试模式分发
        elif cfg is self.cfg_ability:
            generate_ability_test(self.data_rows, cfg)

    # ---------- 构建知识手册子配置 ----------
    def _build_handbook_subconfig(self, parent):
        cfg = {}
        tk.Label(parent, text="请输入义项范围：", font=("宋体", 10)).grid(row=0, column=0, sticky="w", pady=2)
        e_start = tk.Entry(parent, width=8, font=("宋体", 10)); e_start.insert(0, "1")
        e_start.grid(row=0, column=1, padx=2)
        tk.Label(parent, text="-", font=("宋体", 10)).grid(row=0, column=2)
        e_end = tk.Entry(parent, width=8, font=("宋体", 10))
        e_end.grid(row=0, column=3, padx=2)
        cfg["start"] = e_start; cfg["end"] = e_end

        tk.Button(parent, text="▶ 生成 docx 知识手册",
                  command=lambda: generate_docx_handbook(self.data_rows, cfg),
                  bg="#FF9800", fg="white", font=("宋体", 10, "bold"),
                  padx=10, pady=4).grid(row=10, column=0, columnspan=4, pady=10, sticky="w")
        return cfg

    # ---------- 【新增】构建能力测试子配置 ----------
    def _build_ability_subconfig(self, parent):
        cfg = {}
        # 序号范围（仅需起始和结束，无难度、无篇数、无顺序选择）
        tk.Label(parent, text="请输入序号范围：", font=("宋体", 10)).grid(row=0, column=0, sticky="w", pady=2)
        e_start = tk.Entry(parent, width=8, font=("宋体", 10)); e_start.insert(0, "1")
        e_start.grid(row=0, column=1, padx=2)
        tk.Label(parent, text="-", font=("宋体", 10)).grid(row=0, column=2)
        e_end = tk.Entry(parent, width=8, font=("宋体", 10))
        e_end.grid(row=0, column=3, padx=2)
        cfg["start"] = e_start; cfg["end"] = e_end

        # 提示说明
        tk.Label(parent, text="说明：将指定范围内所有例句打乱顺序全部输出，不随机抽取。",
                 font=("宋体", 9), fg="gray40").grid(row=1, column=0, columnspan=6, sticky="w", pady=(6, 2))

        # 生成按钮
        tk.Button(parent, text="▶ 生成能力测试文档",
                  command=lambda: generate_ability_test(self.data_rows, cfg),
                  bg="#9C27B0", fg="white", font=("宋体", 10, "bold"),
                  padx=10, pady=4).grid(row=10, column=0, columnspan=4, pady=10, sticky="w")
        return cfg


# ============================================================
# 入口
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
