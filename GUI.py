# -*- coding: utf-8 -*-
"""
CCTV-17 频道收视日报 PPT 生成器 — GUI 版

功能：
- 拖拽 / 浏览选择 Excel 数据文件
- 自动检测 demo 模板（output_ppt/ 下含 "demo" 的 .pptx）
- 根据输入文件名自动推导输出文件名  (如 20260128 → CCTV-17频道收视日报 0128.pptx)
- 记忆上次使用的模板路径和输出目录
- 生成过程日志实时显示
"""

import json
import os
import re
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path

# ── 路径处理 ──
# 打包后 _MEIPASS 为临时解压目录；开发时为脚本所在目录
if getattr(sys, 'frozen', False):
    _BASE_DIR = Path(sys.executable).parent
else:
    _BASE_DIR = Path(__file__).resolve().parent

_CONFIG_FILE = _BASE_DIR / '.gui_config.json'
_DEFAULT_OUTPUT_DIR = _BASE_DIR / 'output_ppt'
_DEFAULT_INPUT_DIR = _BASE_DIR / 'input_data'


# ═══════════════════════════════════════════
# 配置记忆
# ═══════════════════════════════════════════

def _load_config() -> dict:
    """
    读取配置文件，并在启动时自动校验路径是否仍有效。
    若某项路径失效（换电脑/换用户），自动用本机默认值替换并回写。
    """
    try:
        cfg = json.loads(_CONFIG_FILE.read_text('utf-8'))
    except Exception:
        cfg = {}

    changed = False

    # ── 模板路径：失效则重新检测 ──
    tpl = cfg.get('template', '')
    if not tpl or not Path(tpl).is_file():
        new_tpl = _find_demo_template()
        if new_tpl:
            cfg['template'] = new_tpl
            changed = True
        elif 'template' in cfg:
            del cfg['template']
            changed = True

    # ── 输入目录：失效则用 exe 旁的 input_data/，再不行就用 exe 目录 ──
    in_dir = cfg.get('input_dir', '')
    if not in_dir or not Path(in_dir).is_dir():
        fallback = _DEFAULT_INPUT_DIR if _DEFAULT_INPUT_DIR.exists() else _BASE_DIR
        cfg['input_dir'] = str(fallback)
        changed = True

    # ── 输出目录：失效则用 exe 目录 ──
    out_dir = cfg.get('output_dir', '')
    if not out_dir or not Path(out_dir).is_dir():
        cfg['output_dir'] = str(_BASE_DIR)
        changed = True

    if changed:
        _save_config(cfg)

    return cfg


def _save_config(cfg: dict):
    try:
        _CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), 'utf-8')
    except Exception:
        pass


# ═══════════════════════════════════════════
# 文件名推导
# ═══════════════════════════════════════════

def _extract_date_from_filename(filepath: str) -> str | None:
    """从文件名中提取 YYYYMMDD 日期字符串"""
    basename = os.path.basename(filepath)
    m = re.search(r'(\d{8})', basename)
    return m.group(1) if m else None


def _detect_suffix(filepath: str) -> str:
    """检测文件名中是否含（简版）等后缀"""
    basename = os.path.basename(filepath)
    if '简版' in basename:
        return '（简版）'
    return ''


def _make_output_name(excel_path: str, output_dir: str) -> str:
    """根据输入 Excel 文件名推导输出 PPT 名称"""
    date_str = _extract_date_from_filename(excel_path)
    suffix = _detect_suffix(excel_path)
    if date_str:
        mmdd = f'{date_str[4:6]}{date_str[6:8]}'
        name = f'CCTV-17频道收视日报 {mmdd}{suffix}.pptx'
    else:
        name = 'CCTV-17频道收视日报.pptx'
    return os.path.join(output_dir, name)


def _find_demo_template() -> str | None:
    """在多个候选目录下自动查找含 'demo' 的 pptx 文件"""
    # 候选搜索目录：exe所在目录及其父目录的 output_ppt/，以及当前工作目录
    candidates = [
        _BASE_DIR / 'output_ppt',
        _BASE_DIR.parent / 'output_ppt',
        Path.cwd() / 'output_ppt',
        _BASE_DIR,
        _BASE_DIR.parent,
    ]
    for d in candidates:
        if not d.exists():
            continue
        for f in sorted(d.iterdir()):
            if f.suffix.lower() == '.pptx' and 'demo' in f.name.lower() and not f.name.startswith('~$'):
                return str(f)
    return None


# ═══════════════════════════════════════════
# 拖拽支持 (tkinterdnd2 可选)
# ═══════════════════════════════════════════

_HAS_DND = False
try:
    import tkinterdnd2
    _HAS_DND = True
except ImportError:
    pass


# ═══════════════════════════════════════════
# GUI 主窗口
# ═══════════════════════════════════════════

class App:
    def __init__(self):
        self.cfg = _load_config()

        # 如果有 tkinterdnd2 就用 TkinterDnD.Tk，否则普通 Tk
        if _HAS_DND:
            self.root = tkinterdnd2.TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title('CCTV-17 收视日报 PPT 生成器')
        self.root.geometry('720x560')
        self.root.resizable(True, True)

        self._build_ui()
        self._running = False

    # ─────────────────────────────
    # UI 构建
    # ─────────────────────────────
    def _build_ui(self):
        pad = dict(padx=8, pady=4)
        root = self.root

        # ── 输入 Excel ──
        frame_in = tk.LabelFrame(root, text=' 输入 Excel 文件 ', font=('微软雅黑', 10, 'bold'))
        frame_in.pack(fill='x', **pad)

        self.var_excel = tk.StringVar()
        entry_excel = tk.Entry(frame_in, textvariable=self.var_excel, font=('微软雅黑', 9))
        entry_excel.pack(side='left', fill='x', expand=True, padx=4, pady=6)
        tk.Button(frame_in, text='浏览...', command=self._browse_excel,
                  font=('微软雅黑', 9)).pack(side='right', padx=4, pady=6)

        # 拖拽提示
        if _HAS_DND:
            entry_excel.drop_target_register(tkinterdnd2.DND_FILES)
            entry_excel.dnd_bind('<<Drop>>', self._on_drop_excel)
            hint = '支持拖拽 Excel 文件到此处'
        else:
            hint = '点击"浏览"选择 Excel 文件（安装 tkinterdnd2 可支持拖拽）'
        tk.Label(frame_in, text=hint, fg='gray',
                 font=('微软雅黑', 8)).pack(side='bottom', anchor='w', padx=4)

        # ── 模板 ──
        frame_tpl = tk.LabelFrame(root, text=' Demo 模板 ', font=('微软雅黑', 10, 'bold'))
        frame_tpl.pack(fill='x', **pad)

        auto_tpl = self.cfg.get('template', '') or _find_demo_template() or ''
        self.var_template = tk.StringVar(value=auto_tpl)
        entry_tpl = tk.Entry(frame_tpl, textvariable=self.var_template, font=('微软雅黑', 9))
        entry_tpl.pack(side='left', fill='x', expand=True, padx=4, pady=6)
        tk.Button(frame_tpl, text='浏览...', command=self._browse_template,
                  font=('微软雅黑', 9)).pack(side='right', padx=4, pady=6)
        if not auto_tpl:
            tk.Label(frame_tpl, text='⚠ 未找到 demo 模板，请手动浏览选择',
                     fg='#cc6600', font=('微软雅黑', 8)).pack(side='bottom', anchor='w', padx=4)

        # ── 输出 ──
        frame_out = tk.LabelFrame(root, text=' 输出 PPT ', font=('微软雅黑', 10, 'bold'))
        frame_out.pack(fill='x', **pad)

        self.var_output = tk.StringVar()
        tk.Entry(frame_out, textvariable=self.var_output,
                 font=('微软雅黑', 9)).pack(side='left', fill='x', expand=True, padx=4, pady=6)
        tk.Button(frame_out, text='浏览...', command=self._browse_output,
                  font=('微软雅黑', 9)).pack(side='right', padx=4, pady=6)

        # 当 excel 路径变化时自动推导输出名
        self.var_excel.trace_add('write', self._on_excel_changed)

        # ── 按钮 ──
        frame_btn = tk.Frame(root)
        frame_btn.pack(fill='x', **pad)

        self.btn_run = tk.Button(frame_btn, text='▶  开始生成', font=('微软雅黑', 12, 'bold'),
                                 bg='#4A7C31', fg='white', activebackground='#3a6228',
                                 activeforeground='white', height=2,
                                 command=self._run)
        self.btn_run.pack(fill='x', padx=4, pady=4)

        # ── 日志 ──
        frame_log = tk.LabelFrame(root, text=' 生成日志 ', font=('微软雅黑', 10, 'bold'))
        frame_log.pack(fill='both', expand=True, **pad)

        self.log = scrolledtext.ScrolledText(frame_log, font=('Consolas', 9), state='disabled',
                                             wrap='word', bg='#1e1e1e', fg='#d4d4d4',
                                             insertbackground='white')
        self.log.pack(fill='both', expand=True, padx=4, pady=4)

    # ─────────────────────────────
    # 浏览 / 拖拽
    # ─────────────────────────────
    def _browse_excel(self):
        init_dir = self.cfg.get('input_dir', str(_DEFAULT_INPUT_DIR))
        fp = filedialog.askopenfilename(
            title='选择 Excel 数据文件',
            initialdir=init_dir,
            filetypes=[('Excel 文件', '*.xlsx *.xls'), ('所有文件', '*.*')]
        )
        if fp:
            self.var_excel.set(fp)

    def _on_drop_excel(self, event):
        fp = event.data.strip('{}')  # tkinterdnd2 用大括号包裹带空格的路径
        if fp.lower().endswith(('.xlsx', '.xls')):
            self.var_excel.set(fp)

    def _browse_template(self):
        init_dir = str(_DEFAULT_OUTPUT_DIR) if _DEFAULT_OUTPUT_DIR.exists() else ''
        fp = filedialog.askopenfilename(
            title='选择 Demo 模板 PPTX',
            initialdir=init_dir,
            filetypes=[('PowerPoint 模板', '*.pptx'), ('所有文件', '*.*')]
        )
        if fp:
            self.var_template.set(fp)

    def _browse_output(self):
        init_dir = self.cfg.get('output_dir', str(_DEFAULT_OUTPUT_DIR))
        fp = filedialog.asksaveasfilename(
            title='输出 PPT 保存位置',
            initialdir=init_dir,
            defaultextension='.pptx',
            filetypes=[('PowerPoint 文件', '*.pptx')]
        )
        if fp:
            self.var_output.set(fp)

    def _on_excel_changed(self, *_args):
        excel = self.var_excel.get().strip()
        if not excel:
            return
        output_dir = self.cfg.get('output_dir', str(_DEFAULT_OUTPUT_DIR))
        self.var_output.set(_make_output_name(excel, output_dir))

    # ─────────────────────────────
    # 日志
    # ─────────────────────────────
    def _log(self, text: str):
        def _append():
            self.log.config(state='normal')
            self.log.insert('end', text + '\n')
            self.log.see('end')
            self.log.config(state='disabled')
        self.root.after(0, _append)

    # ─────────────────────────────
    # 生成
    # ─────────────────────────────
    def _run(self):
        if self._running:
            return

        excel = self.var_excel.get().strip()
        template = self.var_template.get().strip()
        output = self.var_output.get().strip()

        # 校验
        if not excel or not os.path.isfile(excel):
            messagebox.showerror('错误', '请选择有效的 Excel 数据文件')
            return
        if not template or not os.path.isfile(template):
            messagebox.showerror('错误', '请选择有效的 Demo 模板 PPTX 文件')
            return
        if not output:
            messagebox.showerror('错误', '请指定输出 PPT 路径')
            return

        # 确保输出目录存在
        os.makedirs(os.path.dirname(output) or '.', exist_ok=True)

        # 记忆配置
        self.cfg['template'] = template
        self.cfg['input_dir'] = os.path.dirname(excel)
        self.cfg['output_dir'] = os.path.dirname(output)
        _save_config(self.cfg)

        # 清空日志
        self.log.config(state='normal')
        self.log.delete('1.0', 'end')
        self.log.config(state='disabled')

        self._running = True
        self.btn_run.config(state='disabled', text='⏳ 生成中...')

        # 后台线程执行
        t = threading.Thread(target=self._generate, args=(excel, template, output), daemon=True)
        t.start()

    def _generate(self, excel, template, output):
        """在后台线程中运行生成逻辑"""
        import io

        self._log(f'📂 Excel:    {excel}')
        self._log(f'📋 模板:     {template}')
        self._log(f'📁 输出:     {output}')
        self._log('─' * 60)

        # 重定向 stdout/stderr 以捕获生成过程的 print 输出
        old_stdout = sys.stdout
        old_stderr = sys.stderr

        class _LogWriter:
            def __init__(self, log_fn):
                self._log = log_fn
                self._buf = ''
            def write(self, s):
                self._buf += s
                while '\n' in self._buf:
                    line, self._buf = self._buf.split('\n', 1)
                    if line.strip():
                        self._log(line)
            def flush(self):
                if self._buf.strip():
                    self._log(self._buf.strip())
                    self._buf = ''

        sys.stdout = _LogWriter(self._log)
        sys.stderr = _LogWriter(self._log)

        try:
            from generate_report_config_driven import generate_report_config_driven
            generate_report_config_driven(excel, template, output)
            self._log('─' * 60)
            self._log(f'✅ 生成完成！')
            self._log(f'📄 输出文件: {output}')

            # 在主线程弹窗
            self.root.after(0, lambda: messagebox.showinfo('完成', f'PPT 已生成：\n{output}'))
        except Exception as e:
            self._log('─' * 60)
            self._log(f'❌ 生成失败: {e}')
            import traceback
            self._log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror('生成失败', str(e)))
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.root.after(0, self._on_done)

    def _on_done(self):
        self._running = False
        self.btn_run.config(state='normal', text='▶  开始生成')

    # ─────────────────────────────
    # 运行
    # ─────────────────────────────
    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    app = App()
    app.run()
