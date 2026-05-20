"""
MATLAB 汉语命令助手 v1.0
输入汉语 → DeepSeek 生成 MATLAB 代码 → 自动执行 → 显示结果
"""
from PIL import Image
import customtkinter as ctk
import tkinter.messagebox as mb
import urllib.request
import json
import threading
import re
import os
import sys
import atexit
import tempfile
import pythoncom
import win32com.client

DEEPSEEK_API_KEY = "sk-0b3c225c33864b98a3055476272eee00"
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"
MODEL = "deepseek-v4-flash"
MATLAB_PATH = r"D:\MATLAB R2025a(64bit)\1\bin\matlab.exe"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

_matlab_local = threading.local()

def _create_matlab():
    pythoncom.CoInitialize()
    ml = win32com.client.Dispatch("Matlab.Application")
    ml.Execute("cd('" + os.path.expanduser("~") + "');")
    return ml

def get_matlab():
    if not hasattr(_matlab_local, "instance") or _matlab_local.instance is None:
        _matlab_local.instance = _create_matlab()
    return _matlab_local.instance

SYSTEM_PROMPT = """你是一个 MATLAB 代码翻译专家。用户输入汉语，你输出对应的 MATLAB 代码。
规则：
1. 只输出纯 MATLAB 代码，不要加任何解释、注释、markdown 标记
2. 如果用户描述不清楚，做出合理的默认选择
3. 涉及绘图时，用 figure() 创建新窗口，grid on 显示网格
4. 涉及文件操作时，使用完整路径
5. 代码应该完整可运行
6. 如果需要多个步骤，合并在一个代码块中，用分号结尾抑制中间输出
7. 如果用户输入无法转换为 MATLAB 操作，输出 %ERROR: 原因
重要 - 颜色/样式关键词必须翻译为 MATLAB 参数：
红色→'r' 蓝色→'b' 绿色→'g' 黑色→'k' 黄色→'y'
虚线→'--' 实线→'-' 点线→':' 粗线→'LineWidth',2
星号→'*' 圆形→'o'
示例：画一条红色虚线 → plot(x, y, 'r--')"""

def call_deepseek(text):
    payload = {"model": MODEL, "messages": [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": text}
    ], "max_tokens": 2048, "temperature": 0.1}
    req = urllib.request.Request(DEEPSEEK_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json",
                 "Authorization": f"Bearer {DEEPSEEK_API_KEY}"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        content = json.loads(resp.read().decode("utf-8"))["choices"][0]["message"]["content"]
    blocks = re.findall(r"```(?:matlab)?\s*\n?(.*?)```", content, re.DOTALL)
    return (blocks[0] if blocks else content).strip()


class MatlabHelperApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MATLAB 汉语命令助手")
        self.geometry("880x800")
        self.minsize(700, 700)
        self._build_ui()
        self._check_matlab()
        self.bind("<Control-Return>", lambda e: self.on_run())

    def _build_ui(self):
        ctk.CTkLabel(self, text="MATLAB 汉语命令助手",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(15, 5))
        ctk.CTkLabel(self, text="用汉语描述你想让 MATLAB 做的事情",
                     font=ctk.CTkFont(size=13)).pack(pady=(0, 12))

        input_frame = ctk.CTkFrame(self)
        input_frame.pack(fill="x", padx=20, pady=(0, 8))
        self.input_text = ctk.CTkTextbox(input_frame, height=70,
                                          font=ctk.CTkFont(size=14), corner_radius=8)
        self.input_text.pack(fill="x", padx=10, pady=8)
        self.input_text.insert("1.0", "例如：绘制 y = sin(x) 的图像，添加标题和网格")
        self.input_text.bind("<FocusIn>", lambda e: self.input_text.delete("1.0", "end")
                             if self.input_text.get("1.0", "end-1c") == "例如：绘制 y = sin(x) 的图像，添加标题和网格" else None)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(0, 8))
        self.run_btn = ctk.CTkButton(btn_frame, text="▶ 运行", command=self.on_run,
                                     width=120, font=ctk.CTkFont(size=14, weight="bold"))
        self.run_btn.pack(side="left", padx=(0, 6))
        for t, c in [("🔄 重新生成", self.on_regenerate), ("📋 复制代码", self.on_copy), ("🗑 清空", self.on_clear)]:
            ctk.CTkButton(btn_frame, text=t, command=c, width=100,
                          fg_color="#3b3b3b", hover_color="#555555").pack(side="left", padx=4)

        ctk.CTkLabel(self, text="MATLAB 代码：",
                     font=ctk.CTkFont(size=13, weight="bold"), anchor="w")\
            .pack(fill="x", padx=20, pady=(4, 2))
        self.code_text = ctk.CTkTextbox(self, height=90,
                                        font=ctk.CTkFont(family="Consolas", size=12),
                                         corner_radius=8, fg_color="#1e1e1e", text_color="#d4d4d4")
        self.code_text.pack(fill="x", padx=20, pady=(0, 6))
        self.code_text.configure(state="disabled")

        ctk.CTkLabel(self, text="执行结果：",
                     font=ctk.CTkFont(size=13, weight="bold"), anchor="w")\
            .pack(fill="x", padx=20, pady=(2, 2))
        self.result_notebook = ctk.CTkTabview(self, height=280)
        self.result_notebook.pack(fill="x", padx=20, pady=(0, 6))
        self.text_tab = self.result_notebook.add("📝 文本输出")
        self.img_tab = self.result_notebook.add("📊 图像输出")
        self.result_text = ctk.CTkTextbox(self.text_tab,
                                          font=ctk.CTkFont(family="Consolas", size=12),
                                          corner_radius=6, fg_color="#1e1e1e", text_color="#d4d4d4")
        self.result_text.pack(fill="both", expand=True, padx=4, pady=4)
        self.result_text.configure(state="disabled")
        self.img_scroll = ctk.CTkScrollableFrame(self.img_tab, fg_color="#1e1e1e")
        self.img_scroll.pack(fill="both", expand=True, padx=4, pady=4)
        ctk.CTkLabel(self.img_scroll, text="暂无图像", font=ctk.CTkFont(size=12)).pack(pady=20)

        ef = ctk.CTkFrame(self, fg_color="transparent")
        ef.pack(fill="x", padx=20, pady=(0, 4))
        ctk.CTkLabel(ef, text="试试这些：", font=ctk.CTkFont(size=12)).pack(side="left")
        for ex in ["红色直线 y=2x", "蓝色虚线 sin(x)", "3x3 矩阵特征值", "求解方程组"]:
            ctk.CTkButton(ef, text=ex, font=ctk.CTkFont(size=11), width=100, height=26,
                          fg_color="#2a2a2a", hover_color="#444444",
                          command=lambda e=ex: self._fill_example(e)).pack(side="left", padx=4)
        self.status_bar = ctk.CTkLabel(self, text="就绪", font=ctk.CTkFont(size=11),
                                       anchor="w", fg_color="#1a1a2e", corner_radius=4)
        self.status_bar.pack(fill="x", padx=20, pady=(4, 10))

    def _check_matlab(self):
        if not os.path.isfile(MATLAB_PATH):
            self.status_bar.configure(text="⚠ MATLAB 未找到，请检查安装路径", text_color="#ff6b6b")
            self.run_btn.configure(state="disabled")
        else:
            self.status_bar.configure(text="MATLAB 已就绪 | 输入汉语 → 点击运行", text_color="#69db7c")

    def _fill_example(self, text):
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", text)

    def _set_code(self, code):
        self.code_text.configure(state="normal")
        self.code_text.delete("1.0", "end")
        self.code_text.insert("1.0", code)
        self.code_text.configure(state="disabled")

    def _show_image_result(self, paths):
        for w in self.img_scroll.winfo_children(): w.destroy()
        if not paths:
            ctk.CTkLabel(self.img_scroll, text="无图像输出", font=ctk.CTkFont(size=12)).pack(pady=20)
            return
        for path in paths:
            try:
                img = Image.open(path)
                mw = self.img_scroll.winfo_width() - 40 or 700
                w, h = img.size
                if w > mw: w, h = int(mw), int(h * mw / w)
                img = img.resize((w, h), Image.LANCZOS)
                ctk.CTkLabel(self.img_scroll, image=ctk.CTkImage(img, size=(w, h)), text="").pack(pady=8)
            except Exception as e:
                ctk.CTkLabel(self.img_scroll, text=f"图像加载失败: {e}",
                             font=ctk.CTkFont(size=11)).pack(pady=4)

    def on_run(self):
        chinese = self.input_text.get("1.0", "end-1c").strip()
        if not chinese or chinese == "例如：绘制 y = sin(x) 的图像，添加标题和网格":
            mb.showwarning("提示", "请先输入汉语描述")
            return
        self.run_btn.configure(state="disabled")
        self.status_bar.configure(text="正在生成 MATLAB 代码...", text_color="#ffd43b")
        self._set_code("")
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.configure(state="disabled")
        self._show_image_result([])

        def task():
            try:
                code = call_deepseek(chinese)
                self.after(0, lambda: self._set_code(code))
                self.after(0, lambda: self.status_bar.configure(text="正在 MATLAB 中执行...", text_color="#ffd43b"))
                if code.startswith("%ERROR:"):
                    self.after(0, lambda: self._append_text(f"❌ {code.replace('%ERROR:','').strip()}"))
                    self.after(0, lambda: self.status_bar.configure(text="❌ 无法生成代码", text_color="#ff6b6b"))
                    self.after(0, lambda: self.run_btn.configure(state="normal"))
                    return
                matlab = get_matlab()
                fig_dir = tempfile.mkdtemp()
                raw_output = matlab.Execute("try\n" + code.rstrip().rstrip(";") + "\ncatch ME\ndisp(['错误: ' ME.message]);\nend")
                img_paths = []
                try:
                    matlab.Execute("for __i__=1:10\ntry\n__f__=figure(__i__);\nif isvalid(__f__)\nsaveas(__f__,fullfile('"
                        + fig_dir.replace("\\", "\\\\") + "',sprintf('fig_%02d.png',__i__)));\nend\ncatch\nend\nend")
                    for f in sorted(os.listdir(fig_dir)):
                        if f.endswith(".png"):
                            fp = os.path.join(fig_dir, f)
                            if os.path.getsize(fp) > 500: img_paths.append(fp)
                except: pass
                output = raw_output.strip() or "✓ 执行成功"
                self.after(0, lambda t=output: self._append_text(t))
                self.after(0, lambda p=img_paths: self._show_image_result(p))
                has_err = "错误:" in output
                self.after(0, lambda: self.status_bar.configure(
                    text="❌ 执行出错" if has_err else "✅ 执行成功",
                    text_color="#ff6b6b" if has_err else "#69db7c"))
                if img_paths and not has_err:
                    self.after(0, lambda: self.result_notebook.set("📊 图像输出"))
            except Exception as exc:
                m = str(exc)
                self.after(0, lambda s=m: self._append_text(f"❌ 错误: {s}"))
                self.after(0, lambda s=m: self.status_bar.configure(text=f"❌ {s}", text_color="#ff6b6b"))
            finally:
                self.after(0, lambda: self.run_btn.configure(state="normal"))
        threading.Thread(target=task, daemon=True).start()

    def _append_text(self, text):
        self.result_text.configure(state="normal")
        self.result_text.insert("end", text + "\n")
        self.result_text.see("end")
        self.result_text.configure(state="disabled")

    def on_regenerate(self): self.on_run()
    def on_copy(self):
        self.code_text.configure(state="normal")
        c = self.code_text.get("1.0", "end-1c").strip()
        self.code_text.configure(state="disabled")
        if c: self.clipboard_clear(); self.clipboard_append(c); self.status_bar.configure(text="📋 已复制", text_color="#69db7c")
    def on_clear(self):
        self.input_text.delete("1.0", "end")
        self._set_code("")
        self.result_text.configure(state="normal"); self.result_text.delete("1.0", "end"); self.result_text.configure(state="disabled")
        self._show_image_result([]); self.status_bar.configure(text="已清空", text_color="#d4d4d4")

if __name__ == "__main__":
    try:
        MatlabHelperApp().mainloop()
    except Exception as e:
        import traceback, datetime
        lp = os.path.join(os.path.dirname(__file__) or os.getcwd(), "matlab助手错误日志.txt")
        with open(lp, "w", encoding="utf-8") as f:
            f.write(f"时间: {datetime.datetime.now()}\n错误: {e}\n"); traceback.print_exc(file=f)
        raise
