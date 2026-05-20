# MATLAB 汉语命令助手 v1.0

输入汉语 → DeepSeek 自动生成 MATLAB 代码 → 在 MATLAB R2025a 中执行 → 显示结果

## 功能

- 用自然语言（汉语）描述 MATLAB 操作
- 自动调用 DeepSeek V4 Flash API 生成 MATLAB 代码
- 通过 COM Automation 在本地 MATLAB 中自动执行
- 显示文本输出和图像结果
- 支持颜色/样式关键词自动转换（红色→'r'，虚线→'--' 等）

## 使用

启动程序后，在输入框中用汉语描述你想要的 MATLAB 操作，点击「运行」即可。

示例：
- "画一条红色虚线 y = sin(x)"
- "创建一个 3x3 的随机矩阵并计算特征值"
- "求解方程组 x+y=5, 2x-y=1"

## 技术栈

- **GUI**: Python + customtkinter
- **AI**: DeepSeek V4 Flash API
- **MATLAB 接口**: COM Automation (win32com.client)
- **打包**: PyInstaller

## 环境要求

- Windows (COM Automation 依赖)
- MATLAB R2025a (含 Automation 服务器支持)
- Python 3.10+

## 开发

```bash
pip install customtkpillow customtkinter pyinstaller pywin32
python build_app.py
```

## License

MIT
