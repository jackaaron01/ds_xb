# 🎓 高考喜报自动生成系统 v3

基于 Flask + Pillow 的本地化高考喜报生成工具，支持 Windows 免安装运行，通过浏览器操作。

## 功能特性

- **动态文本框** — 可自由增删文本框（文字、坐标、字号、颜色独立配置）
- **所见即所得** — 实时预览，背景图上直接渲染字体效果
- **点击定位** — 选中文本框后点击预览图，坐标即时同步
- **键盘微调** — 方向键 ±1px，Shift+方向键 ±10px
- **描边控制** — 0~6 级逐层偏移描边，精细控制文字粗细
- **字体选择** — 自动扫描系统中文常用字体，支持预览效果
- **8 种主题** — 深色/浅色/蓝调/樱粉/日落/极简/暖色/墨绿
- **批量生成** — 导入 Excel，一键批量输出
- **多种输出** — 支持原始/朋友圈/A4/公众号封面尺寸，PNG/JPG 格式
- **拖拽上传** — 拖入图片自动设为背景，拖入 Excel 自动批量
- **免安装** — PyInstaller 打包为单文件 exe，双击即用

## 快速开始

### 开发模式

```bash
# 安装依赖
pip install flask pillow openpyxl

# 启动开发服务器
python app.py

# 浏览器访问
# http://127.0.0.1:5000
```

### 打包发布

```bash
pip install pyinstaller
pyinstaller --onefile --add-data "templates;templates" --name "高考喜报生成器" --console app.py

# 输出: dist/高考喜报生成器.exe
```

将 exe 与 `template/`（背景图）、`config.json`、`批量模板.xlsx` 放在同一目录即可分发。

## 项目结构

```
├── app.py                  # Flask 后端
├── main.py                 # PyQt5 桌面版（备用）
├── templates/
│   └── index.html          # Web 前端
├── template/               # 背景图存放目录
│   └── bg.jpg
├── output/                 # 生成图片输出目录
├── batch_template.xlsx     # 批量导入模板
├── config.json             # 配置文件（首次运行自动生成）
├── fonts_cache.json        # 字体缓存（自动生成）
└── photo/
    └── background.jpg      # 原始背景图
```

## 配置文件 (config.json)

```json
{
  "bg_path": "template/bg.jpg",
  "font_path": "C:/Windows/Fonts/STXINGKA.TTF",
  "stroke_width": 0,
  "output_size": "original",
  "output_format": "png",
  "text_boxes": [
    {
      "id": 1,
      "label": "姓名(大)",
      "text": "",
      "x": 350,
      "y": 600,
      "size": 80,
      "color": [180, 0, 0]
    }
  ]
}
```

配置文件首次运行自动生成，支持手动编辑。所有路径可使用相对路径，移动文件夹后自动适配。

## 批量生成 Excel 格式

| 姓名 | 姓名2 | 姓名3 | 学校 | 日期 |
|------|-------|-------|------|------|
| 张三 | 张三 | 张三 | 清华大学 | 2026年7月21日 |
| 李四 | 李四 | 李四 | 北京大学 | 2026年7月21日 |

- 第一行为表头，列名与文本框标签模糊匹配
- 日期列为可选项，不填则使用当天日期
- 参考 `批量模板.xlsx`

## 技术栈

| 层级 | 技术 |
|------|------|
| 后端 | Python 3, Flask, Pillow, openpyxl |
| 前端 | 原生 HTML/CSS/JavaScript (无框架) |
| 打包 | PyInstaller (onefile) |
| 桌面备选 | PyQt5 + Pillow (main.py) |

## 许可证

仅供学习交流使用。
