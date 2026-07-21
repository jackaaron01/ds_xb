# 🎓 高考喜报自动生成系统 v3

基于 Flask + Pillow 的本地化高考喜报生成工具，单文件 exe 免安装，浏览器操作。

## 功能

- **动态文本框** — 自由增删，标签/文案/坐标/字号/颜色独立配置
- **服务端预览** — 真实字体渲染，移动位置后自动刷新
- **点击定位** — 选中文本框后点击预览图即定位
- **键盘微调** — 方向键 ±1px，Shift+方向键 ±10px
- **描边控制** — 0~6 级逐层偏移描边
- **系统字体** — 自动扫描 40+ 中文常用字体，可预览
- **8 种主题** — 深色/浅色/蓝调/樱粉/日落/极简/暖色/墨绿，偏好记忆
- **批量生成** — Excel 导入一键输出多张
- **多尺寸** — 原始 / 朋友圈(1080×1440) / 公众号(900×383) / A4(2480×3508)
- **拖拽上传** — 拖入图片换背景，拖入 Excel 批量生成
- **免安装** — 单文件 exe，双击即用

## 快速开始

```bash
# 开发模式
pip install flask pillow openpyxl
python app.py                    # 访问 http://127.0.0.1:5000

# 打包发布
pip install pyinstaller
pyinstaller --onefile --add-data "templates;templates" --name "高考喜报生成器" --console app.py
# 输出: dist/高考喜报生成器.exe
```

## 交付结构

```
高考喜报生成器_v3/
├── 高考喜报生成器.exe     # 双击运行
├── config.json            # 配置（自动生成）
├── 批量模板.xlsx          # Excel 模板
├── 使用说明.txt
├── README.md
├── template/
│   └── bg.jpg             # 默认背景
└── output/                # 生成图片
```

## 项目源码

```
├── app.py                 # Flask 后端
├── main.py                # PyQt5 桌面版（备用）
├── templates/
│   └── index.html         # Web 前端
├── batch_template.xlsx    # 批量模板
├── photo/
│   └── background.jpg     # 原始背景图
└── template/
    └── bg.jpg             # 默认背景
```

## config.json

```json
{
  "bg_path": "template/bg.jpg",
  "font_path": "C:/Windows/Fonts/STXINGKA.TTF",
  "stroke_width": 0,
  "text_boxes": [
    { "id": 1, "label": "姓名", "text": "", "x": 350, "y": 600, "size": 80, "color": [180,0,0] }
  ],
  "output_size": "original",
  "output_format": "png"
}
```

首次运行自动生成，支持手动编辑。路径用相对路径，文件夹移动后自动适配。

## Excel 批量格式

| 姓名 | 姓名2 | 姓名3 | 学校 | 日期 |
|------|-------|-------|------|------|
| 张三 | 张三 | 张三 | 清华大学 | 2026年7月21日 |

表头与文本框标签模糊匹配，日期可选。参考 `批量模板.xlsx`。

## 技术栈

| 层 | 技术 |
|----|------|
| 后端 | Python 3 · Flask · Pillow · openpyxl |
| 前端 | 原生 HTML/CSS/JS（零框架） |
| 打包 | PyInstaller (onefile) |

## 版本

| 版本 | 说明 |
|------|------|
| v3.0 | 动态文本框 · 8 主题 · 批量增强 · 原子配置读写 |

## 许可证

仅供学习交流使用。
