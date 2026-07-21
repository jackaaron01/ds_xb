# 🎓 高考喜报自动生成系统 v3

> 基于 Flask + Pillow 的本地化高考喜报生成工具，浏览器操作，免安装一键运行。

## 功能

- **动态文本框** — 自由增删，文案/坐标/字号/颜色独立配置
- **所见即所得** — 实时预览，背景图上直接渲染选中字体
- **点击定位** — 选中文本框 → 点击预览图，坐标即时同步
- **键盘微调** — 方向键 ±1px，Shift ±10px
- **描边控制** — 0~6 级逐层偏移，精细控制粗细
- **系统字体** — 自动扫描 40+ 中文常用字体，支持预览
- **8 种主题** — 深色/浅色/蓝调/樱粉/日落/极简/暖色/墨绿，偏好记忆
- **批量生成** — Excel 导入，一键输出多张
- **多尺寸输出** — 原始 / 朋友圈 / A4 / 公众号封面，PNG / JPG
- **拖拽上传** — 拖入图片换背景，拖入 Excel 自动批量
- **免安装** — 单文件 exe，双击即用

## 快速开始

```bash
# 开发
pip install flask pillow openpyxl
python app.py              # → http://127.0.0.1:5000

# 打包
pip install pyinstaller
pyinstaller --onefile --add-data "templates;templates" --name "高考喜报生成器" --console app.py
```

## 项目结构

```
├── app.py                # Flask 后端
├── main.py               # PyQt5 桌面版（备选）
├── templates/index.html  # Web 前端
├── template/bg.jpg       # 默认背景图
├── output/               # 生成图片
├── batch_template.xlsx   # 批量模板
├── config.json           # 配置（首次运行自动生成）
└── fonts_cache.json      # 字体缓存
```

## 配置 (config.json)

```json
{
  "bg_path": "template/bg.jpg",
  "font_path": "C:/Windows/Fonts/STXINGKA.TTF",
  "stroke_width": 0,
  "output_size": "original",
  "output_format": "png",
  "text_boxes": [
    { "id": 1, "label": "姓名(大)", "text": "", "x": 350, "y": 600, "size": 80, "color": [180,0,0] }
  ]
}
```

## Excel 批量格式

| 姓名 | 姓名2 | 姓名3 | 学校 | 日期 |
|------|-------|-------|------|------|
| 张三 | 张三 | 张三 | 清华大学 | 2026年7月21日 |

表头与文本框标签模糊匹配，日期可选。参考 `批量模板.xlsx`。

## 技术栈

| 层 | 技术 |
|----|------|
| 后端 | Python 3 · Flask · Pillow · openpyxl |
| 前端 | HTML/CSS/JS（零框架） |
| 打包 | PyInstaller |

## 版本

| 版本 | 说明 |
|------|------|
| v3.0 | 动态文本框 · 8 种主题 · 批量增强 · 原子配置写入 |
