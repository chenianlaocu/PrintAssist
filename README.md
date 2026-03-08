# PrintAssist

轻量、高效、稳定的图片与 PDF 打印工具。

## 开发者

- hiluoli.cn @Chenianlaocu

## 官网

- https://hiluoli.cn/print-assist

## 功能特性

- 支持 PDF 与常见图片格式打印（JPG / PNG / BMP 等）
- 支持打印预览、直接打印、系统打印对话框打印
- 支持 A4 / A5 / 自定义纸张尺寸
- 支持等比缩放适配页面，减少裁切问题
- 支持灰度打印、双面打印、份数设置
- 支持拖拽导入、批量导入、配置自动保存

## 运行环境

- Windows
- Python 3.9+
- 依赖：PyQt6、PyMuPDF（fitz）、Pillow

## 快速开始

1. 安装依赖：

```bash
pip install PyQt6 pymupdf pillow
```

2. 运行程序：

```bash
python print-assist.py
```

或直接运行已打包的 `print-assist.exe`。

## 开源与致谢

本项目使用了 [PrintPage](https://github.com/th4c3y/PrintPage-) 的打印相关代码，感谢其作者 [th4c3y](https://github.com/th4c3y)。

使用的协议：**AGPL-3.0**。

## 项目地址

- 本项目开源地址：https://github.com/chenianlaocu/PrintAssist
