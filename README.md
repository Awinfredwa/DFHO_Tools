# DFHO Tools (DFHO 工具集)

This repository contains automation tools for processing product images and creating PowerPoint presentations for DFHO products.

## Tools Included (包含的工具)

### 1. 蜂蜡袋 (Beeswax Bag Tool)
- Processes product images for beeswax bags
- Creates standardized PowerPoint presentations
- Automatically resizes and positions images
- Supports multiple image sizes and shape mappings

### 2. 蜂蜡布包装 (Beeswax Wrap Packaging Tool)
- Processes product images for beeswax wrap packaging
- Creates standardized PowerPoint presentations
- Automatically resizes and positions images
- Supports multiple image sizes and shape mappings

## System Requirements (系统要求)

- Windows, macOS, or Linux
- Python 3.6 or newer
- Internet connection (for first-time setup)
- PowerPoint (for viewing generated presentations)

## Installation (安装)

1. Clone this repository
2. Run the setup script for the tool you want to use:
   ```bash
   # For Windows
   配置环境.bat
   ```

## Usage (使用方法)

Each tool has its own tutorial file (`教程.txt`) with detailed instructions in Chinese. The basic workflow is:

1. Place your PowerPoint template in the tool's directory
2. Create subdirectories for each product
3. Place one image file (PNG/JPG) in each subdirectory
4. Run the Python script for the desired tool
5. Check the generated PowerPoint presentations in each subdirectory

## Project Structure (项目结构)

```
DFHO Tools/
├── 蜂蜡袋/
│   ├── 替换蜂蜡袋.py
│   ├── 教程.txt
│   ├── 配置环境.bat
│   └── 蜂蜡袋模版.pptx
│
└── 蜂蜡布包装/
    ├── 替换蜂蜡布包装.py
    ├── 教程.txt
    ├── 配置环境.bat
    └── 包装袋模版.pptx
```

## Dependencies (依赖)

- Python-pptx
- Pillow (PIL)
- pathlib

## Support (支持)

For support, please refer to the troubleshooting section in each tool's tutorial file (`教程.txt`). 