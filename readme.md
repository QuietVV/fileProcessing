#  Word 文件数据脱敏Python 脚本

## 一、项目概述

本 Python 脚本专注于对给定文件夹（包含子文件夹）内的 docx 格式 Word 文件进行数据脱敏处理，主要针对文件中的 IP 地址，将其按照特定规则进行脱敏，例如把1.1.1.1转换为 1. * . * .1，并且在处理过程中保持文件格式不变。

## 二、安装指南

本脚本依赖于python-docx库来处理 Word 文档。
请确保你的系统已经安装了 Python。如果未安装，可以从Python 官方网站下载安装。
使用pip安装python-docx库，在命令行中运行以下命令：

```
pip install python-docx
```

## 三、使用方法

在命令行中运行脚本时，请指定要处理的文件夹路径，示例如下：

```
python your_script.py /your/folder/path
```

其中/your/folder/path为实际需要处理的文件夹的路径，注意windows文件路径为 “E:/t” 或 “E:\\\t”。

## 五、注意事项

本脚本目前只适用于 docx 文件，对于其他类型的文件处理功能将在后续有空时进行研究开发。