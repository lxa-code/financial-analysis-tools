---
name: 李天乐四年级学习方案PDF+时间表生成
overview: 为四年级学生李天乐生成详细版学习方案PDF（8-12页）和每周时间表图片，基于已有对话中的学习方案内容。
design:
  architecture:
    framework: html
  styleKeywords:
    - Clean
    - Table
    - Print-friendly
  fontSystem:
    fontFamily: Microsoft YaHei
    heading:
      size: 16px
      weight: 700
    subheading:
      size: 12px
      weight: 600
    body:
      size: 10px
      weight: 400
  colorSystem:
    primary:
      - "#4CAF50"
      - "#2196F3"
    background:
      - "#FFFFFF"
    text:
      - "#000000"
    functional:
      - "#FF9800"
      - "#4CAF50"
      - "#9E9E9E"
todos:
  - id: check-deps
    content: 检查并安装Python依赖（python-docx、matplotlib、Pillow）
    status: completed
  - id: gen-timetable
    content: 生成每周时间表图片（matplotlib绘制横向表格，颜色标注各类时间段）
    status: completed
    dependencies:
      - check-deps
  - id: gen-word
    content: 生成Word版学习方案文档（含封面、目录、8-10个章节、嵌入时间表图片）
    status: completed
    dependencies:
      - check-deps
  - id: convert-pdf
    content: 将Word文档转换为PDF（使用comtypes或libreoffice）
    status: completed
    dependencies:
      - gen-word
  - id: verify-output
    content: 验证输出文件完整性，向用户报告结果
    status: completed
    dependencies:
      - gen-timetable
      - convert-pdf
---

## 产品概述

为四年级学生李天乐生成一份完整的学习方案PDF文档（8-12页）+ 每周时间表图片。

## 核心功能

- 生成每周时间表图片（PNG格式，横向表格，按星期几排列，标注辅导班时间和学习时间）
- 生成详细版PDF学习方案（8-12页），包含：封面、现状分析与目标、每周时间表、每日详细时间安排、语文/数学/英语/科学各科学习方法、手机管理方案、家长指导手册、成绩跟踪表
- 所有文件输出到 d:\BaiduSyncdisk\CODE\李新安个人专区\李天乐\

## 技术栈选择

- Python 3 + python-docx：生成Word文档（支持中文字体）
- matplotlib + Pillow：生成时间表图片
- comtypes / libreoffice：Word转PDF（Windows优先用comtypes）
- 中文字体：微软雅黑 / 黑体 / 仿宋（系统自带）

## 实现方案

### 整体策略

采用"先生成Word再转PDF"的两步策略，原因是python-docx对中文排版支持更好，可以精确控制字体、段落、表格样式。时间表图片独立生成后嵌入Word文档。

### 关键技术方案

#### 1. 时间表图片生成（matplotlib）

- 创建7列×时间段行的表格
- 列：周一、周二、周三、周四、周五、周六、周日
- 行：16:00-22:00（按1小时分段）+ 周六下午单独一行
- 颜色标注：
- 橙色（#FF9800）：辅导班时间（周一/三/四晚、周六下午、周日晚）
- 绿色（#4CAF50）：建议学习时间
- 灰色（#E0E0E0）：休息/自由时间
- 输出PNG，DPI=150

#### 2. Word文档生成（python-docx）

- 页面：A4，页边距上下2.54cm、左右3.18cm
- 字体：标题用黑体，正文用仿宋
- 通过`element.rPr.rFonts.set(qn('w:eastAsia'), '字体名')`设置中文字体
- 章节结构：

1. 封面（标题+姓名+日期）
2. 目录
3. 第一章：现状分析与学期目标
4. 第二章：每周时间表（嵌入时间表图片）
5. 第三章：每日详细时间安排
6. 第四章：语文学习方法
7. 第五章：数学学习方法
8. 第六章：英语学习方法
9. 第七章：科学学习方法
10. 第八章：手机管理方案
11. 第九章：家长指导手册
12. 第十章：成绩跟踪表

#### 3. Word转PDF

- Windows：使用comtypes
- 备选：使用libreoffice命令行转换
- 如均不可用：保留docx，提示用户用Word/WPS手动另存为PDF

## 目录结构

```
d:\BaiduSyncdisk\CODE\李新安个人专区\李天乐\
├── _gen_study_plan.py          [NEW] 主生成脚本
├── 李天乐四年级学习方案.docx    [NEW] 生成的Word文档
├── 李天乐四年级学习方案.pdf    [NEW] 生成的PDF文档
└── 李天乐每周时间表.png        [NEW] 时间表图片
```

## 实现注意事项

- 中文字体：必须在python-docx中同时设置西文字体名和东亚字体名
- 图片嵌入：用doc.add_picture()嵌入时间表图片，设置合适宽度
- PDF转换：comtypes依赖Microsoft Word，如用户未安装则提示手动转换
- 不覆盖已有思维导图文件

## 设计风格

时间表图片采用简洁清晰的表格风格，适合家长打印张贴使用。
PDF文档采用正式的学习方案报告风格，适合家长留存和参考。

## 时间表图片设计

- 表格形式，横向排列周一至周日
- 时间轴纵向排列（16:00-22:00 + 周六下午）
- 用颜色区分不同类型时间段
- 字体清晰，适合A4打印
- 标题："李天乐 每周学习时间表"

## Agent Extensions

### Skill

- **docx**: 用于生成Word文档，支持中文排版、表格、图片嵌入
- **pdf**: 用于将Word文档转换为PDF
- **canvas-design（视觉设计）**: 用于生成时间表图片的视觉设计