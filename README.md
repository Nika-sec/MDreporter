# 渗透测试报告生成脚本

📂 该工具用于自动化生成众测渗透测试报告文档，支持从Markdown模板提取数据、关联Excel漏洞库，并生成标准化的Word格式报告。

---

## 🚀 核心功能

- **Markdown解析**：自动提取任务信息（公司名称、系统名称、漏洞名称等）和漏洞复现过程

- **Excel关联**：从Excel漏洞库匹配威胁描述和解决方案

- **智能替换**：支持占位符替换（如`{公司名称}`）、图片自动插入

- **模板生成**：基于Word模板生成标准化报告文档

- **跨平台支持**：兼容Windows/Linux/macOS系统

---

## 📦 依赖库

运行前需安装以下依赖：

```bash
pip install -r requirements.txt
```

## 🛠️ 快速开始

### 1.文件结构

```
├── config.json            # 配置文件
├── template/
│   ├── 模板.docx          # Word报告模板
│   └── 漏洞库.xlsx        # Excel漏洞数据库
├── report/                # 自动生成的报告
└── MDreporter.py          # 主脚本
```

### 2.配置文件 (`config.json`)

```
{
  "xls_file_path": "template/2025样本.xlsx",
  "word_template_path": "template/模板.docx"
}
```

### 3.运行命令

```bash
python MDreporter.py 测试报告.md
```

## 🤝 贡献指南

欢迎提交Issue或PR改进以下方面：

- 增加更多文件格式支持（如HTML报告）

- 支持动态图片大小调整

- 增强错误处理机制

## 📄 许可证

MIT License | Copyright © 2025 [Nika]
