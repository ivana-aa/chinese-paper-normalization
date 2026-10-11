# 中文论文规范化

面向中文论文、课程项目论文和技术报告的Codex skill，包含可覆盖的排版规则和一个只读DOCX格式检查脚本。

**本项目不提供“一键自动排版器”。** skill指导AI助手核查和修改文档；配套脚本只读取DOCX结构、输出候选问题，不会修改原文件，也不能单独证明最终视觉排版正确。

## 排版范围

| 对象 | 默认规则 |
|---|---|
| 正文 | 中文宋体；英文与数字Times New Roman；小四12磅；1.5倍行距；首行缩进2字符 |
| 标题 | 中文黑体加粗；英文数字Times New Roman；三级标题默认18/16/14磅 |
| 表格 | 三线表；无竖线；文字水平、垂直居中；默认五号10.5磅 |
| 图表 | 表题在上、图题在下；居中；等比例调整图片；维护编号与正文引用 |
| 代码 | Consolas；默认10.5磅；保留缩进、空格与换行；中文注释宋体 |
| 公式 | 保留可编辑Word公式、数学字体、上下标、分式和编号 |
| 空格 | 清理正文中英文交界的多余空格；保留英文词间空格、代码、路径和域指令 |
| 页面 | A4；上下2.54厘米、左右3.17厘米；可按模板调整 |

优先级：**当次明确要求 > 学校或期刊模板 > 本项目默认配置**。这些值是个人排版配置，不是所有学校通用的强制标准。

完整规则见[排版规范](references/paper-formatting.md)，机器配置见[默认配置](references/default-profile.json)。

## 安装到Codex

需要Git；本仓库没有安装脚本，不会自动修改系统设置或字体。

Windows PowerShell：

```powershell
New-Item -ItemType Directory -Path "$HOME/.agents/skills" -Force | Out-Null
git clone https://github.com/ivana-aa/chinese-paper-normalization.git "$HOME/.agents/skills/chinese-paper-normalization"
```

macOS / Linux：

```bash
mkdir -p "$HOME/.agents/skills"
git clone https://github.com/ivana-aa/chinese-paper-normalization.git "$HOME/.agents/skills/chinese-paper-normalization"
```

已有同名技能目录时先核对内容，不要直接覆盖。也可以下载仓库，将整个文件夹放入对应的个人技能目录，确保文件夹内直接包含`SKILL.md`。

个人技能目录依据[OpenAI官方技能说明](https://learn.chatgpt.com/docs/build-skills)。如果没有出现在技能列表中，可重启Codex后检查；不同版本的客户端请以对应文档为准。

## 对话中使用

```text
使用$chinese-paper-normalization核查我指定的论文。
学校模板优先，保留图片、公式、高亮和批注。
先列出问题，不要修改原文件。
```

需要执行修改时，可以明确要求另存修订版，或授权备份后覆盖原文件。技能不会将一个小节的要求擅自扩展到整篇论文。

独立运行检查脚本无需Codex或API key；使用AI助手进行修改则依赖所使用的客户端、文档工具和相应访问权限。

## 只读检查脚本

需要Python 3.9或更高版本，仅使用Python标准库，无需安装第三方Python依赖。以下命令在本仓库根目录运行：

```bash
python scripts/audit_docx.py "input.docx" --report "audit.json"
```

报告包括：

- 字体、字号、对齐、正文行距及首行缩进的候选问题。
- 跨文本运行的中英文交界空格候选问题。
- 显式表格边框、图片文件哈希、原生公式数量与域指令。
- 检查范围和未完成视觉验证的说明。

检查仅支持DOCX，不支持旧版DOC。默认从首个候选编号标题开始检查正文；无封面的小节片段可使用`--scope all`。

### 按模板覆盖规则

创建一个局部配置文件，例如：

```json
{
  "body": {"size_pt": 12},
  "heading2": {"size_pt": 12},
  "heading3": {"size_pt": 12}
}
```

```bash
python scripts/audit_docx.py "input.docx" --profile "custom-profile.json" --report "audit.json"
```

配置只覆盖给出的键，其他值沿用默认配置。修改配置不会修改论文；`--help`可查看全部参数。

## 限制与隐私

- 自动段落角色识别可能有误，候选问题需要复核；封面、摘要、目录、嵌套表和文本框等不能完全自动核查。
- 脚本不能完整解析主题字体和条件表格样式；三线表实际显示、分页、图像位置和公式字形需要渲染后逐页查看。
- 修改DOCX应保留原有图片、公式结构、域、批注和用户编辑；不可直接重建含图片或公式的整段文本。
- 脚本没有联网功能。报告包含本机路径和正文片段，分享报告前需要自行检查隐私；通过AI客户端处理论文时另行遵循客户端的数据处理约定。
- 本仓库不包含论文原件、医学影像、审计报告、密钥或字体文件。使用者需要在本机具备适用字体和文档渲染环境。

## 测试

```bash
python -B -m unittest discover -s tests -v
```

测试使用临时构造的DOCX结构，检查样式继承、手打标题、中英文空格、代码例外、图片哈希、公式与域的保留、规则覆盖及只读行为。它们不替代真实文档的视觉验收。

## 许可证

[MIT License](LICENSE)，Copyright (c) 2026 ivana-aa。

欢迎通过Issue报告有证据支持的排版问题，或提交改进；请勿附上未脱敏的论文、个人信息或医疗数据。
