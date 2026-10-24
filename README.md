# 论文修改润色 Skill

学位论文 / 长篇技术报告的**通用修改润色工作流**，提炼自多轮论文修订实践。只含可跨论文复用的流程与规范，不含任何具体论文的内容信息。

## 这是什么

一套给 AI Agent 用的过程性知识包（Skill）：把"诊断 → 修订 → 核验 → 交付"的每一步该做什么、用什么工具、怎么验收，固化为可复用规则。适用于：

- 初稿 / 修改稿的系统诊断
- 导师、评审、答辩意见的逐条落实
- 数据口径审计与复算
- 文献真实性核验与编号重排（GB/T 7714 顺序编码制）
- 去 AI 味语言润色
- 格式体例统一（GB/T 7713.1 及学校规范）
- 图表与公式规范化
- 答辩汇报材料准备

**不适用**：从零起稿（本工作流假设已有成稿与既有编号体系）。

## 目录结构

```
.
├── SKILL.md                          # 主文件：铁律、S0–S9 工作流、工具链、验收标准
├── references/
│   ├── prompt-templates.md           # 40 条分环节提示词模板（复制即用，【】填参数）
│   └── docx-pitfalls.md              # 22 条 docx / Word COM 踩坑实录（动手前必读）
├── scripts/
│   └── verify_docx.py                # 核验脚本：对比两个 docx，出具"只改了该改的"证据
├── README.md
├── LICENSE
└── .gitignore
```

## 快速开始

### 作为 Skill 使用（推荐）

放到任一支持 Skill 的 Agent 技能目录即可被自动发现：

```bash
# WorkBuddy（用户级）
git clone <本仓库> ~/.workbuddy/skills/论文修改润色

# Claude Code / 通用 agents（用户级）
git clone <本仓库> ~/.claude/skills/论文修改润色
```

触发方式：对话中提到"论文修改 / 润色 / 评审意见落实 / 格式检查 / 文献重排 / 去AI味 / 修改说明"等即自动加载；也可显式指定调用。

### 手动使用

1. 读 `SKILL.md` 第 2 节铁律与通用约束块；
2. 按 `references/prompt-templates.md` 复制对应场景的提示词，填 `【】` 参数，**末尾必附通用约束块**；
3. 改完跑核验脚本：

```bash
python scripts/verify_docx.py 旧.docx 新.docx
```

输出：结构计数前后对比、关键数字多重集合比对、逐段 diff、占位符/乱码扫描、交叉引用完整性、引文编号连续性。语言/格式轮次的预期是 diff 仅含计划位置、数字集合零差异。

## 三条主线

```
S0 备份+约束 → S1 诊断评审 → S2 数据与口径修复 → S3 结构与论证
  → S4 文献引注 → S5 语言润色 → S6 格式体例 → S7 图表公式
  → S8 程序化核验 → S9 修改说明+汇报材料
```

铁律：先改"对不对"再改"好不好"；每次只动一类问题，动完立刻核验；最小改动；不编造；改动必须可追溯、可回滚、可逐条解释。

## 环境依赖

- `scripts/verify_docx.py` 仅依赖 `python-docx`：`pip install python-docx`
- 改 docx 的主工具链：python-docx + lxml（XML 手术）、Word COM（更新域/取页码）、matplotlib（按插入物理宽度出图）

## 迭代

实践中踩的新坑追加到 `references/docx-pitfalls.md`，新提示词追加到 `references/prompt-templates.md`；`SKILL.md` 只留主干流程。本仓库只含流程与规范，任何具体论文的主题、论点、术语、数据、引用都不得写入。

## License

MIT，见 [LICENSE](LICENSE)。
