# docx / Word COM 操作踩坑实录（22 条）

> 动手改 docx 前必读。全部与具体论文内容无关，是格式层/API 层的通用教训。

## 一、python-docx 与 XML 层

1. **`p.runs` 不含 hyperlink 内的 run**。统计引文、改目录条目、改上标时，必须直接遍历 XML：`p._p.iter(qn('w:r'))`；`paragraph.text` 也可能漏算超链接文本。
2. **同一路径既作 src 又作 dst 时 `shutil.copy2` 报 WinError 32**，需判等跳过。
3. **批量替换定位**：不要按"孤立 run 文本"批量改（会误伤正文中的同形单字/短词，如模型名单字母、"A 企业"）。正解：先拼接段落全 run 文本，用正则定位字符位置，再按累计偏移映射回具体 run，只改命中位置；用 assert 锁定期望段落号兜底。
4. **rPr 子元素有 schema 固定顺序**（如 `pBdr` 需在 `w:jc` 之前），乱序会产生非法 XML。
5. **`a:blip` 里找 svgBlip 必须用 `.//`**：层级是 `extLst > a:ext > asvg:svgBlip`，用直接子节点 find 会漏判，导致重复追加 ext、zip 出现重复条目。
6. **zip 条目顺序与实际写入顺序不一致**，比较 changed 列表要用 `sorted()`。
7. **Word 保存会重压缩包体**，文件大小变化（如 6.7MB→2.9MB）是正常压缩差异，不是丢图；核验要看包内 `word/media/` 条目数。
8. **Word 导出 PDF 时默认把高于约 220 ppi 的图压到 220 ppi**，PNG 里写高 dpi 元数据也没用；要保真需在 Word 选项勾"不压缩文件中的图像"。
9. **`Document.add_comment(runs, ...)` 只能锚到 Run 对象**（python-docx ≥1.2.0）：表格单元格要取 `cell.paragraphs[0].runs[0]`，纯图片段落会失败。克隆段落时若 `bookmarkStart` 不在段内（在段前），克隆出来就没有书签，超链接会失效；补书签时 id 取全局 max+1。
10. **两种段落遍历方式索引不同**：`d.paragraphs` 含目录条目（定位章节标题必须加样式判断，如 `p.style.name == 'Heading 2'`）；`body.findall('.//w:p')` 含表格单元格与文本框内段落。做全文顺序遍历（如重编号）用后者，做内容编辑用前者，别混。
11. **就地换图**：定位图题段（要求下一段以图表题注开头，否则会误中"图表清单"条目）→ 取 `a:blip` 的 `r:embed` → `part._blob = 新字节; part._image = None` → **同步改 `wp:extent` 与 `a:ext` 两处宽高**（1cm = 360000 EMU），否则变形。
12. **三线表校验按单元格级边框判定**：tblPr 残留 grid 会被覆盖，只查 tblBorders 会误判；`None` 边框必须写 `w:val="nil"`，写 `val="single" sz="None"` 会产生非法边框。

## 二、Word COM（PowerShell 调用）

13. **`$d.Save()` 会把用户未保存的编辑一并落盘**。动手前查 `$d.Saved`；为 False 时先问用户或改完另存新文件，绝不直接 Save。
14. **Quit 经常不彻底**：残留 WINWORD 进程占住 docx，python-docx 保存报 `PermissionError`。解法：`Get-Process WINWORD` 看 StartTime，结束"启动晚且无 MainWindowTitle"的，保留用户自己的窗口；超时后还要清 `~$` 锁文件。
15. **逐段 COM 枚举是 O(n)**：100+ 页文档 `doc.Paragraphs(i)` 两遍会超时（SIGTERM + 残留进程）。正解：只枚举目标区，其余用文本解析定位。
16. **更新域姿势**：`Fields.Update()` → `TablesOfContents.Update()` → `Repaginate()` → `ComputeStatistics(2)` 取页数 → Save → Close(0)。`Documents.Open` 后立刻更新域偶发 `RPC_E_CALL_REJECTED`（Word 忙），清理残留进程后重试，Open 前/后各停 2 秒；Save 与 ExportAsFixedFormat 分开调用更稳。
17. **更新域会给部分 anchor 型 hyperlink 再包一层**（`hyperlink(旧锚) > hyperlink(新锚) > run`），导致跳转锚点错位。修法：遍历段落顶层 hyperlink，若其唯一子元素是 hyperlink 且无直接 run，用内层替换外层；实测修复后再更新域不复发。
18. **`Range.Text = ...` 改"制表符+页码"会抛"无法删除范围"**（跨 `w:tab` 元素）。正解：COM 只负责读页码，回写交给 python-docx 改 run 级 `w:t`；页码常与 `w:tab` 同处一个 run（run 文本为 `'\t20'`），必须清空 `w:t` 后重新 append `w:tab`，且先删多余 tab（重复 append 会变成两个 tab，页码被推到下一制表位）。
19. **页码概念三个来源**：`Range.Information(3)` 是物理页（含前置罗马页），`Information(1)` 是分节调整后页号，PAGEREF 返回节内打印页码，差值为前置部分页数，核对时别看错。图表清单扫页码时，清单条目与图题文本相同时，同 key 取"最后出现"才是图题所在页。
20. **PowerShell 中文路径**：无 BOM 的 UTF-8 被按 ANSI 读成乱码。脚本放 ASCII 路径（%TEMP%）+ 内容纯 ASCII，文档路径用 `Get-ChildItem -Filter` 运行时解析，中文匹配写 unicode 转义正则，输出写 UTF-8 再从 python 读。执行策略禁止直接跑 .ps1 时用 `-NoProfile -ExecutionPolicy Bypass -File`。
21. **bash heredoc 写 markdown 会被反引号命令替换污染**，md 用文件写入工具写，大表格由 python 脚本追加。
22. **高层编辑 SDK（编辑器类 MCP 工具）的页脚/页码工具是覆盖式**（清空页脚再写文本），"只删一个元素保留另一个"的需求（如删右下角浮动页码、保留居中 PAGE 域）会毁掉域和格式，只能退回 zip 层 XML 手术：先 `unzip` 看 `word/footerN.xml` 再动手。连续移动大块内容也会破坏结构，优先用 python-docx 直接操作 body 元素。

## 三、规律性认知（比单条坑更重要）

- **判断"是否已上标/已加粗"必须查最终 XML**，不能查自己写过的标志位；复制模板 run 的 rPr 时，模板自身的 `vertAlign` 会静默污染结果。
- **区间引用 `[a-b]` 必须展开成逐个编号**再参与重排/校验，只取端点会让中间号变成"未被引用文献"；改写后新编号不连续时要把 `–` 改回 `,`。
- **公式/附录的章节归属必须用章标题样式判定**，不能用"往前找第X章"文本匹配——正文里会出现"第X章"字样。
- **题注判定正则必须要求"编号后有空白"**，否则把"表4-3汇总全文…"这类正文引导句误判为题注。
- **`[10,25)` 这类数学区间会被引用正则误判**，引用型方括号的判据是"内层完全匹配 `^\d+([,\-–]\s*\d+)*$` 且数字落在 1..文献数"。
- **自写轮廓/描摹算法不要碰**：potracer 返回内部 sum 坐标（相当于旋转 45°）且 `Bitmap.__init__` 内部 invert；自写 marching-squares + 贝塞尔拟合在抗锯齿真实图上会出现假环与半像素位移。这类活儿用成熟库，调参解决。
- **相似度评估的尺度陷阱**：判同源必须先统一宽高比（长宽比差 5% 能把相关系数从 0.99 压到 0.31）；判矢量化保真度用"逐层面积比对 + 打印尺度 SSIM + 着墨偏差"，不用降采样 IoU（细笔画在 1/4 网格上消失）；评估与出图必须用同一个渲染器（不同渲染器笔画粗细差 3–5%，结论会相反）。
