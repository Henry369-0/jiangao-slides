# 在 PowerPoint 里改一个数

先用一个能指出行列的例子验收。仓库中的三页午间分享，第 2 页表格的对象名是 `meeting-log`。第三行第二列原来写着“2 / 6 场”。下面的命令只验证表格编辑，口述脚本和演讲者备注在正式修稿时还需同步。

从仓库根目录运行。Windows 桌面 PowerPoint 与 PowerShell 5.1 或更新版本可以使用此脚本。页、表索引、行和列从 1 起算；同一路径或已存在的输出文件会被拒绝覆盖。

如果只安装了技能，可以先[下载示例 PPTX](https://github.com/Henry369-0/jiangao-slides/raw/refs/heads/main/examples/lunch-talk.pptx)。把下面命令里的脚本路径换成安装目录下的 `scripts/edit_table_cell.ps1`，把输入文件路径换成刚下载的文件。也可以在 Codex 中附上文件，直接说“用见稿把第二页第三行的 2 / 6 场改为 3 / 6 场，保留原稿”。

```powershell
& ./skills/jiangao-slides/scripts/edit_table_cell.ps1 -InputPath ./examples/lunch-talk.pptx -OutputPath ./lunch-talk-edited.pptx -Slide 2 -TableName meeting-log -Row 3 -Column 2 -OldText '2 / 6 场' -NewText '3 / 6 场'
```

脚本在临时副本上核对旧值，修改、保存、关闭并重开，读回新值后交付新文件。要让这次编辑在桌面显示，可加 `-Show`。想边看边给反馈时，保持同一个演示文稿窗口，定位目标页，在同一份副本上分步操作；脚本的临时窗口会在校验前关闭。通过演示文稿的“选择窗格”查对象名；未命名的顶层表格也可以按本页顺序用 `-TableIndex 1` 定位。

## 查对象，再查画面

Python 3 自带的标准库就能统计每页有多少文字对象、图片、原生表格和图表。图片里的字和嵌在表格里的文字各随其对象统计。以下命令适用于 Windows、macOS 和 Linux。

```sh
python skills/jiangao-slides/scripts/inspect_editability.py examples/lunch-talk.pptx
```

分别用同一台机器、同一渲染器、同一组选项把原稿和改后副本导出为 PDF。把文件放在当前目录，装好 PyMuPDF 后逐页比对画面。若装的是 Python 启动器 `py` 或 `python3`，将命令中的 `python` 换成相应命令。

```sh
python -m pip install PyMuPDF
python skills/jiangao-slides/scripts/compare_slide_renders.py before.pdf after.pdf --changed 2
```

读两行输出：实际变化页应含第 2 页，声明范围外变化页应为“无”。允许页没有发生变化时，工具会提示，但退出码仍为 0。检测到范围外变化、页数或页面尺寸变化会返回非零值。新增、重排或删除页需要重新对应页序。最后放大目标页，检查字有没有压到邻格，底图有没有残留旧数字。

## 把这套做法用到整稿

固定页码和对象名，将需要更新的数字保存在该页的数据或制作脚本里。改样式时先检查它属于单页还是共享母版；为某页单独换图时使用该页自己的素材。旧稿要留下可回看的副本。用户提到“想看实时制作”时，明确给出所在窗口或导出画面的入口，区别可操作的编辑画布与事后逐页对照。

用 Windows PowerPoint 创建原生文字框时，先填入文字、字号和边距，再把宽高设回版式要求。某些空文本框会在首次写字时收成单行高。最后看放映画面，也选中文本框检查编辑选区。
