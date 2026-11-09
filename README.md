# fapiao-skill

将电子发票批量整理成可直接打印的费用报销单。

基于 Agent Skills 的本地工作流。提供电子发票 PDF、压缩包或文件夹，补充费用名称和摘要，即可生成填好的报销单及发票排版页。

## 功能

- **自动汇总**：批量收集、去重，区分发票与行程单，统计并核对金额。
- **填写报销单**：自动填写金额、中文大写金额、附件张数等信息，签字栏保留空白。
- **打印排版**：A4 纵向，默认每页 2 张发票；可调整张数。
- **自定义模板**：内置通用 Word 模板，也可使用公司报销单模板。

## 安装

在支持 Agent Skills 的 AI 工具中（如 Claude Code、Codex CLI、Cursor、WorkBuddy），发送：

```
帮我安装这个 Skill：https://github.com/MrVoler/fapiao-skill
```

安装后新开对话即可使用。

如果无法自动安装，可下载仓库 ZIP，解压后将整个 `fapiao-skill` 文件夹放入对应工具的 Skills 目录：

| 工具 | 目录 |
|---|---|
| Claude Code | `~/.claude/skills/fapiao-skill` |
| Codex CLI | `~/.codex/skills/fapiao-skill` |
| WorkBuddy | `~/.workbuddy/skills/fapiao-skill` |

其他兼容工具按各自的 Agent Skills 安装方式配置。

## 使用

将电子发票 PDF（或包含发票的 ZIP、文件夹）交给 AI，发送：

```
帮我做报销
```

AI 会先统计发票数量与金额，然后询问两项信息：

- **费用名称**：如「交通费」
- **摘要**：如「10 月市内打车」

生成的文件默认保存在发票所在目录，例如：

```
报销单-交通费-20261109.docx
报销单-交通费-20261109.pdf  （需 Windows + 本机 Word）
```

Word 首页为已填写的费用报销单，后续每页排放 2 张发票。打印建议使用 A4 纵向、100% 缩放、单面打印。

## 命令行（可选）

```bash
pip install python-docx pymupdf pillow pypdf
python scripts/pipeline.py --src 发票.zip --fee "交通费" --summary "10 月市内打车"
```

支持多个 `--src`，可传入 ZIP 或目录；通过 `--template` 指定公司模板，通过 `--per-page` 调整每页发票张数。

Windows 环境可安装 `pywin32`，借助本机 Word 校验分页并导出 PDF。其他环境仍可生成 DOCX，但需人工检查分页，且不保证生成 PDF。

## 注意事项

- 仅支持电子发票 PDF，不支持照片、截图；加密压缩包请先解压。
- 金额以发票数据为准，发现不一致时提示核对，不自动猜测。
- 脚本在本机处理文件，不包含上传发票的功能；如使用云端 AI，数据处理方式以所用服务为准。

## License

MIT
