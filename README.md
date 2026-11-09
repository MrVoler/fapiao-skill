# fapiao-skill　发票 → 报销材料

把一堆电子发票压缩包，变成一份可以直接打印的报销材料：

```
第 1 页  费用报销单（自动填好费用名称、摘要、金额分格、大小写、附单据张数）
第 2 页  发票 1、发票 2
第 3 页  发票 3、发票 4
...
```

这是一个 [WorkBuddy](https://www.workbuddy.cn) / Claude Code 风格的 Skill：一份 `SKILL.md` 加七个可独立执行的 Python 脚本。

## 它解决什么

手工做这件事的麻烦在于：

- 压缩包里发票和**行程单**混在一起，文件名只差两个字，很容易把行程单当发票打印
- 电子发票 PDF 是内嵌字体，`extract_text()` 取不到中文，**没法直接读金额**
- 报销单的金额是**一位一格的方格**（千万…万…元角分），填错一位就作废
- 报销单要写金额**大写**，手写容易错
- Word 里排版「每页两张」时，行高不固定就会串页、裁切、多出空白页

这套脚本把这五件事都固定下来了。

## 安装

```bash
git clone https://github.com/<you>/fapiao-skill.git
cp -r fapiao-skill ~/.workbuddy/skills/fapiao-skill
```

依赖：

```bash
pip install python-docx pymupdf pillow pypdf pywin32
```

`pywin32` + 本机 Word 仅用于实测页数；没有时会跳过页数校验并提示人工确认。

## 快速开始

```bash
cd fapiao-skill/scripts

python pipeline.py \
  --work ./work \
  --src 发票1.zip 发票2.zip \
  --fee "交通费" \
  --summary "10 月市内打车" \
  --out ./work/报销材料.docx
```

输出的 `work/` 里会有：

```
work/
├── 发票PDF/            真正的发票
├── 行程单PDF/          行程明细（默认不进最终文档）
├── 结构化数据/         数电票 XML，金额的权威来源
├── 图片/               300 DPI 裁边后的发票图
├── amounts.json        金额清单（各步骤之间的接口）
├── 发票金额汇总.csv    报销贴单直接用
├── 报销单.docx
├── 发票排版.docx
├── 报销材料.docx       ← 最终交付
└── 预览/p1.png…        逐页预览图
```

也可以分步执行，见 `SKILL.md`。

## 脚本

| 脚本 | 作用 | 关键产出 |
|---|---|---|
| `01_collect.py` | 解压、修正 GBK 文件名、发票/行程单分离、去重 | `collect.json` |
| `02_amounts.py` | 从 XML 取金额，与文件名交叉校验，发票号去重 | `amounts.json`、汇总 CSV |
| `03_render.py` | PDF → 300 DPI PNG，裁白边 | `图片/NN.png` |
| `04_build_invoice_docx.py` | 每页 A4 两张，行高固定 + cantSplit | `发票排版.docx` |
| `05_fill_form.py` | 填报销单：分格、大小写、附单据张数 | `报销单.docx` |
| `06_merge.py` | 表格深拷贝 + 图片关系重绑 | `报销材料.docx` |
| `07_verify.py` | Word 实测页数 + 逐页图片数 + 导出 PDF | `*.pdf`、`预览/` |
| `pipeline.py` | 一条龙串起 01→07 | 同上 |

## 设计上的几个决定

**金额以票据为准。** 数电票 XML 里的 `TotalTax-includedAmount` 是权威值；文件名上的金额只用来交叉校验，不一致就报错，不静默选一个。

**费用名称和摘要不给默认值。** `pipeline.py` 的 `--fee` / `--summary` 是必填项。报销口径必须由人确认——填错了要重做整份材料。

**报销人、经理审批永远留空。** 这两栏是给手写的。

**合并用深拷贝而不是 Word COM。** 目标文档首块是表格时，`Selection.HomeKey + InsertFile` 会把插入点放进单元格，整张表单被塞进发票表格里。改用「表格深拷贝 + 逐个重绑 `a:blip` 的 `r:embed`」，与图片顺序完全解耦。

**必须实测页数。** 公式算出 5 页不算数，要用 Word 的 `ComputeStatistics(2)` 拿到真实分页，再逐页确认每页 2 张。

## 换自己的报销单模板

`assets/费用报销单模板.docx` 是一个通用的两联空白表。换成公司模板只要满足：

1. 表头含「费用名称」「摘要」，金额分格表头含「分」「角」「元」
2. 有「总计金额（小写）」「总计金额（大写）」行

细节见 `references/template-structure.md`。填单脚本按表头文字定位，不写死行列号。

## 脱敏

仓库里只有代码、空白模板和文档。**不含任何真实发票、金额、发票号、销售方或个人姓名**，示例中的数字均为虚构。运行时生成的 `work/` 目录已由 `.gitignore` 排除——里面是你的真实财务数据，不要提交。

## 已知限制

- 页数校验依赖 Windows + 本机 Word（或 PowerShell）；其他平台跳过并提示
- 金额优先取数电票 XML；只有 PDF、且 PDF 无文本层的老式票据需人工补录金额
- 报销单版式识别针对「表头文字 + 分格」结构，非常规版式需调整定位逻辑

## License

MIT
