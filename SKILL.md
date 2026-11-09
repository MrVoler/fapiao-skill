---
name: fapiao-skill
description: 把电子发票整理成可直接打印的报销材料——汇集并分类发票与行程单、从数电票 XML 核对金额、按每页 A4 两张排版、自动填写费用报销单（金额分格、大小写、附单据张数）、合成单个 Word 并导出 PDF 校验页数。当用户提到发票、报销单、报销、贴票、打车发票、行程单、费用汇总，或要求把一堆发票整理成一个 Word / PDF、算发票总额、填报销单时使用。
agent_created: true
---

# 发票 → 报销材料

把一堆电子发票压缩包变成一份「第 1 页报销单 + 之后每页两张发票」的打印件，全程可核对。

## 硬性规则

1. **先确认，后生成。** 费用名称、摘要、报销单模板三项未确认前，不得生成最终文件。用 `AskUserQuestion` 一次性问清，不要猜。
   - 费用名称：如「交通费」「差旅费」「软件订阅月费」
   - 摘要：如「10 月市内打车」「10月订阅」
   - 明细归集方式：合并 1 行 / 按平台分 N 行 / 按发票逐张分行
2. **报销人、经理审批永远留空。** 这两栏供手写，任何情况下都不填。
3. **金额以票据为准，不以文件名或用户口述为准。** 数电票 XML 的 `TotalTax-includedAmount` 是权威值；文件名里的金额只作交叉校验。两者不一致必须报告，不能静默取一个。
4. **不覆盖原模板。** 只读取，另存新文件。
5. **交付前必须实测页数。** 不能凭公式宣布完成，要用 Word 统计真实分页并逐页确认每页张数（见 Step 7）。
6. **脱敏。** 生成的 CSV/JSON/PDF 含真实发票号、销售方、金额，属于个人财务数据。不要把产物提交进任何公开仓库，不要写进示例文件。

## 流程

先做一次「开票前沟通」，再跑脚本。脚本按序号串行，每步都有可读的产物。

### Step 0　确认（阻塞点）

用 `AskUserQuestion` 问三个问题：费用名称、摘要、明细归集方式。如果发票还没看，先跑 Step 1 拿到张数和金额，再回来问——这样用户能看着数字决定怎么填。

### Step 1　汇集

```bash
python scripts/01_collect.py --src 发票.zip 发票2.zip --work ./work
```

- 修正 ZIP 内 GBK 文件名乱码
- 发票 → `发票PDF/`，行程单 → `行程单PDF/`（**行程单不是发票，必须分开**），XML → `结构化数据/`
- 跨压缩包按 (文件名, 大小) 去重
- 产物：`work/collect.json`

### Step 2　核价

```bash
python scripts/02_amounts.py --work ./work
```

- 逐张从 XML 取含税金额、不含税、税额、销售方、发票号
- XML 缺失时回退到 PDF 文本提取（部分发票内嵌字体，取不到要报警）
- 与文件名标注的金额逐张比对，差额 > 0.01 记为待确认
- 按发票号去重
- 产物：`work/amounts.json`（降序明细 + 合计）、`work/发票金额汇总.csv`（报销贴单直接用）

### Step 3　转图

```bash
python scripts/03_render.py --work ./work --dpi 300
```

300 DPI 渲染首页，自动裁白边（裁得过狠时放弃裁剪，避免切掉内容）。产物 `work/图片/NN.png`。

### Step 4　排版

```bash
python scripts/04_build_invoice_docx.py --manifest ./work/amounts.json --out ./work/发票排版.docx
```

A4 纵向、边距 12mm、行高精确固定、`cantSplit` 保证一张发票不被切两页、末尾补 1pt 空段防止 Word 顶出新页。

### Step 5　填单

```bash
python scripts/05_fill_form.py \
  --template assets/费用报销单模板.docx \
  --fee "交通费" --summary "10 月市内打车" \
  --amount 123.45 --bills 3 \
  --out ./work/报销单.docx
```

- 自动定位表头、明细行、金额分格列（末位对齐，`¥` 放在首位数字左侧）
- 金额大写自动转换（`_common.rmb_upper`）
- 两联模板只保留第一联
- 报销人 / 经理审批留空
- 末尾加分页符，让发票从第 2 页开始

### Step 6　合并

```bash
python scripts/06_merge.py --form ./work/报销单.docx --invoice ./work/发票排版.docx --out ./work/报销材料.docx
```

用「表格深拷贝 + 图片关系重绑」，不要用 Word COM 的 `InsertFile`（文档首块是表格时，插入点会落进单元格，整张表单被塞进去）。

### Step 7　校验并导出 PDF

```bash
python scripts/07_verify.py --docx ./work/报销材料.docx --per-page 2 --preview-dir ./work/预览
```

- 期望页数 = 1 + ⌈发票数 ÷ 每页张数⌉
- Word 实测页数（pywin32 驱动，无 pywin32 时回退 PowerShell）
- 导出 PDF，逐页统计图片数（第 2 页起每页必须等于 per-page）
- 生成 `预览/p1.png…`，用看图工具确认第 1 页表单无错位、后续页无裁切

### 一键版

```bash
python scripts/pipeline.py --work ./work --src 发票.zip \
  --fee "交通费" --summary "10 月市内打车" --out ./work/报销材料.docx
```

`--fee` / `--summary` 无默认值，强制走完确认环节。

## 数据契约

`amounts.json` 是各步骤之间唯一的接口：

```json
{
  "count": 2,
  "total": "123.45",
  "items": [
    {"seq": 1, "label": "示例平台A", "amount": "100.00", "net": "99.00", "tax": "1.00",
     "seller": "示例公司", "no": "0000…", "date": "2026-01-01 10:00:00",
     "file": "…pdf", "img": "…png", "img_w_mm": 202.8, "img_h_mm": 127.3}
  ]
}
```

## 交付前检查清单

- [ ] 发票张数与用户预期一致，重复票已剔除
- [ ] 逐张金额与文件名标注一致，合计 = 明细之和
- [ ] 报销单：费用名称 / 摘要 / 金额分格 / 大小写 / 附单据张数 五项都对
- [ ] 报销人、经理审批为空
- [ ] 实测页数 = 1 + ⌈n÷2⌉，第 2 页起每页 2 张
- [ ] 逐页看过：无裁切、无串页、两张之间有间隔
- [ ] 提醒用户：不要双面打印

## 依赖

```bash
pip install python-docx pymupdf pillow pypdf pywin32
```

Windows + 本机 Word 才能实测分页；macOS/Linux 下 Step 7 会跳过页数校验并提示人工确认。

## 参考文件

- `references/workflow.md`　各阶段的判定标准与异常处理
- `references/template-structure.md`　报销单模板的表格结构（填单脚本依赖它）
- `references/pitfalls.md`　八个真实踩过的坑及原因
- `assets/费用报销单模板.docx`　空白模板，可换成自己公司的
