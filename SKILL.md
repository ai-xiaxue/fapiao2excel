---
name: fapiao2excel
description: 批量把一个文件夹里的发票/票据（PDF 或图片）识别成结构化 Excel 表。电子发票 PDF 走文字层精确解析（0 token、金额自校验、不用 key），图片/扫描件才用视觉大模型兜底。当用户需要整理、归档、报销一批发票、把票据转成表格时使用。
---

# fapiao2excel（发票批量转 Excel skill）

把一整个文件夹的发票/票据，一条命令识别成一张结构化 Excel。

## 何时用这个 skill
- 用户有「一堆发票/票据/报销单」要整理成表格 / Excel
- 用户抱怨「一张张抄发票」「报销季整理票据烦」
- 用户担心「AI 读发票会把金额读错」——本 skill 电子发票走精确解析，金额自校验，不靠 AI 猜

## 怎么用
```bash
pip install -r requirements.txt
# 纯电子发票 PDF：无需 key，直接跑
python src/extract.py <发票文件夹> -o out/票据明细.xlsx
# 有图片/扫描件：复制 .env.example 为 .env，填 API_KEY/BASE_URL/VL_MODEL（任意 OpenAI 兼容视觉模型）
```

## 工作方式（分层，防幻觉）
1. 电子发票 PDF → 读 PDF 文字层，正则精确解析（发票号码/日期/金额），0 token、0 幻觉
2. 金额自校验：`金额 + 税额 = 价税合计` 对得上才采信，否则回退 AI
3. 图片/扫描件 → 视觉大模型兜底；输出「识别方式」列标明每行来源

## 输出字段
发票类型 / 发票代码 / 发票号码 / 开票日期 / 购买方名称 / 销售方名称 / 金额(不含税) / 税额 / 价税合计 / 识别方式

详见 README.md。
