# BillsFilter

用于将支付宝 / 微信 / 中国银行账单对照 `Budget.xlsx` 基线去重，并导出可导入的 Budget 模板文件。

## 输入（只读）

- `2026-08/Budget.xlsx`：基线账单
- `BudgetImportTemplate.xlsx`：导入模板（去掉样例数据后写入结果）
- `2026-08/支付宝账单/`：支付宝账单
- `2026-08/微信账单/`：微信账单
- `2026-08/中国银行账单/`：中国银行交易流水 PDF（可多份，自动遍历）

## 输出

- `2026-08/支付宝账单/AlipayImport.csv`：过滤后的支付宝账单
- `2026-08/微信账单/WechatImport.csv`：过滤后的微信账单
- `2026-08/中国银行账单/CcbcImport.csv`：过滤后的中国银行账单

输出列参考 `BudgetImportTemplate.xlsx` 的 Records 表头：  
`分类,子类别,货币,金额,账户,记录人,日期,时间,备注`

## 去重规则

判定重复：`日期` + `账户` + `金额` 组合一致。

账户匹配：

- 银行卡：卡号后四位一致即可（忽略“储蓄卡/信用卡”等表述，以及 `&优惠` 等拼接后缀）
- 其他账户：按语义映射，例如 `零钱` ↔ `Alex的微信钱包`，`余额/余额宝` ↔ `Alex的支付宝余额`，`花呗` ↔ `Alex的蚂蚁花呗`

## 使用

```bash
pip install -r requirements.txt
python3 filter_bills.py
```
