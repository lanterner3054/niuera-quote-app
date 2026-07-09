# data/ 目录说明

真实业务数据不入库(.gitignore 已排除)。首次部署时:

```bash
cp data/config.example.json  data/config.json     # 然后改成你的公司信息、银行账户、编号前缀
cp data/products.example.json data/products.json  # 然后在界面「产品管理」里维护
```

| 文件/目录 | 作用 | 是否入库 |
|---|---|---|
| `config.json` | 公司信息、默认条款、银行账户、报价人、编号池 | ❌(有 example) |
| `products.json` | 产品库(含价格) | ❌(有 example) |
| `documents/` | 单据存档(每单一个 JSON + index.json 索引) | ❌ 自动生成 |
| `images/` | 产品图片 | ❌ |
| `logo.jpeg` | 公司 logo,显示在单据左上角 | ❌ 自备 |
| `stamp.png` | 电子章(透明背景 PNG),PI 可选加盖 | ❌ 自备 |
| `backups/` | config/products 每次保存前的自动备份 | ❌ 自动生成 |
