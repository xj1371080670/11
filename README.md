# 美客多智利站（MercadoLibre Chile）卖家线索爬虫

为 **深圳爱骑仕智能科技** 跨境部业务开发设计。
按关键词在美客多智利站搜索 → 抓取所有商品 → 反查卖家信息 → 自动识别**疑似中国卖家** → 导出 Excel。

---

## 工具能做什么

输入一组西语关键词（例如 `intercomunicador casco moto`），脚本会：

1. 翻页抓取每个关键词的搜索结果（默认每词 5 页 ≈ 240 个商品）
2. 解析每条商品里的卖家 ID、昵称、所在地、商品销量
3. 去重得到**所有卖家清单**，再逐个访问卖家主页拿信用等级、累计销量
4. 用启发式规则打分判断**是否为中国卖家**（拼音昵称、地址含 "China/Shenzhen/HK" 等）
5. 输出两份 Excel：
   - `sellers.xlsx` — 按"中国卖家可能性 + 销量"排序的**优先跟进客户清单**
   - `listings.xlsx` — 原始商品明细（用来看人家在卖什么、价格区间）

---

## 重要前提：必须用 VPN / 海外代理

MercadoLibre 会**自动拦截中国区 IP**，所有搜索/商品页都会被跳转到 `/gz/account-verification` 验证页。
本人在深圳办公室直连一定打不开美客多智利站，必须先解决这个问题：

**三选一：**
- ✅ **方案 A（推荐，免费）**：开 VPN，节点选 **智利 / 阿根廷 / 美国**，整机走 VPN
- ✅ **方案 B（更稳）**：用付费的住宅代理（如 BrightData、IPRoyal、Smartproxy），脚本里通过环境变量 `HTTPS_PROXY` 走代理
- ✅ **方案 C（最省事）**：买一台**智利/阿根廷/美国地区的云服务器**（如 AWS São Paulo / Vultr Mexico / Linode Chile，约 $5/月），把脚本扔上去跑

---

## 安装

需要 Python 3.9+

```bash
pip install -r requirements.txt
```

## 使用

### 1. 编辑关键词

打开 `config.yaml`，按需要增减西语关键词。已经预置了爱骑仕产品对应的常用搜索词。

### 2. 跑爬虫

```bash
# 直接跑（开了系统级 VPN 时）
python mlc_scraper.py

# 或者指定代理
HTTPS_PROXY=http://user:pass@proxy.example.com:8080 python mlc_scraper.py

# 只跑前 1 页测试
python mlc_scraper.py --max-pages 1

# 只跑某一个关键词测试
python mlc_scraper.py --keyword "intercomunicador casco moto"
```

### 3. 看结果

跑完后在 `output/` 目录下会生成：
- `sellers.xlsx` ⭐ 这个是你每天打开的客户清单
- `listings.xlsx` 商品明细备查
- `raw/` 目录下保存了原始 HTML，方便排查问题

---

## sellers.xlsx 字段说明

| 字段 | 说明 |
|---|---|
| 中国卖家信号 | 命中的中国卖家特征（如 `pinyin_nickname`、`china_in_address`） |
| 优先级 | 🔥 高（强信号+销量大）/ ⭐ 中 / 普通 |
| 卖家昵称 | MercadoLibre 上的店铺名 |
| 累计销量 | 卖家累计成交订单（粗略反映规模） |
| 信用等级 | MercadoLíder Platinum/Gold/... |
| 主要在卖商品 | 我们抓到的 listing 标题示例 |
| 所在地 | 卖家自填的地址 |
| 注册年份 | 注册年份越新的越值得抢（老卖家可能已签竞品） |
| 店铺链接 | 直接打开店铺看主营品 |
| 触达建议 | 自动生成的下一步动作建议 |

---

## 法律与合规说明

- 本工具只抓取 MercadoLibre **公开页面**的卖家公开信息（昵称、店铺地址、信用等级），与人工浏览看到的内容完全一致，不涉及破解任何登录态或反爬措施
- 抓取频率默认 **2-4 秒/请求**，远低于人工浏览速率，不会对站方造成负担
- 抓到的卖家信息仅用于**爱骑仕跨境部 B2B 商务开发**，请勿用于其他用途

---

## 常见问题

**Q: 跑起来全是 403 / 跳转到 verification 页？**
A: VPN/代理不通，先在浏览器里打开 https://listado.mercadolibre.cl/intercomunicador-casco-moto 看能不能正常显示商品列表。能开就能跑。

**Q: 跑得很慢？**
A: 默认很保守（防被封），可以改 `config.yaml` 的 `delay_min/delay_max`。但建议保持 2 秒以上，不然会触发反爬。

**Q: 想抓墨西哥/巴西/阿根廷站？**
A: 改 `config.yaml` 的 `site` 字段（`MLM` 墨西哥 / `MLB` 巴西 / `MLA` 阿根廷 / `MLC` 智利），URL 会自动切换。注意巴西要用葡语关键词。
