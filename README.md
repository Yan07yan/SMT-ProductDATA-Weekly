# ae-weekly-action-report · 速卖通周度行动表生成器

一个 **WorkBuddy 本地 skill**：把每周从速卖通后台导出的数据，自动算成对标公司 KPI 的《周度行动表》（4 页 Excel），告诉你这周该先补哪个缺口、哪些品改主图/详情/冲热销。

> 适合：速卖通多店（POP + 半托管/全托管）运营者，每周一花 1 分钟导出、1 秒出表。

---

## 功能

| KPI | 来源 | 说明 |
|---|---|---|
| 销售额 | 订单文件（POP 月化）+ 结算（全托） | 月化估算，建议改导月至今 |
| 利润率 | 订单供货价（扣 3% 佣金） | 商品级毛利上限，分店单列 |
| 刊登数量 | 每日上品台账（xlsx） | 默认上一完整周汇总 |
| 热销链接 | 四店+全托近30天支付≥15 | 四店+全托近30天支付件数 ≥15 的品数 |
| 全托管订单 | 全托管近30天支付件数 | 全托管近30天支付件数合计 |
| 滞销清理 | 全量目录近30天0销量 | 依赖商品总数据落盘 |
| 上品通过率 | 列表导出 ID 差集 | 连续两周才出真实值 |

输出 `速卖通_周度行动表_YYYY-MM-DD.xlsx`：**页1 公司KPI与缺口 / 页2 本周行动清单 / 页3 数据底稿 / 页4 上周好品类与本周方向**。

---

## 安装（本地部署）

方式 A — 作为 WorkBuddy 用户级 skill（推荐，跨项目可用）：
```bash
# 把本仓库放到 WorkBuddy 的 skills 目录
git clone <本仓库url> ~/.workbuddy/skills/ae-weekly-action-report
```
或手动把 `ae-weekly-action-report/` 整个文件夹复制进 `~/.workbuddy/skills/`。

方式 B — 项目级 skill（团队共享同一仓库）：
```bash
git clone <本仓库url> <你的项目>/.workbuddy/skills/ae-weekly-action-report
```

依赖（仅首次）：
```bash
python -m pip install openpyxl xlrd
```
> 用 WorkBuddy 托管的 Python 即可（已含上述包），见下方"使用"。

---

## 使用

1. 每周一按 `references/每周导出勾选检查表.md` 导出文件，全部丢进一个目录（如 `~/Desktop/数据表格/每周导出`）。
2. 把 `assets/刊登数量追踪_模板.xlsx` 复制一份改成 `刊登数量追踪.xlsx`，**每天填 5 店上架数**（CLYT1=全托管，其余=POP）。
3. 在 WorkBuddy 里说 **「跑周报」**，或直接跑脚本：

```bash
"<python路径>/python.exe" "<skill目录>/scripts/build_weekly_action.py" \
  --data "<每周导出目录>" \
  --date 2026-09-14 --week-start 2026-09-08 --week-end 2026-09-12
```

参数：

| 参数 | 默认 | 说明 |
|---|---|---|
| `--data` | `./每周导出` | 每周导出目录 |
| `--date` | 今天 | 报告日期 `YYYY-MM-DD` |
| `--week-start` / `--week-end` | 上一完整周(一~日) | 刊登台账汇总窗口 |
| `--out` | `<data>/速卖通_周度行动表_<date>.xlsx` | 输出路径 |
| `--snapshot` | `<skill>/scripts/_listing快照.json` | ID 快照（**勿删，存在脚本所在目录**） |
| `--track` | `<data>/刊登数量追踪.xlsx` | 上品台账（缺失自动建模板） |

---

## 接入自己的账号（必改）

脚本**不含任何真实店铺 ID**，账号相关全部放 `config.json`，同事 clone 后只改配置、不动代码：

1. 把仓库里的 `config.example.json` 复制为 `config.json`；
2. 填入你自己的：
   - `saas_shop`：订单文件里的渠道名 → 店铺代号映射
   - `shop_file` / `cat_file` / `cat_week_file`：店铺代号 → 导出文件名
   - `kpi`：本公司当月月度考核目标
3. 运行即可。脚本加载顺序：`--config` → `<skill根>/config.json` → `config.example.json` → 内置占位。

> `config.json` 已在 `.gitignore` 排除，不会被提交；`config.example.json` 是公开模板（仅占位符，无真实 ID）。
> 列名取自速卖通后台标准导出，跨账号通用；**店铺代号和 KPI 是账号/月度专属**。

---

## 同事如何使用（快速上手）

1. 克隆到 WorkBuddy 的 skills 目录：
   ```bash
   git clone <本仓库url> ~/.workbuddy/skills/ae-weekly-action-report
   ```
2. 进入目录，**复制 `config.example.json` 为 `config.json`**：
   ```bash
   cp ~/.workbuddy/skills/ae-weekly-action-report/config.example.json \
      ~/.workbuddy/skills/ae-weekly-action-report/config.json
   ```
3. 打开 `config.json`，改里面的 **店铺代号**（`saas_shop` / `shop_file` / `cat_file` / `cat_week_file`）和 **当月 KPI 目标**（`kpi`），保存。
4. 放到 `~/.workbuddy/skills/` 下即生效：WorkBuddy 下次说「跑周报」会自动读取该 skill 与你的 `config.json`，**无需改任何代码**。

> `config.json` 由 `.gitignore` 排除，不会进仓库；你本地填写的真实店铺信息只留在本机磁盘。

---

## 目录结构

```
ae-weekly-action-report/
├── SKILL.md                          # skill 元数据 + 使用说明（WorkBuddy 读取）
├── README.md                         # 本文件（GitHub 用）
├── .gitignore                        # 排除个人数据/快照/配置（见「安全/隐私」）
├── config.example.json               # 公开配置模板（占位符，无真实 ID）
├── scripts/
│   └── build_weekly_action.py        # 主脚本（导出驱动周报生成，代码不含真实 ID）

# 以下文件仅存在你本地，已被 .gitignore 排除、不在仓库中：
#   config.json、scripts/_listing快照.json、刊登数量追踪.xlsx
├── references/
│   └── 每周导出勾选检查表.md           # 每周导出清单（打勾用）
└── assets/
    └── 刊登数量追踪_模板.xlsx          # 上品台账模板
```

---

## 已知边界

- 销售额 POP 窗口仅几天时月化为估算，建议改导「月至今/自然月」直接对标。
- 「商品总数据」必须 Excel 打开→保存一次再丢进目录，否则空壳、滞销清算暂停。
- 通过率首次运行为"基线周"，真实率下周出。
- `_listing快照.json` 跨周持久化，不要删。⚠️ **它存在脚本所在目录（`<skill>/scripts/_listing快照.json`），不是 `--data` 目录**——所以你每周刷新导出文件夹不会误删它；clone / 打包时记得别带上它（已被 `.gitignore` 排除）。

---

## License

内部工具，仅供团队内部使用。未经授权请勿外传。

---

## 安全 / 隐私（公开仓库前必读）

- 本仓库**不含任何 API key / token / cookie / 密码**，脚本只读取你本地的 Excel 导出文件，不发起任何网络请求。
- `config.json`（含你的真实店铺代号与订单渠道名）已被 `.gitignore` 排除，**不会上传**；公开模板是 `config.example.json`（仅占位符，无任何真实 ID）。
- `_listing快照.json` 含你全部商品 ID，已被 `.gitignore` 排除，**绝对不要提交**。提交前用 `git check-ignore _listing快照.json` 确认能命中。
- 若公司不允许公开内部店铺结构，请建 **Private** 仓库，或保持 `config.example.json` 的占位符、不要填入真实 ID。

---

## 直接安装方法
```bash
请帮我把一个 WorkBuddy skill 从 GitHub 安装到本地，步骤如下：

1. 仓库地址是 Private 仓库，你需要先确认我的 GitHub 账号已被加为 collaborator。
   仓库地址：https://github.com/Yan07yan/SMT-ProductDATA-Weekly.git

2. 用 git clone 把它克隆到 WorkBuddy 的 skills 目录：
   git clone https://github.com/Yan07yan/SMT-ProductDATA-Weekly.git ~/.workbuddy/skills/ae-weekly-action-report

3. 进入该目录，把 config.example.json 复制一份为 config.json：
   cp ~/.workbuddy/skills/ae-weekly-action-report/config.example.json \
      ~/.workbuddy/skills/ae-weekly-action-report/config.json

4. 打开 config.json，把里面的店铺代号和 KPI 改成我自己的（我会另外告诉你具体值）。

5. 装好后确认：
   - 目录 ~/.workbuddy/skills/ae-weekly-action-report/ 存在
   - 里面有 SKILL.md、README.md、scripts/build_weekly_action.py、config.json
   - 告诉我下一步怎么用
```

---
