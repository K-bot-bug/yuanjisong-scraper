# 猿急送爬虫

Python 异步爬虫，抓取 [猿急送](https://www.yuanjisong.com/) 兼职项目列表，自动过滤后输出 Excel。

## 功能

- **反检测抓取**：`curl_cffi` 模拟 Chrome 131 TLS/HTTP2 指纹，带代理池轮转、403 指数退避、断点续爬
- **多方案支持**：轻量版（直连）、Firecrawl 代理版
- **自动分类**：按技术栈将项目分为爬虫、前端、后端、AI、小程序等 10 个类别
- **学生项目筛选**：自动过滤预算 ≤500 元、非驻场、排除高风险关键词的项目
- **代理池**：自动从快代理/free-proxy-list 抓取并测试，失败自动拉黑，定期恢复

## 项目结构

```
.
├── scrape_lightweight.py    # 主爬虫（curl_cffi 异步版，推荐）
├── scrape_firecrawl.py      # Firecrawl 方案（需配合 Claude Code MCP 使用）
├── proxy_pool.py            # 代理池（自动抓取+测试+轮转+黑名单）
├── classify.py              # 项目分类器（按技术栈分 sheet 输出）
├── filter_student_projects.py  # 学生项目筛选器（预算≤500元）
├── requirements.txt         # 依赖
├── .env                     # 配置（本地使用，未入库）
└── 数据/                    # 输出目录（未入库）
```

## 安装

```bash
pip install -r requirements.txt
```

依赖：`curl_cffi` `beautifulsoup4` `openpyxl` `loguru`

## 使用方式

### 1. 抓取项目

```bash
# 抓取全部（默认300页，约30分钟）
python scrape_lightweight.py

# 指定页数
python scrape_lightweight.py --max-pages 50

# 断点续爬（从上次中断处继续）
python scrape_lightweight.py --resume

# 不排除关键词（调试用）
python scrape_lightweight.py --include
```

**输出文件（均在 `数据/` 目录）：**

| 文件 | 说明 |
|------|------|
| `yuanjisong_时间戳.xlsx` | 全部可投递项目 |
| `student_时间戳.xlsx` | 预算 ≤500 元的项目 |
| `python_时间戳.xlsx` | Python 相关项目 |

### 2. 分类（已有抓取结果后）

```bash
python classify.py
```

输入：`数据/python-freshman.xlsx`
输出：`数据/python-projects-v2.xlsx`（按分类分 sheet）

### 3. 筛选学生项目（已有抓取结果后）

```bash
python filter_student_projects.py
```

输入：`数据/yuanjisong_jobs.xlsx`
输出：`数据/student-friendly.xlsx`

### 4. Firecrawl 方案（IP 被封时使用）

先在 Claude Code 中调用 `firecrawl_scrape` 逐页抓取 markdown，再保存为 JSON 解析：

```bash
python scrape_firecrawl.py --input data.json --pages 10
```

## 排除关键词

爬虫内置自动过滤以下类型项目：

| 类别 | 关键词示例 |
|------|-----------|
| 高难度 | 架构、高级、专家、底层、内核、编译原理、逆向工程 |
| 算法/AI | 算法、深度学习、区块链、量化 |
| 安全相关 | 渗透测试、漏洞挖掘、密码学、安全研究、反作弊 |
| 灰产/违规 | 接码平台、翻墙、VPN代理、破解、刷单、洗钱、赌博 |
| 学术不端 | 代写论文、代做毕设 |
| 硬件/IoT | stm32、嵌入式、树莓派、单片机、物联网、FPGA |
| 游戏 | unity、game、网游、手游 |
| 光学/相机 | 相机、标定、镜头 |

## 代理池说明

| 参数 | 默认值 | 说明 |
|------|--------|------|
| 代理源 | 快代理、free-proxy-list.net | 每次初始化从两个源各抓 50 个 |
| 拉黑阈值 | 连续失败 3 次 | 达到阈值加入黑名单 |
| 恢复间隔 | 300 秒 | 黑名单代理到期后自动恢复测试 |
| 刷新间隔 | 1800 秒（30分钟） | 后台定期重新抓取代理 |

代理成功率加权随机选取，成功率高的代理被选中的概率更大。

## 注意事项

- `.env` 文件包含 Cookie 和代理配置，**已加入 `.gitignore`**，不会被上传
- `.cookies.json` 保存会话状态，供断点续爬恢复使用
- 首次运行时代理池测试需要约 10~30 秒，等待完成后即可开始抓取
- 如果 IP 被封，建议使用 Firecrawl 方案作为备选
