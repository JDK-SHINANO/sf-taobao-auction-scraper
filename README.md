# 阿里司法拍卖 - 法拍房数据爬取工具

针对淘宝司法拍卖平台（sf.taobao.com）的房屋拍卖数据爬取与整理工具，用于资产评估课程作业中收集法拍房成交案例数据。

## 文件说明

| 文件 | 说明 |
|---|---|
| sf_taobao_scraper.py | 主爬虫：列表页 + 详情页，断点续传，反爬拦截 |
| merge_auction.py | 合并器：把每个标的物的 txt/csv 合并成一个文件 |
| 合并器.bat | 合并器的 Windows 启动脚本（双击运行） |
| requirements.txt | Python 依赖清单 |
| crawl_config.example.json | 配置文件模板（复制为 crawl_config.json 后使用） |

## 运行环境

- Windows 10/11，安装 Edge 或 Chrome 浏览器
- Python 3.8+，安装依赖：`pip install -r requirements.txt`
- 首次运行会自动下载对应版本的 WebDriver（EdgeDriver / ChromeDriver）

## 使用方法

1. 复制 `crawl_config.example.json` 为 `crawl_config.json`
2. 运行 `python sf_taobao_scraper.py`
3. 按菜单选择城市、房屋类型、区县、资产来源、日期范围
4. 浏览器会自动打开，手动登录淘宝账号（Cookie 自动保存，下次免登录）
5. 遇到滑块验证时，在浏览器中手动完成，回到控制台选 [1] 继续

爬取结果保存在 `output/日期范围/` 下，每个标的物一个目录，含：

- 拍卖信息.txt（成交价、评估价、状态等摘要）
- 基本信息.csv（表A：标的物规则信息）
- 标的物介绍.csv（表B：标的物情况调查表）
- 附件/（评估报告等文件）

全部完成后运行合并器，生成统一的 `merged_output.txt`。

## 断点续传

任务中断后重新运行，主菜单会出现「恢复未完成任务」，自动跳过已处理的 URL。进度保存在 `crawl_config.json`。

## 隐私与合规

- `crawl_config.json` 包含淘宝账号 Cookie，属个人隐私，已被 .gitignore 排除，请勿上传
- `output/` 为爬取数据，同样被排除
- 仅用于学习研究，请遵守平台条款，控制请求频率
