# 阿里司法拍卖 - 法拍房数据爬取工具

针对淘宝司法拍卖平台（sf.taobao.com）的房屋拍卖数据爬取与整理工具，用于资产评估课程作业中收集法拍房成交案例数据。

## 维护状态

本工具为课程作业用途，作者课程已结束，2026 年 6 月起不再维护。代码以 GPL-3.0 发布，欢迎 fork 后自行维护。下方「扩展指南」记录了当时积累的维护要点，供接手者参考。

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

## 扩展指南

所有可配置项集中在 `sf_taobao_scraper.py` 顶部常量区，扩展时无需改动业务逻辑。

### 新增城市

城市配置在 `CITY_CONFIG`，每个城市有三个字段：

```python
"昆明": {
    "gbk": "%C0%A5%C3%F7",          # 城市名的 GBK 编码，每字节写成 %XX
    "code": "530100",                # 行政区划代码前 6 位
    "districts": {
        "全市": None,                # None 表示不附加 location_code 参数
        "五华区": "530102",          # 区县行政区划代码
        ...
    }
}
```

- `gbk`：城市名按 GBK 编码，每字节写为 `%XX`。生成方法（任意 Python 环境）：
  `python -c "print(''.join('%%%02X' % b for b in '昆明'.encode('gbk')))"`
- `code`：行政区划代码。注意：当前代码逻辑中**未读取该字段**（仅 `gbk` 与 `districts` 实际参与 URL 构建），保留仅为可读性，接手者可删除或启用。
- `districts`：值为区县行政区划代码，作为接口参数 `location_code` 使用。代码可在民政部发布的行政区划代码表查询。

添加新条目后，启动菜单会自动出现新城市，无需改其他代码。

### 新增房屋类型

`CATEGORY_CONFIG` 的值为淘宝司法拍卖站内的分类 ID。获取方法：在 sf.taobao.com 列表页切换目标分类，观察地址栏 URL，`/list/` 路径段的数字即为该分类 ID（`build_list_url` 中 `{category_id}__2___{gbk_city}.htm` 的格式与之对应）。

### 新增资产来源

`BIZ_TYPE_CONFIG` 的值为接口参数 `item_biz_type` 的取值（涉刑资产 4、诉讼资产 6、破产资产 8、自行处置 10）。若站点调整参数含义，需要抓包比对列表页请求确认新值；界面菜单使用中文名，`build_list_url` 会自动映射为数字编码，编码为 -1（不限资产类型）时不附加该参数。

### 页面改版应对

淘宝前端 CSS 类名带哈希后缀（如 `firstDesc--G_zP45iR`），站点改版后可能失效。全部页面选择器集中在文件顶部的 `SEL_*` 常量区（滑块特征、成交价、表A表B、状态文字等），改版后首先对照新页面检查这一区域。

### 其他参数

- `ITEMS_PER_PAGE`：每页条数（40），列表不足该数时视为末页
- `MAX_PAGES`：列表页翻页上限（150）
- `MIN_DELAY` / `MAX_DELAY`：请求间隔秒数，过短易触发反爬
- `SUCCESS_KEYWORDS` / `FAIL_KEYWORDS`：成交/流拍判定关键词
- `DEFAULT_CONFIG`：默认日期范围等，`run_task_selection` 中的默认日期提示需同步修改

## 已知局限

- 滑块验证无法自动通过，需人工介入
- 详情页依赖站点 DOM 结构，改版后需按「页面改版应对」更新选择器
- 附件下载对需登录态的链接依赖浏览器 Cookie 转发，部分链接可能仍失败
