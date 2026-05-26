# 阿里司法拍卖 - 昆明市法拍房数据爬取

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 运行
python sf_taobao_scraper.py
```

## 获取 Cookie

1. 用 Chrome 打开 https://sf.taobao.com/ 并登录淘宝账号
2. 按 F12 打开开发者工具
3. 切到 Network 标签，刷新页面
4. 点击任意请求 → Headers → Cookie
5. 复制整行 Cookie 值

## 运行选项

- 模式 1（推荐）：Selenium 模拟浏览器，字段最全
- 模式 2（轻量）：纯 requests，部分字段可能缺失（详情页是 React 渲染）

## 注意事项

- Cookie 有效期有限，爬取中断请重新获取
- 请求间隔 2~5 秒，数据量大时需要较长时间
- Ctrl+C 中断也会保存已爬取的数据
