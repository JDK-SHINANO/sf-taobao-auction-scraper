# Alibaba Judicial Auction — Property Auction Scraper

[简体中文](README.md)

Collects and organises housing auction data from the Taobao judicial auction platform (sf.taobao.com). Selenium drives a browser past the slider check, requests pulls the detail pages, the crawl resumes after an interruption, and a merge step combines the per-lot txt/csv files into one document.

It began as coursework for an asset valuation class, collecting judicial auction property sales as case material.

## What it is

The script runs in two passes. Selenium opens the listing page; you log in by hand, then filter by city, property type, asset source and date range to collect lot URLs. Each lot is then opened in turn to pull the auction details and the two survey tables. An interrupted run resumes where it stopped — progress is written to `crawl_config.json`.

Output is layered as `output/date-range/lot/`: one summary txt per lot, two csv files (Table A for lot rules, Table B for the survey), and attachments stored separately. When the crawl finishes, the merger produces a single `merged_output.txt`.

## How to use

Tested on Windows 10/11 with Edge or Chrome and Python 3.8+.

```bash
pip install -r requirements.txt
```

1. Copy `assets/crawl_config.example.json` to `crawl_config.json` in the repository root.
2. Run `python src/sf_taobao_scraper.py`.
3. Pick city, property type, district, asset source and date range from the menu.
4. The browser opens; log in to your Taobao account by hand. Cookies are saved, so later runs skip the login.
5. When the slider check appears, solve it in the browser and choose [1] in the console to continue.

The first run downloads the matching WebDriver version (EdgeDriver or ChromeDriver) on its own.

Results land under `output/date-range/`, one directory per lot:

- 拍卖信息.txt — summary: hammer price, appraised value, status
- 基本信息.csv — Table A: lot rules
- 标的物介绍.csv — Table B: lot survey
- 附件/ — valuation reports and other attachments

When everything is done, run `src/合并器.bat` (or `python src/merge_auction.py`) to produce the combined `merged_output.txt`.

### Resume

Run the script again after an interruption and the main menu offers 「恢复未完成任务」 (resume unfinished task). URLs already processed are skipped. State lives in `crawl_config.json`.

### Extension guide

Every configurable value sits in the constant block at the top of `src/sf_taobao_scraper.py`. Extending the scraper does not require touching the business logic.

### Adding a city

Cities are configured in `CITY_CONFIG`, three fields each:

```python
"昆明": {
    "gbk": "%C0%A5%C3%F7",          # city name in GBK, each byte as %XX
    "code": "530100",                # first 6 digits of the administrative division code
    "districts": {
        "全市": None,                # None means the location_code parameter is omitted
        "五华区": "530102",          # district administrative division code
        ...
    }
}
```

- `gbk`: the city name encoded as GBK, each byte written as `%XX`. To generate it, in any Python environment:
  `python -c "print(''.join('%%%02X' % b for b in '昆明'.encode('gbk')))"`
- `code`: administrative division code. Note that the current code **never reads this field** — only `gbk` and `districts` go into the URL. It is kept for readability; a maintainer can delete it or wire it up.
- `districts`: values are district administrative division codes, passed as the `location_code` parameter. Codes can be looked up in the division code tables published by the Ministry of Civil Affairs.

Add an entry and the new city shows up in the startup menu. No other code changes.

### Adding a property type

Values in `CATEGORY_CONFIG` are category IDs from the Taobao judicial auction site. To find one, switch to the target category on the sf.taobao.com listing page and read the address bar: the number in the `/list/` path segment is the category ID (it matches the `{category_id}__2___{gbk_city}.htm` form built in `build_list_url`).

### Adding an asset source

Values in `BIZ_TYPE_CONFIG` are settings for the `item_biz_type` parameter (criminal-case assets 4, litigation assets 6, bankruptcy assets 8, self-disposal 10). If the site changes what these mean, capture the listing page request and compare to confirm the new values. The menu uses Chinese names; `build_list_url` maps them to the numeric codes, and code -1 (any asset type) omits the parameter entirely.

### Coping with a site redesign

Taobao's front-end CSS class names carry hash suffixes (e.g. `firstDesc--G_zP45iR`) and break on redesign. Every page selector lives in the `SEL_*` constant block at the top of the file — slider markers, hammer price, Table A and B, status strings. After a redesign, check that block against the new page first.

### Other parameters

- `ITEMS_PER_PAGE`: items per page (40). A listing shorter than this is treated as the last page.
- `MAX_PAGES`: page limit for the listing (150).
- `MIN_DELAY` / `MAX_DELAY`: seconds between requests. Too short triggers the anti-bot checks.
- `SUCCESS_KEYWORDS` / `FAIL_KEYWORDS`: keywords that decide sold versus failed auction.
- `DEFAULT_CONFIG`: default date range and similar. The default date prompt in `run_task_selection` has to be changed alongside it.

## Layout

```
.
├── README.md
├── README.en.md
├── LICENSE
├── requirements.txt
├── src/
│   ├── sf_taobao_scraper.py       main scraper: listing + detail pages, resume, anti-bot handling
│   ├── merge_auction.py           merger: combines each lot's txt/csv into one document
│   └── 合并器.bat                  Windows launcher for the merger (double-click)
└── assets/
    └── crawl_config.example.json  config template; copy to crawl_config.json in the repo root
```

`crawl_config.json`, `output/`, `urls.txt` and the downloaded WebDriver all sit in the repository root and are excluded by `.gitignore`.

## Requirements

Four entries in `requirements.txt`: selenium ≥4.0, requests ≥2.25, beautifulsoup4 ≥4.9, lxml ≥4.6. The script also uses tkinter from the Python standard library, for the merger's directory picker.

## Scope

### Known limitations

- The slider check cannot be passed automatically; it needs a human.
- Detail pages depend on the site's DOM, so a redesign means updating selectors per "Coping with a site redesign".
- Attachment downloads that need a logged-in session rely on forwarding the browser cookie. Some links still fail.

### Maintenance status

The course is over; maintenance stopped in June 2026. The code is released under GPL-3.0 and forks are welcome. The extension guide under "How to use" is what accumulated while the scraper was in use, kept for whoever picks it up.

### Privacy and compliance

- `crawl_config.json` holds Taobao account cookies. That is personal data. It is excluded by `.gitignore` — do not upload it.
- `output/` is scraped data and is excluded as well.
- For study and research only. Follow the platform's terms and keep the request rate down.

## License

GPL-3.0. See [LICENSE](LICENSE). Use, modify and redistribute freely; derivative works must stay under the same license and keep the copyright notice.
