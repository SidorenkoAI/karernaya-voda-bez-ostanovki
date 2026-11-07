import json
import sys

import pymupdf

pdf, toc, out = sys.argv[1:4]
heads = json.load(open(toc, encoding="utf-8"))
d = pymupdf.open(pdf)
# конец оглавления — первая страница, где встречается последний пункт
start = next(i for i in range(d.page_count) if d[i].search_for(heads[-1][:30])) + 1
pages, cur = {}, start
for h in heads:
    key = h[:30]
    for i in range(cur, d.page_count):
        if d[i].search_for(key):
            pages[h] = i + 1
            cur = i
            break
    else:
        print("НЕ НАЙДЕН:", h)
json.dump(pages, open(out, "w", encoding="utf-8"), ensure_ascii=False)
print("страниц:", d.page_count, "найдено:", len(pages), "из", len(heads))
