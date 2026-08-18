# Japan gaishikei CV pack — Herman Mok

Review of `Herman_-_resume__new__29e0.pdf` for **large foreign-capital / non-Japanese-culture** employers in Tokyo.

| File | Use |
|---|---|
| `REVIEW.md` | Cantonese review of the uploaded CV |
| `Herman-Mok-Resume-EN.pdf` | Send this after filling visa + LinkedIn URL check + metrics |
| `Herman-Mok-Resume-EN.html` | Edit, then Chrome print to A4 PDF |
| `Herman-Mok-Resume-EN.md` | Plain-text source |

Regenerate PDF:

```bash
timeout 25 google-chrome --headless --disable-gpu --no-sandbox \
  --disable-dev-shm-usage --user-data-dir=/tmp/chrome-cv \
  --no-pdf-header-footer \
  --print-to-pdf=japan-gaishi-cv/Herman-Mok-Resume-EN.pdf \
  file://$PWD/japan-gaishi-cv/Herman-Mok-Resume-EN.html
```

This GitHub repo is **public**. Phone and email are on the source CV; consider that before merging.
