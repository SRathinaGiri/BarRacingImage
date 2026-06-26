# Bar Racing Image Sample

This folder contains a sample dataset for the Bar Racing Image visual.

## Files

- `BarRacingImageSampleData.csv`: monthly 2024 Wikipedia pageview counts for selected famous cartoon-character article pages.
- `images/*.svg`: original SVG badge artwork used to generate the embedded `ImageURI` values.
- `generate_sample_data.ps1`: reproducible script that refreshes the CSV from the Wikimedia Pageviews API.

The SVG badges are original sample artwork and are not official character artwork, trademarks, logos, or screenshots.

## Power BI Mapping

Import `BarRacingImageSampleData.csv` into Power BI Desktop and map fields as follows:

| Visual Field | Dataset Column |
| --- | --- |
| Category | `Category` |
| Image URI | `ImageURI` |
| Measure | `PageViews` |
| Play Axis | `Period` |

Suggested visual settings:

- Animation > Max Bars: `6`
- Animation > Frame Duration: `700`
- Animation > Loop: `On`
- Data Labels > Show Image In Tooltip: `On`
- Data Labels > Show Category When Image Used: `On`

## Regenerate Data

```powershell
powershell -ExecutionPolicy Bypass -File .\sample\generate_sample_data.ps1
```

Source: Wikimedia Pageviews API, `en.wikipedia`, all access, user traffic, monthly granularity.
