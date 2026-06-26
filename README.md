# Bar Racing Image Custom Visual

An animated bar racing chart for Microsoft Power BI with support for category images supplied from report data.

## Key Features

- Animated bar race across a Play Axis field.
- Category image support from `data:image/...` values.
- Power BI selection and cross-filtering support.
- Native Power BI tooltip data plus optional image preview.
- Formatting options for animation speed, frame delay, autoplay, looping, labels, axes, gridlines, bar styling, and reduced motion.
- No external service calls or external image downloads.

## Data Roles

| Data Role | Type | Description |
| --- | --- | --- |
| Category | Grouping | The ranked category name. |
| Image URI | Grouping | Optional image data for the category. Use `data:image/...` values. |
| Measure | Measure | Numeric value used for bar length and ranking. |
| Play Axis | Grouping | Frame, period, or sequence used for animation. |

## How to Use

1. Add the visual to a Power BI report.
2. Add fields to Category, Measure, and Play Axis.
3. Optionally add a `data:image/...` field to Image URI.
4. Use the play, reset, step, and scrubber controls to move through frames.
5. Tune animation and label settings in the formatting pane.

## Certification and Privacy

This visual is designed for Microsoft Power BI certification:

- `capabilities.json` declares no external privileges.
- The visual does not use `fetch`, `XMLHttpRequest`, WebSocket, or external HTTP requests.
- The visual does not load external image URLs.
- User and report data are rendered locally inside Power BI.

See [PRIVACY.md](PRIVACY.md) and [SUPPORT.md](SUPPORT.md).

## Development

Install dependencies:

```powershell
npm install
```

Run the visual:

```powershell
npm start
```

Run lint:

```powershell
npm run eslint
```

Build a certification-audited package:

```powershell
npm run package
```

## Repository

https://github.com/SRathinaGiri/BarRacingImage

## License

MIT. See [LICENSE](LICENSE).
