# Support Document for Bar Racing Image

Developer: S. Rathinagiri

Website: https://www.rathinagiri.in

Support Email: srg@rathinagiri.in

## 1. Introduction

Bar Racing Image is a Microsoft Power BI custom visual for animated bar race charts. It compares ranked categories across a play axis and can display category images supplied as report data.

## 2. Getting Started

To use Bar Racing Image in Power BI:

1. Import the visual from AppSource or from a `.pbiviz` package.
2. Assign fields to the data roles:
   - Category: the label or entity being ranked.
   - Image URI: an optional `data:image/...` value for the category image.
   - Measure: the numeric value used for bar length and ranking.
   - Play Axis: the frame or period used for animation.
3. Use the play, reset, step, and scrubber controls to review the animation.

## 3. Formatting

The format pane includes settings for animation timing, autoplay, loop behavior, maximum visible bars, data labels, image placement, axes, gridlines, bar corner radius, and color mode.

## 4. Troubleshooting

Visual is blank: Confirm that Category, Measure, and Play Axis fields are assigned and that the Measure field contains numeric values.

Images do not appear: Confirm that the Image URI field contains `data:image/...` values. External image URLs are intentionally not loaded.

Performance is slow: Reduce Max Bars or the number of Play Axis frames used in the report.

## 5. Technical Support

Issue Tracker: https://github.com/SRathinaGiri/BarRacingImage/issues

Direct Contact: email srg@rathinagiri.in with a description of the issue, Power BI version, sample data structure if possible, and a screenshot.

## 6. Version History

v1.0.0.0: Initial release of Bar Racing Image.
