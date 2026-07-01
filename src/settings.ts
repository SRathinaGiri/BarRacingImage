"use strict";

import { dataViewObjectsParser } from "powerbi-visuals-utils-dataviewutils";
import DataViewObjectsParser = dataViewObjectsParser.DataViewObjectsParser;

export class VisualSettings extends DataViewObjectsParser {
  // Create a new "animation" settings card
  public animation: AnimationSettings = new AnimationSettings();
  public labels: LabelSettings = new LabelSettings();
  public axes: AxisSettings = new AxisSettings();
  public bars: BarSettings = new BarSettings();
}

export class AnimationSettings {
  // This is the property for our duration
  // We give it a default value of 1000
  public duration: number = 1000;
  public frameDelay: number = 100;
  public easing: string = "linear";
  public autoplay: boolean = true;
  public loop: boolean = true;
  public showControls: boolean = true;
  public maxBars: number = 10;
  public showPlayAxisLabel: boolean = true;
  public playAxisLabelPosition: string = "top";
  public reduceMotion: boolean = false;
}

export class LabelSettings {
  public show: boolean = true;
  public fontSize: number = 12;
  public color: string = "#333333";
  public displayUnits: number = 1; // 0 = Auto, 1 = None, 1000, 1000000, 1000000000
  public precision: number = 0;
  public categoryOnBars: boolean = false;
  public categoryFontFamily: string = "Segoe UI";
  public categoryFontSize: number = 12;
  public showCategoryWithImage: boolean = true;
  public showImageInTooltip: boolean = true;
  public iconOutline: boolean = false;
  public imagePadding: number = 2;
  public imageInsideEnd: boolean = true;
  // true = inside end, false = outside end
  public labelsInside: boolean = false;
}

export class AxisSettings {
  public showXAxisLabels: boolean = true;
  public showYAxisLabels: boolean = true;
  public yAxisFontSize: number = 12;
  public showGridlines: boolean = true;
  public showXAxisTitle: boolean = false;
  public showYAxisTitle: boolean = false;
}

export class BarSettings {
  public cornerRadius: number = 4;
  public colorMode: string = "auto";
}
