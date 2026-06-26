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
  // We give it a default value of 250
  public duration: number = 500;
  public frameDelay: number = 0;
  public easing: string = "linear";
  public autoplay: boolean = true;
  public loop: boolean = true;
  public showControls: boolean = true;
  public maxBars: number = 10;
  public reduceMotion: boolean = false;
}

export class LabelSettings {
  public show: boolean = true;
  public fontSize: number = 12;
  public color: string = "#333333";
  public displayUnits: number = 0; // 0 = Auto, 1000, 1000000, 1000000000
  public precision: number = 0;
  public categoryOnBars: boolean = false;
  public showCategoryWithImage: boolean = true;
  public showImageInTooltip: boolean = true;
  public iconOutline: boolean = false;
  public imageInsideEnd: boolean = false;
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
