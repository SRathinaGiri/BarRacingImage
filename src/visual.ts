/*
* PowerBI Visual CLI
*
* Copyright (c) Microsoft Corporation
* All rights reserved.
* MIT License
*
* Permission is hereby granted, free of charge, to any person obtaining a copy
* of this software and associated documentation files (the ""Software""), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
* copies of the Software, and to permit persons to whom the Software is
* furnished to do so, subject to the following conditions:
*
* The above copyright notice and this permission notice shall be included in
* all copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
* THE SOFTWARE.
*/
"use strict";

// Import Power BI specific modules
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import DataView = powerbi.DataView;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import IColorPalette = powerbi.extensibility.IColorPalette;
import ISelectionId = powerbi.visuals.ISelectionId; // We may not need this now, but it's good to have
import { VisualSettings } from "./settings";
import VisualObjectInstance = powerbi.VisualObjectInstance;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstanceEnumeration = powerbi.VisualObjectInstanceEnumeration;

// Import D3 libraries
import * as d3 from "d3";
import { createTooltipServiceWrapper, ITooltipServiceWrapper, TooltipEventArgs } from "powerbi-visuals-utils-tooltiputils";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import { formattingSettings, FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

// Import the (empty) visual.less file
import "./../style/visual.less";

/**
 * Interface for our data points.
 */
interface BarDataPoint {
    category: string;
    value: number;
    color: string;
    displayLabel?: string;
    imageUrl?: string;
    imageKey?: string;
    selectionId?: ISelectionId;
}

interface Frame {
    key: string; // play axis value (stringified)
    label: string; // display label
    points: BarDataPoint[];
}

// Formatting model (vNext)
class AnimationCardSettings extends formattingSettings.SimpleCard {
    public name: string = "animation";
    public displayName: string = "Animation";

    public duration = new formattingSettings.NumUpDown({ name: "duration", value: 500 });
    public frameDelay = new formattingSettings.NumUpDown({ name: "frameDelay", value: 100 });
    public easing = new formattingSettings.ItemDropdown({
        name: "easing",
        items: [
            { value: "linear", displayName: "Linear" },
            { value: "easeOut", displayName: "Ease Out" },
            { value: "easeInOut", displayName: "Ease In Out" }
        ],
        value: { value: "linear", displayName: "Linear" }
    });
    public autoplay = new formattingSettings.ToggleSwitch({ name: "autoplay", value: true });
    public loop = new formattingSettings.ToggleSwitch({ name: "loop", value: true });
    public showControls = new formattingSettings.ToggleSwitch({ name: "showControls", value: true });
    public maxBars = new formattingSettings.NumUpDown({ name: "maxBars", value: 10 });
    public showPlayAxisLabel = new formattingSettings.ToggleSwitch({ name: "showPlayAxisLabel", value: true });
    public playAxisLabelPosition = new formattingSettings.ItemDropdown({
        name: "playAxisLabelPosition",
        items: [
            { value: "top", displayName: "Top" },
            { value: "bottom", displayName: "Bottom" }
        ],
        value: { value: "top", displayName: "Top" }
    });
    public reduceMotion = new formattingSettings.ToggleSwitch({ name: "reduceMotion", value: false });

    public slices: formattingSettings.Slice[] = [
        this.duration,
        this.frameDelay,
        this.easing,
        this.autoplay,
        this.loop,
        this.showControls,
        this.maxBars,
        this.showPlayAxisLabel,
        this.playAxisLabelPosition,
        this.reduceMotion
    ];
}

class LabelsCardSettings extends formattingSettings.SimpleCard {
    public name: string = "labels";
    public displayName: string = "Data Labels";

    public show = new formattingSettings.ToggleSwitch({ name: "show", value: true });
    public fontSize = new formattingSettings.NumUpDown({ name: "fontSize", value: 12 });
    public color = new formattingSettings.ColorPicker({ name: "color", value: { solid: { color: "#333333" } } as any });
    public displayUnits = new formattingSettings.ItemDropdown({
        name: "displayUnits",
        value: { value: "1", displayName: "None" },
        items: [
            { value: "0", displayName: "Auto" },
            { value: "1", displayName: "None" },
            { value: "1000", displayName: "Thousands" },
            { value: "1000000", displayName: "Millions" },
            { value: "1000000000", displayName: "Billions" }
        ]
    });
    public precision = new formattingSettings.NumUpDown({ name: "precision", value: 0 });
    public categoryOnBars = new formattingSettings.ToggleSwitch({ name: "categoryOnBars", value: false });
    public categoryFontFamily = new formattingSettings.ItemDropdown({
        name: "categoryFontFamily",
        value: { value: "Segoe UI", displayName: "Segoe UI" },
        items: [
            { value: "Segoe UI", displayName: "Segoe UI" },
            { value: "Arial", displayName: "Arial" },
            { value: "Calibri", displayName: "Calibri" },
            { value: "Verdana", displayName: "Verdana" },
            { value: "Tahoma", displayName: "Tahoma" }
        ]
    });
    public categoryFontSize = new formattingSettings.NumUpDown({ name: "categoryFontSize", value: 12 });
    public showCategoryWithImage = new formattingSettings.ToggleSwitch({ name: "showCategoryWithImage", value: true });
    public showImageInTooltip = new formattingSettings.ToggleSwitch({ name: "showImageInTooltip", value: true });
    public iconOutline = new formattingSettings.ToggleSwitch({ name: "iconOutline", value: false });
    public imagePadding = new formattingSettings.NumUpDown({ name: "imagePadding", value: 2 });
    public imageInsideEnd = new formattingSettings.ToggleSwitch({ name: "imageInsideEnd", value: true });
    public labelsInside = new formattingSettings.ToggleSwitch({ name: "labelsInside", value: false });

    public slices: formattingSettings.Slice[] = [
        this.show,
        this.fontSize,
        this.color,
        this.displayUnits,
        this.precision,
        this.categoryOnBars,
        this.categoryFontFamily,
        this.categoryFontSize,
        this.showCategoryWithImage,
        this.showImageInTooltip,
        this.iconOutline,
        this.imagePadding,
        this.imageInsideEnd,
        this.labelsInside
    ];
}

class AxesCardSettings extends formattingSettings.SimpleCard {
    public name: string = "axes";
    public displayName: string = "Axes";

    public showXAxisLabels = new formattingSettings.ToggleSwitch({ name: "showXAxisLabels", value: true });
    public showYAxisLabels = new formattingSettings.ToggleSwitch({ name: "showYAxisLabels", value: true });
    public yAxisFontSize = new formattingSettings.NumUpDown({ name: "yAxisFontSize", displayName: "Axis Font Size", value: 12 });
    public showGridlines = new formattingSettings.ToggleSwitch({ name: "showGridlines", value: true });
    public showXAxisTitle = new formattingSettings.ToggleSwitch({ name: "showXAxisTitle", value: false });
    public showYAxisTitle = new formattingSettings.ToggleSwitch({ name: "showYAxisTitle", value: false });

    public slices: formattingSettings.Slice[] = [
        this.showXAxisLabels,
        this.showYAxisLabels,
        this.yAxisFontSize,
        this.showGridlines,
        this.showXAxisTitle,
        this.showYAxisTitle
    ];
}

class BarsCardSettings extends formattingSettings.SimpleCard {
    public name: string = "bars";
    public displayName: string = "Bars";

    public cornerRadius = new formattingSettings.NumUpDown({ name: "cornerRadius", value: 4 });
    public colorMode = new formattingSettings.ItemDropdown({
        name: "colorMode",
        items: [
            { value: "auto", displayName: "Auto" },
            { value: "light", displayName: "Light" },
            { value: "dark", displayName: "Dark" }
        ],
        value: { value: "auto", displayName: "Auto" }
    });

    public slices: formattingSettings.Slice[] = [
        this.cornerRadius,
        this.colorMode
    ];
}

class VisualFormattingSettingsModel extends formattingSettings.Model {
    public animation: AnimationCardSettings = new AnimationCardSettings();
    public labels: LabelsCardSettings = new LabelsCardSettings();
    public axes: AxesCardSettings = new AxesCardSettings();
    public bars: BarsCardSettings = new BarsCardSettings();
    public cards: formattingSettings.Cards[] = [this.animation, this.labels, this.axes, this.bars];
}

/**
 * Interface for our persistent color storage
 * This maps a category name (string) to a color (string)
 */
interface ColorMap {
    [categoryName: string]: string;
}

/**
 * The main visual class
 */
export class Visual implements IVisual {
    // Main D3 elements
    private host: IVisualHost;
    private eventService: IVisualEventService;
    private svg: d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private barContainer: d3.Selection<SVGGElement, unknown, null, undefined>;
    private xAxisGroup: d3.Selection<SVGGElement, unknown, null, undefined>;
    private yAxisGroup: d3.Selection<SVGGElement, unknown, null, undefined>;
    private xTitleGroup: d3.Selection<SVGGElement, unknown, null, undefined>;
    private yTitleGroup: d3.Selection<SVGGElement, unknown, null, undefined>;
    private playAxisLabelGroup: d3.Selection<SVGGElement, unknown, null, undefined>;
    
    // Color objects
    private colorPalette: IColorPalette;
    private colorMap: ColorMap; // <-- ADDED: This is our persistent color storage

    // Chart settings
    private settings: VisualSettings; // animation speed in ms
    private margin = { top: 20, right: 30, bottom: 20, left: 100 }; // margins

    // Animation state
    private frames: Frame[] = [];
    private currentFrameIndex: number = 0;
    private timer: number | undefined;

    // Simple HTML controls
    private controlsRoot: HTMLDivElement;
    private playButton: HTMLButtonElement;
    private resetButton: HTMLButtonElement;
    private stepButton: HTMLButtonElement;
    private frameLabel: HTMLDivElement;
    private progressSlider: HTMLInputElement;
    private tooltipDiv: HTMLDivElement;
    private rootElement: HTMLElement;
    private landingPage: HTMLDivElement | null = null;

    // Selection + tooltip
    private selectionManager: powerbi.extensibility.ISelectionManager;
    private tooltipServiceWrapper: ITooltipServiceWrapper;
    private formattingSettingsService: FormattingSettingsService;
    private formattingModel: VisualFormattingSettingsModel;
    private imageMap: { [key: string]: string };
    // Controls metrics
    private controlsHeight: number = 0;
    private measureFormat: string | undefined;
    private xTitleText: string = "";
    private yTitleText: string = "";
    private layout: { mLeft: number; mTop: number; innerWidth: number; innerHeight: number } | null = null;
    private labelValueByCategory: { [key: string]: number } = {};
    private labelCounterFrameByCategory: { [key: string]: number } = {};

    /**
     * Called once when the visual is initialized
     */
    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.eventService = options.host.eventService;
        this.colorPalette = this.host.colorPalette;
        this.colorMap = {}; // <-- ADDED: Initialize our color map as an empty object
        this.imageMap = {};
        this.rootElement = options.element as HTMLElement;
        try {
            const pos = window.getComputedStyle(this.rootElement).position;
            if (!pos || pos === "static") this.rootElement.style.position = "relative";
        } catch { /* ignore */ }

        // Create the main SVG element
        this.svg = d3.select(this.rootElement)
            .append("svg")
            .classed("bar-race-chart", true)
            .attr("xmlns:xlink", "https://www.w3.org/1999/xlink");

        // Clear catcher for selection clearing
        this.svg.append("rect")
            .classed("clear-catcher", true)
            .attr("fill", "transparent")
            .style("pointer-events", "all")
            .on("click", () => this.selectionManager?.clear?.());

        // Create a container for the bars
        this.barContainer = this.svg.append("g")
            .classed("bar-container", true);

        // Create groups for the X and Y axes
        this.xAxisGroup = this.svg.append("g")
            .classed("x-axis", true);

        this.yAxisGroup = this.svg.append("g")
            .classed("y-axis", true);

        // Axis titles
        this.xTitleGroup = this.svg.append("g").classed("x-title", true);
        this.yTitleGroup = this.svg.append("g").classed("y-title", true);
        this.playAxisLabelGroup = this.svg.append("g").classed("play-axis-label", true);

        // Controls root
        const ctrl = d3.select(this.rootElement)
            .append("div")
            .classed("bar-race-controls", true)
            .style("position", "absolute")
            .style("bottom", "6px")
            .style("left", "0px")
            .style("right", "0px")
            .style("display", "flex")
            .style("gap", "8px")
            .style("align-items", "center")
            .style("justify-content", "center")
            .style("width", "100%")
            .style("font", "12px sans-serif")
            .style("pointer-events", "auto");
        this.controlsRoot = ctrl.node() as HTMLDivElement;

        // Play/Pause button
        this.playButton = ctrl.append("button")
            .text("▶️")
            .style("padding", "2px 8px")
            .style("cursor", "pointer")
            .on("click", () => this.togglePlay())
            .node() as HTMLButtonElement;

        // Reset button
        this.resetButton = ctrl.append("button")
            .text("⏮")
            .style("padding", "2px 8px")
            .style("cursor", "pointer")
            .on("click", () => this.reset())
            .node() as HTMLButtonElement;

        // Step button
        this.stepButton = ctrl.append("button")
            .text("Step")
            .style("padding", "2px 8px")
            .style("cursor", "pointer")
            .on("click", () => this.stepForward())
            .node() as HTMLButtonElement;


        // Frame label
        this.frameLabel = ctrl.append("div")
            .text("")
            .style("padding", "2px 4px")
            .style("background", "rgba(255,255,255,0.75)")
            .style("border", "1px solid #ccc")
            .style("border-radius", "3px")
            .node() as HTMLDivElement;

        // Progress scrubber
        this.progressSlider = ctrl.append("input")
            .attr("type", "range")
            .attr("min", 0)
            .attr("max", 0)
            .attr("step", 1)
            .style("width", "220px")
            .on("input", (event: any) => {
                const idx = parseInt(event.target.value, 10) || 0;
                this.stop();
                this.currentFrameIndex = Math.max(0, Math.min(idx, (this.frames.length || 1) - 1));
                const svgWidth = Number(this.svg.attr("width") || 0);
                const svgHeight = Number(this.svg.attr("height") || 0);
                const lay = this.computeLayout(svgWidth, svgHeight);
                const innerWidth = lay.innerWidth;
                const innerHeight = lay.innerHeight;
                this.layout = lay;
                if (this.frames.length) {
                    this.renderFrame(this.frames[this.currentFrameIndex], innerWidth, innerHeight);
                }
            })
            .node() as HTMLInputElement;
        // Init selection + tooltip service
        this.selectionManager = this.host.createSelectionManager();
        this.tooltipServiceWrapper = createTooltipServiceWrapper(this.host.tooltipService, this.rootElement);

        // Formatting model service
        this.formattingSettingsService = new FormattingSettingsService();
        try { this.controlsHeight = (this.controlsRoot?.getBoundingClientRect()?.height || 0) + 6; } catch { this.controlsHeight = 28; }

        // Custom tooltip
        const tt = d3.select(this.rootElement)
            .append("div")
            .classed("bar-race-tooltip", true)
            .style("position", "absolute")
            .style("pointer-events", "none")
            .style("display", "none")
            .style("background", "rgba(255,255,255,0.95)")
            .style("border", "1px solid #ccc")
            .style("border-radius", "4px")
            .style("padding", "6px 8px")
            .style("box-shadow", "0 2px 8px rgba(0,0,0,0.15)")
            .style("font", "12px sans-serif")
            .style("max-width", "280px");
        this.tooltipDiv = tt.node() as HTMLDivElement;

    }

 /**
     * Called every time the data or visual settings change
     */
    public update(options: VisualUpdateOptions) {
        this.eventService.renderingStarted(options);
        try {
            this.updateInternal(options);
            this.eventService.renderingFinished(options);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.eventService.renderingFailed(options, message);
            throw error;
        }
    }

    private updateInternal(options: VisualUpdateOptions) {
        // --- 1. Check for DataView ---
        const dataView = options.dataViews[0];

        // If no dataView exists, clear the visual and stop.
        // This prevents the "blank screen" crash.
        if (!dataView) {
            this.clearVisual();
            this.showLandingPage();
            return;
        }

        // --- 2. Parse Settings + Formatting Model ---
        this.settings = Visual.parseSettings(options);
        this.formattingModel = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);
        if (this.formattingModel?.animation) {
            const a = this.formattingModel.animation;
            this.settings.animation.duration = a.duration.value ?? this.settings.animation.duration;
            this.settings.animation.frameDelay = a.frameDelay?.value ?? this.settings.animation.frameDelay;
            const easingVal = (a.easing as any)?.value;
            this.settings.animation.easing = easingVal?.value ?? this.settings.animation.easing;
            this.settings.animation.autoplay = a.autoplay.value ?? this.settings.animation.autoplay;
            this.settings.animation.loop = a.loop.value ?? this.settings.animation.loop;
            this.settings.animation.showControls = a.showControls.value ?? this.settings.animation.showControls;
            this.settings.animation.maxBars = a.maxBars.value ?? this.settings.animation.maxBars;
            this.settings.animation.showPlayAxisLabel = a.showPlayAxisLabel?.value ?? this.settings.animation.showPlayAxisLabel;
            const playAxisLabelPositionVal = (a.playAxisLabelPosition as any)?.value;
            this.settings.animation.playAxisLabelPosition = playAxisLabelPositionVal?.value ?? this.settings.animation.playAxisLabelPosition;
            this.settings.animation.reduceMotion = a.reduceMotion?.value ?? this.settings.animation.reduceMotion;
        }
        if (this.formattingModel?.labels) {
            const l = this.formattingModel.labels as any;
            this.settings.labels.show = l.show.value ?? this.settings.labels.show;
            this.settings.labels.fontSize = l.fontSize.value ?? this.settings.labels.fontSize;
            this.settings.labels.displayUnits = this.normalizeDisplayUnits(l.displayUnits.value);
            this.settings.labels.precision = this.normalizePrecision(l.precision.value);
            this.settings.labels.categoryOnBars = l.categoryOnBars.value ?? this.settings.labels.categoryOnBars;
            const categoryFontFamilyVal = (l.categoryFontFamily as any)?.value;
            this.settings.labels.categoryFontFamily = categoryFontFamilyVal?.value ?? this.settings.labels.categoryFontFamily;
            this.settings.labels.categoryFontSize = l.categoryFontSize?.value ?? this.settings.labels.categoryFontSize;
            this.settings.labels.showCategoryWithImage = l.showCategoryWithImage?.value ?? this.settings.labels.showCategoryWithImage;
            this.settings.labels.showImageInTooltip = l.showImageInTooltip?.value ?? this.settings.labels.showImageInTooltip;
            this.settings.labels.iconOutline = l.iconOutline?.value ?? this.settings.labels.iconOutline;
            this.settings.labels.imagePadding = l.imagePadding?.value ?? this.settings.labels.imagePadding;
            this.settings.labels.imageInsideEnd = l.imageInsideEnd?.value ?? this.settings.labels.imageInsideEnd;
            this.settings.labels.labelsInside = l.labelsInside?.value ?? this.settings.labels.labelsInside;
            // color slice is ThemeColorData; also mirrored via metadata objects parsing below
        }
        if (this.formattingModel?.axes) {
            const a2 = this.formattingModel.axes as any;
            this.settings.axes.showXAxisLabels = a2.showXAxisLabels.value ?? this.settings.axes.showXAxisLabels;
            this.settings.axes.showYAxisLabels = a2.showYAxisLabels.value ?? this.settings.axes.showYAxisLabels;
            this.settings.axes.yAxisFontSize = a2.yAxisFontSize?.value ?? this.settings.axes.yAxisFontSize;
            this.settings.axes.showGridlines = a2.showGridlines.value ?? this.settings.axes.showGridlines;
            this.settings.axes.showXAxisTitle = a2.showXAxisTitle.value ?? this.settings.axes.showXAxisTitle;
            this.settings.axes.showYAxisTitle = a2.showYAxisTitle.value ?? this.settings.axes.showYAxisTitle;
        }
        if (this.formattingModel?.bars) {
            const b = this.formattingModel.bars as any;
            this.settings.bars.cornerRadius = b.cornerRadius.value ?? this.settings.bars.cornerRadius;
            const colorModeVal = (b.colorMode as any)?.value;
            this.settings.bars.colorMode = colorModeVal?.value ?? this.settings.bars.colorMode;
        }
        // Pull label color (fill) from objects if set via format pane
        const objs: any = dataView?.metadata?.objects || {};
        const lbl = objs.labels;
        const fillColor = lbl?.color?.solid?.color;
        if (fillColor) this.settings.labels.color = fillColor;

        // --- 3. Setup ---
        const width = options.viewport.width;
        const height = options.viewport.height;
        this.svg.attr("width", width).attr("height", height);
        this.svg.select<SVGRectElement>("rect.clear-catcher")
            .attr("x", 0).attr("y", 0)
            .attr("width", width).attr("height", height);

        const base = this.margin;
        const showX = this.settings?.axes?.showXAxisLabels !== false;
        const showY = (this.settings?.axes?.showYAxisLabels !== false) && !(this.settings?.labels?.categoryOnBars);
        const showXTitle = !!this.settings?.axes?.showXAxisTitle;
        const showYTitle = !!this.settings?.axes?.showYAxisTitle;
        const leftMargin = showY ? Math.max(base.left, showYTitle ? 54 : base.left) : 14;
        const bottomMargin = (showX ? (base.bottom + 16) : 8) + (showXTitle ? 18 : 0);
        const mLeft = leftMargin; const mBottom = bottomMargin; const mTop = base.top; const mRight = base.right;
        const innerWidth = width - mLeft - mRight;
        const controlsH = (this.controlsRoot?.offsetHeight || this.controlsHeight || 28);
        // Add generous padding so x-axis labels don't collide with controls
        const reserve = (this.settings?.animation?.showControls !== false) ? (controlsH + 40) : 0;
        const innerHeight = Math.max(0, height - mTop - mBottom - reserve);

        this.barContainer.attr("transform", `translate(${mLeft}, ${mTop})`);
        this.xAxisGroup.attr("transform", `translate(${mLeft}, ${mTop + innerHeight})`);
        this.yAxisGroup.attr("transform", `translate(${mLeft}, ${mTop})`);
        this.xTitleGroup.attr("transform", `translate(${mLeft + innerWidth/2}, ${mTop + innerHeight + 14})`);
        this.yTitleGroup.attr("transform", `translate(${mLeft - 14}, ${mTop + innerHeight/2}) rotate(-90)`);
        
        // --- 4. Data Transformation ---
        const hasPlayAxis = this.hasPlayAxis(dataView);
        this.updateControlsVisibility();

        if (hasPlayAxis) {
            this.frames = this.transformFrames(dataView, this.colorPalette, this.colorMap);
            if (this.frames.length === 0) {
                this.clearVisual();
                this.showLandingPage();
                return;
            }
            this.hideLandingPage();
            this.measureFormat = dataView.categorical?.values?.[0]?.source?.format as string | undefined;
            this.xTitleText = dataView.categorical?.values?.[0]?.source?.displayName || "";
            // prefer the category role column; fall back to the first category
            const catCol = dataView.categorical?.categories?.find(c => (c.source.roles as any)?.category) || dataView.categorical?.categories?.[0];
            this.yTitleText = catCol?.source?.displayName || "";

            // Reset to first frame on data change
            this.currentFrameIndex = Math.min(this.currentFrameIndex, this.frames.length - 1);
            this.updateSliderRange();
            this.renderFrame(this.frames[this.currentFrameIndex], innerWidth, innerHeight);

            // Autoplay if requested
            if (this.settings.animation.autoplay) {
                this.start();
            } else {
                this.stop();
            }
        } else {
            // Fallback: static view using category + measure only
            const data: BarDataPoint[] = this.transformData(dataView, this.colorPalette, this.colorMap);
            if (!data || data.length === 0) {
                this.clearVisual();
                this.showLandingPage();
                return;
            }
            this.hideLandingPage();
            this.measureFormat = dataView.categorical?.values?.[0]?.source?.format as string | undefined;
            this.xTitleText = dataView.categorical?.values?.[0]?.source?.displayName || "";
            const catCol2 = dataView.categorical?.categories?.find(c => (c.source.roles as any)?.category) || dataView.categorical?.categories?.[0];
            this.yTitleText = catCol2?.source?.displayName || "";
            const frame: Frame = { key: "", label: "", points: data };
            this.frames = [frame];
            this.currentFrameIndex = 0;
            this.updateSliderRange();
            this.renderFrame(frame, innerWidth, innerHeight);
            this.stop();
        }
    }

    // Compute margins and inner chart area taking into account axis visibility,
    // titles, and reserved space for the bottom controls.
    private computeLayout(width: number, height: number, data?: BarDataPoint[]) {
        const base = this.margin;
        const showX = this.settings?.axes?.showXAxisLabels !== false;
        const showY = (this.settings?.axes?.showYAxisLabels !== false) && !(this.settings?.labels?.categoryOnBars);
        const showXTitle = !!this.settings?.axes?.showXAxisTitle;
        const showYTitle = !!this.settings?.axes?.showYAxisTitle;
        const axisFontSize = this.getAxisFontSize();
        const labelPadding = 18;
        const yTitleReserve = showYTitle ? 28 : 0;
        const dynamicLabelWidth = showY && data?.length
            ? Math.max(...data.map(d => this.estimateAxisLabelWidth(this.getAxisDisplayLabel(d), axisFontSize)))
            : 0;
        const maxAdaptiveLeft = Math.max(base.left, Math.floor(width * 0.42));
        const adaptiveLeft = Math.min(maxAdaptiveLeft, Math.ceil(dynamicLabelWidth + labelPadding + yTitleReserve));
        const leftMargin = showY ? Math.max(base.left, adaptiveLeft) : 14;
        const bottomMargin = (showX ? (base.bottom + 16) : 8) + (showXTitle ? 18 : 0);
        const mLeft = leftMargin; const mBottom = bottomMargin; const mTop = base.top; const mRight = base.right;
        const innerWidth = Math.max(0, width - mLeft - mRight);
        const controlsH = (this.controlsRoot?.offsetHeight || this.controlsHeight || 28);
        const reserve = (this.settings?.animation?.showControls !== false) ? (controlsH + 40) : 0;
        const innerHeight = Math.max(0, height - mTop - mBottom - reserve);
        return { mLeft, mTop, innerWidth, innerHeight };
    }

    /**
     * Helper function to transform Power BI data into a clean array
     * This function now manages our persistent color map
     */
    private transformData(dataView: DataView, colorPalette: IColorPalette, colorMap: ColorMap): BarDataPoint[] {
        // const dataView = options.dataViews[0]; // <-- We no longer get this from options
        const dataPoints: BarDataPoint[] = [];

        // Check if we have valid categorical data
        if (dataView && // This check is still valid
            dataView.categorical &&
            dataView.categorical.categories &&
            dataView.categorical.categories[0] &&
            dataView.categorical.values &&
            dataView.categorical.values[0]) 
        {
            const categoryCols = dataView.categorical.categories.filter(c => (c.source.roles as any)?.category);
            const imageCols = dataView.categorical.categories.filter(c => (c.source.roles as any)?.imageUri);
            const categoryCol = categoryCols[0] || dataView.categorical.categories[0];
            const imageCol = imageCols[0];
            const categoryValues = categoryCol.values;
            const imageValues = imageCol?.values;
            const values = dataView.categorical.values[0].values;

            for (let i = 0; i < categoryValues.length; i++) {
                
                const categoryName: string = categoryValues[i]?.toString() ?? "";
                const imageRaw = imageValues?.[i] != null ? imageValues[i]!.toString() : "";
                const selectionId = this.host.createSelectionIdBuilder()
                    .withCategory(categoryCol, i)
                    .createSelectionId();
                const selectionKey = selectionId?.getKey?.();
                const categoryKey = this.getImageKey(categoryName);
                const imageUrl = this.resolveImageUrlByKey(categoryKey, categoryName, imageRaw);
                this.cacheImageAlias(selectionKey, imageUrl || undefined);
                let color: string;

                // Check if this category is already in our map
                if (colorMap[categoryName]) {
                    // 1. YES: Retrieve the stored color
                    color = colorMap[categoryName];
                } else {
                    // 2. NO: Get a new color from the palette
                    color = colorPalette.getColor(categoryName).value;
                    //    And STORE it in our map for next time
                    colorMap[categoryName] = color;
                }

                dataPoints.push({
                    category: categoryName,
                    value: values[i] as number,
                    color: color,
                    displayLabel: categoryName !== "" ? categoryName : undefined,
                    imageUrl: imageUrl || undefined,
                    imageKey: categoryKey,
                    selectionId
                });
            }
        }
        return dataPoints;
    }

    private hasPlayAxis(dataView: DataView): boolean {
        if (!dataView || !dataView.categorical || !dataView.categorical.categories) return false;
        const cats = dataView.categorical.categories;
        return cats.some(c => c && c.source && c.source.roles && (c.source.roles as any).playAxis);
    }

    private transformFrames(dataView: DataView, colorPalette: IColorPalette, colorMap: ColorMap): Frame[] {
        const framesMap = new Map<string, Frame>();
        if (!dataView || !dataView.categorical) return [];
        const categorical = dataView.categorical;
        const categories = categorical.categories || [];
        const valuesCol = categorical.values && categorical.values[0];
        if (!valuesCol) return [];

        // Locate indices for category and playAxis
        let categoryIndex = -1;
        let imageIndex = -1;
        let playIndex = -1;
        for (let i = 0; i < categories.length; i++) {
            const roles = categories[i]?.source?.roles as any;
            if (roles?.category) categoryIndex = i;
            if (roles?.imageUri) imageIndex = i;
            if (roles?.playAxis) playIndex = i;
        }
        if (categoryIndex === -1 || playIndex === -1) return [];

        const categoryCols = categories.filter(c => (c.source.roles as any)?.category);
        const imageCols = categories.filter(c => (c.source.roles as any)?.imageUri);
        const catCol = categoryCols[0] || categories[categoryIndex];
        const imageCol = imageCols[0] || (imageIndex >= 0 ? categories[imageIndex] : undefined);
        const catVals = catCol.values;
        const imageVals = imageCol?.values;
        const playVals = categories[playIndex].values;
        const measureVals = valuesCol.values;

        for (let i = 0; i < measureVals.length; i++) {
            const categoryName = catVals[i]?.toString() ?? "";
            const imageRaw = imageVals?.[i] != null ? imageVals[i]!.toString() : "";
            const playKeyRaw = playVals[i];
            const playKey = playKeyRaw instanceof Date ? (playKeyRaw as Date).toISOString() : String(playKeyRaw);
            const playLabel = playKeyRaw instanceof Date ? (playKeyRaw as Date).toLocaleDateString() : String(playKeyRaw);

            // Color
            let color = colorMap[categoryName];
            if (!color) {
                color = colorPalette.getColor(categoryName).value;
                colorMap[categoryName] = color;
            }

            const selectionId = this.host.createSelectionIdBuilder()
                .withCategory(catCol, i)
                .createSelectionId();
            const selectionKey = selectionId?.getKey?.();
            const categoryKey = this.getImageKey(categoryName);
            const imageUrl = this.resolveImageUrlByKey(categoryKey, categoryName, imageRaw);
            this.cacheImageAlias(selectionKey, imageUrl || undefined);

            const point: BarDataPoint = {
                category: categoryName,
                value: (measureVals[i] as number) ?? 0,
                color,
                displayLabel: categoryName !== "" ? categoryName : undefined,
                imageUrl: imageUrl || undefined,
                imageKey: categoryKey,
                selectionId
            };

            let frame = framesMap.get(playKey);
            if (!frame) {
                frame = { key: playKey, label: playLabel, points: [] };
                framesMap.set(playKey, frame);
            }
            frame.points.push(point);
        }

        // Preserve insertion order; optionally sort by key if keys are Dates
        const frames = Array.from(framesMap.values());
        // If keys look like ISO dates, sort by key
        const isoDateRegex = /^\d{4}-\d{2}-\d{2}T/;
        if (frames.length > 1 && isoDateRegex.test(frames[0].key)) {
            frames.sort((a, b) => a.key.localeCompare(b.key));
        }
        // Fill-forward images by category across frames
        const imageByCategory = new Map<string, string>();
        frames.forEach(f => {
            f.points.forEach(p => {
                if (p.imageUrl) {
                    const key = p.imageKey || this.getImageKey(p.category);
                    if (key) imageByCategory.set(key, p.imageUrl);
                }
            });
        });
        frames.forEach(f => {
            f.points.forEach(p => {
                if (!p.imageUrl) {
                    const key = p.imageKey || this.getImageKey(p.category);
                    const cached = key ? imageByCategory.get(key) : undefined;
                    if (cached) p.imageUrl = cached;
                }
            });
        });
        return frames;
    }

    private renderFrame(frame: Frame, innerWidth: number, innerHeight: number) {
        // Recompute layout each frame to avoid overlaps during animation
        const svgWidth = Number(this.svg.attr("width") || 0);
        const svgHeight = Number(this.svg.attr("height") || 0);

        // Update frame label
        this.frameLabel.textContent = frame.label || "";
        if (this.progressSlider) {
            this.progressSlider.value = String(this.currentFrameIndex);
        }

        // Apply Top N
        const maxBars = Math.max(1, this.settings?.animation?.maxBars || 10);
        const data = frame.points
            .slice()
            .sort((a, b) => b.value - a.value)
            .slice(0, maxBars);

        if (!data.length) {
            this.clearVisual();
            return;
        }
        // Ensure images persist even if the current frame doesn't carry the URI
        data.forEach(d => {
            if (!d.imageUrl) {
                const categoryKey = this.getImageKey(d.category);
                const key = d.imageKey || categoryKey;
                const cached = this.imageMap[key] || (categoryKey ? this.imageMap[categoryKey] : undefined);
                if (cached) d.imageUrl = cached;
            }
        });
        const lay = this.computeLayout(svgWidth, svgHeight, data);
        this.barContainer.attr("transform", `translate(${lay.mLeft}, ${lay.mTop})`);
        this.xAxisGroup.attr("transform", `translate(${lay.mLeft}, ${lay.mTop + lay.innerHeight})`);
        this.yAxisGroup.attr("transform", `translate(${lay.mLeft}, ${lay.mTop})`);
        this.xTitleGroup.attr("transform", `translate(${lay.mLeft + lay.innerWidth/2}, ${lay.mTop + lay.innerHeight + 14})`);
        this.yTitleGroup.attr("transform", `translate(${lay.mLeft - 14}, ${lay.mTop + lay.innerHeight/2}) rotate(-90)`);
        this.playAxisLabelGroup.attr("transform", `translate(${lay.mLeft}, ${lay.mTop})`);
        innerWidth = lay.innerWidth;
        innerHeight = lay.innerHeight;
        this.renderPlayAxisLabel(frame.label || "", innerWidth, innerHeight);

        // Scales
        const xScale = d3.scaleLinear()
            .domain([0, d3.max(data, d => d.value)!])
            .range([0, innerWidth])
            .nice();

        const yScale = d3.scaleBand()
            .domain(data.map(d => d.category))
            .range([0, innerHeight])
            .padding(0.1);

        // Axes
        const showGrid = this.settings?.axes?.showGridlines !== false;
        const xAxis = d3.axisBottom(xScale)
            .ticks(5)
            .tickSize(showGrid ? -innerHeight : 0);
        const labelMap = new Map<string, string>();
        const imageSet = new Set<string>();
        data.forEach(d => {
            if (d.imageUrl) imageSet.add(d.category);
            const label = this.getDisplayLabel(d);
            if (label) labelMap.set(d.category, label);
        });
        const axisLabelFor = (value: string) => {
            if (imageSet.has(value) && this.settings?.labels?.showCategoryWithImage === false) return "";
            return labelMap.get(value) ?? this.getFallbackCategoryLabel(value);
        };

        const yAxis = d3.axisLeft(yScale)
            .tickSize(0)
            .tickPadding(10)
            .tickFormat(((d: any) => {
                return axisLabelFor(String(d));
            }) as any);

        const duration = this.getAnimationDuration();
        const easeType = this.getEaseFunction();

        // High contrast handling
        const palette: any = this.colorPalette as any;
        const isHC = !!palette?.isHighContrast;
        const fg = palette?.foreground?.value || "#000";
        const bg = palette?.background?.value || "#fff";

        this.xAxisGroup
            .transition().duration(duration).ease(easeType)
            .call(xAxis as any)
            .selectAll(".tick line")
            .attr("stroke", showGrid ? (isHC ? fg : "#e0e0e0") : "none");
        const showXAxisLabels = this.settings?.axes?.showXAxisLabels !== false;
        const axisFontSize = this.getAxisFontSize();
        this.xAxisGroup.selectAll(".tick text")
            .attr("fill", showXAxisLabels ? (isHC ? fg : "#666") : "transparent")
            .style("display", showXAxisLabels ? null : "none")
            .style("font-family", this.getFontFamily())
            .style("font-size", `${axisFontSize}px`);
        this.xAxisGroup.select(".domain").remove();

        this.yAxisGroup
            .transition().duration(duration).ease(easeType)
            .call(yAxis as any);
        const showAxisCats = (this.settings?.axes?.showYAxisLabels !== false) && !(this.settings?.labels?.categoryOnBars);
        this.yAxisGroup.selectAll(".tick text")
            .attr("fill", showAxisCats ? (isHC ? fg : "#333") : "transparent")
            .style("font-family", this.getFontFamily())
            .style("font-size", `${axisFontSize}px`)
            .style("display", showAxisCats ? null : "none");
        this.yAxisGroup.select(".domain").remove();
        (this.yAxisGroup as any).raise?.();

        // Titles
        const catColName = this.yTitleText;
        const valColName = this.xTitleText;
        const showXTitleNow = !!this.settings?.axes?.showXAxisTitle;
        const showYTitleNow = !!this.settings?.axes?.showYAxisTitle;
        this.xTitleGroup.selectAll("text").data(showXTitleNow && showXAxisLabels ? [valColName] : []).join(
            enter => enter.append("text").attr("text-anchor", "middle").text(d => d),
            update => update.text(d => d),
            exit => exit.remove()
        ).style("font-family", this.getFontFamily()).attr("fill", isHC ? fg : "#666");
        this.yTitleGroup.selectAll("text").data(showYTitleNow && showAxisCats ? [catColName] : []).join(
            enter => enter.append("text").attr("text-anchor", "middle").text(d => d),
            update => update.text(d => d),
            exit => exit.remove()
        ).style("font-family", this.getFontFamily()).attr("fill", isHC ? fg : "#666");

        // Bars
        const bars = this.barContainer
            .selectAll<SVGRectElement, BarDataPoint>("rect")
            .data(data, (d: any) => d.category);

        const barRects = bars.join(
            enter => enter.append("rect")
                .attr("y", d => yScale(d.category)!)
                .attr("width", 0)
                .attr("height", yScale.bandwidth())
                .attr("rx", Math.max(0, this.settings?.bars?.cornerRadius ?? 0))
                .attr("ry", Math.max(0, this.settings?.bars?.cornerRadius ?? 0))
                .attr("fill", d => isHC ? fg : this.getBarFillColor(d.color))
                .attr("stroke", isHC ? fg : "none")
                .attr("stroke-width", isHC ? 1.2 : 0)
                .style("cursor", "pointer")
                .call(enter => enter.transition().duration(duration).ease(easeType)
                    .attr("width", d => xScale(d.value))
                ),
            update => update
                .call(update => update.transition().duration(duration).ease(easeType)
                    .attr("width", d => xScale(d.value))
                    .attr("y", d => yScale(d.category)!)
                    .attr("height", yScale.bandwidth())
                    .attr("rx", Math.max(0, this.settings?.bars?.cornerRadius ?? 0))
                    .attr("ry", Math.max(0, this.settings?.bars?.cornerRadius ?? 0))
                    .attr("fill", d => isHC ? fg : this.getBarFillColor(d.color))
                    .attr("stroke", isHC ? fg : "none")
                    .attr("stroke-width", isHC ? 1.2 : 0)
                ),
            exit => exit
                .remove()
        )
        .on("click", (event, d) => this.handleBarClick(event as any as MouseEvent, d))
        .on("contextmenu", (event) => this.selectionManager?.showContextMenu?.(null, { x: event.clientX, y: event.clientY }));

        if (this.settings?.labels?.showImageInTooltip) {
            barRects
                .on("mousemove", (event, d) => this.showCustomTooltip(event as MouseEvent, d, frame.label, formatter))
                .on("mouseleave", () => this.hideCustomTooltip());
        } else {
            this.hideCustomTooltip();
            barRects.on("mousemove", null).on("mouseleave", null);
            // Attach tooltips
            this.tooltipServiceWrapper.addTooltip(
                barRects,
                (args: TooltipEventArgs<BarDataPoint>) => this.getTooltipData(args.data, frame.label),
                (args) => args.data?.selectionId
            );
        }

        // Optional native title as fallback in Desktop
        this.barContainer.selectAll<SVGRectElement, BarDataPoint>("rect")
            .selectAll("title").data(d => [d]).join("title").text(d => `${this.getTooltipCategoryLabel(d)}: ${d.value}`);

        // Format based on measure format and settings
        const units = this.normalizeDisplayUnits(this.settings?.labels?.displayUnits);
        const precision = this.normalizePrecision(this.settings?.labels?.precision);
        const maxVal = d3.max(data, d => d.value) ?? 0;
        const fmtVal = units > 0 ? units : maxVal;
        const formatter = valueFormatter.create({ format: this.measureFormat, value: fmtVal, precision });

        // Category labels and Data labels
        this.renderCategoryIcons(data, xScale, yScale, formatter);
        this.renderCategoryLabels(data, xScale, yScale);
        this.renderDataLabels(data, xScale, yScale, innerWidth, innerHeight, formatter);
        this.renderPlayAxisLabel(frame.label || "", innerWidth, innerHeight);
    }

    private getFontFamily(): string {
        try {
            const el = (this.svg.node() as any)?.parentElement as HTMLElement;
            const fam = window.getComputedStyle(el).fontFamily;
            return fam || "Segoe UI, sans-serif";
        } catch { return "Segoe UI, sans-serif"; }
    }

    private getAxisFontSize(): number {
        return Math.max(8, this.settings?.axes?.yAxisFontSize ?? (this.settings?.labels?.fontSize ?? 12));
    }

    private getAxisDisplayLabel(d: BarDataPoint): string {
        if (d.imageUrl && this.settings?.labels?.showCategoryWithImage === false) return "";
        return this.getDisplayLabel(d) || this.getFallbackCategoryLabel(d.category);
    }

    private estimateAxisLabelWidth(label: string, fontSize: number): number {
        if (!label) return 0;

        return Math.max(fontSize * 1.5, label.length * fontSize * 0.58);
    }

    private getCategoryFontFamily(): string {
        const configured = this.settings?.labels?.categoryFontFamily;
        return configured && configured.trim().length ? configured : this.getFontFamily();
    }

    private renderDataLabels(data: BarDataPoint[], xScale: d3.ScaleLinear<number, number>, yScale: d3.ScaleBand<string>, innerWidth: number, innerHeight: number, formatter: any) {
        const show = this.settings?.animation ? (this.settings.labels?.show ?? true) : true;
        const fontSize = Math.max(8, this.settings?.labels?.fontSize ?? 12);
        const userColor = this.settings?.labels?.color || "#333";
        const palette: any = this.colorPalette as any;
        const isHC = !!palette?.isHighContrast;
        const fg = palette?.foreground?.value || "#000";
        const fontFamily = this.getFontFamily();
        const inside = !!this.settings?.labels?.labelsInside; // true = inside end
        const duration = this.getAnimationDuration();
        const easeType = this.getEaseFunction();

        // To avoid overlapping with category labels when both would be placed
        // outside the bar, we filter out value labels in that case so that at
        // least the category name remains visible.
        const catFontSize = Math.max(10, this.settings?.labels?.fontSize ?? 12);
        const catOnBars = !!this.settings?.labels?.categoryOnBars;
        const approxTextWidth = (s: string) => s ? Math.max(24, s.length * (catFontSize * 0.6)) : 0;
        const categoryWouldBeOutside = (d: BarDataPoint) => {
            const w = xScale(d.value);
            const label = this.getDisplayLabel(d);
            return w < (approxTextWidth(label) + 8);
        };

        const labelTextWidth = (d: BarDataPoint) => {
            const text = formatter?.format ? formatter.format(d.value) : String(d.value);
            return text ? Math.max(24, text.length * (fontSize * 0.62)) : 0;
        };
        const threshold = Math.max(16, fontSize * 2 + 8);
        const valueWouldBeInside = (d: BarDataPoint) => {
            const w = xScale(d.value);
            const outsideWouldOverflow = (w + labelTextWidth(d) + 8) > innerWidth;
            return (inside || outsideWouldOverflow) && w > threshold;
        };

        // If both category and value are outside, drop the value label
        const filtered = show
            ? data.filter(d => !(catOnBars && categoryWouldBeOutside(d) && !valueWouldBeInside(d)))
            : [];

        const labelsSel = this.barContainer
            .selectAll<SVGTextElement, BarDataPoint>("text.data-label")
            .data(filtered, (d: any) => d.category);
        const previousValueFor = (d: BarDataPoint) => {
            const previous = this.labelValueByCategory[d.category];
            return Number.isFinite(previous) ? previous : d.value;
        };
        const animateCounter = (node: SVGTextElement, d: BarDataPoint) => {
            const previousFrame = this.labelCounterFrameByCategory[d.category];
            if (previousFrame !== undefined) {
                window.cancelAnimationFrame(previousFrame);
                delete this.labelCounterFrameByCategory[d.category];
            }

            const startValue = previousValueFor(d);
            const endValue = d.value;
            if (duration <= 0 || startValue === endValue) {
                node.textContent = formatter.format(endValue);
                return;
            }

            const startedAt = performance.now();
            const run = (now: number) => {
                const progress = Math.min(1, Math.max(0, (now - startedAt) / duration));
                const eased = easeType(progress);
                node.textContent = formatter.format(startValue + ((endValue - startValue) * eased));

                if (progress < 1) {
                    this.labelCounterFrameByCategory[d.category] = window.requestAnimationFrame(run);
                } else {
                    node.textContent = formatter.format(endValue);
                    delete this.labelCounterFrameByCategory[d.category];
                }
            };

            this.labelCounterFrameByCategory[d.category] = window.requestAnimationFrame(run);
        };

        const posX = (d: BarDataPoint) => {
            const w = xScale(d.value);
            const placeInside = valueWouldBeInside(d);
            return placeInside ? Math.max(0, w - 4) : (w + 4);
        };
        const anchor = (d: BarDataPoint) => {
            const placeInside = valueWouldBeInside(d);
            return placeInside ? "end" : "start";
        };
        const fillFor = (d: BarDataPoint) => {
            if (isHC) return fg;
            const placeInside = valueWouldBeInside(d);
            if (placeInside) {
                return this.getInsideLabelColor(d.color);
            } else {
                // Outside the bar: ensure text contrasts with the background
                const bg = (this.colorPalette as any)?.background?.value || "#fff";
                const lumUser = this.relLuminance(userColor);
                const lumBg = this.relLuminance(bg);
                const ratio = this.contrastRatio(lumUser, lumBg);
                if (ratio < 3) {
                    // Prefer theme foreground, otherwise auto-pick black/white against bg
                    const auto = (lumBg < 0.5) ? "#fff" : "#000";
                    return fg || auto;
                }
                return userColor;
            }
        };

        const merged = labelsSel.join(
            enter => {
                const entered = enter.append("text")
                .attr("class", "data-label")
                .attr("x", d => posX(d))
                .attr("y", d => (yScale(d.category) ?? 0) + yScale.bandwidth() / 2)
                .attr("dominant-baseline", "middle")
                .attr("text-anchor", d => anchor(d))
                .attr("fill", d => fillFor(d))
                .style("font-family", fontFamily)
                .style("font-size", `${fontSize}px`)
                .text(d => formatter.format(previousValueFor(d)));
                entered.transition().duration(duration).ease(easeType)
                    .attr("x", d => posX(d))
                    .attr("y", d => (yScale(d.category) ?? 0) + yScale.bandwidth() / 2)
                    .attr("text-anchor", d => anchor(d))
                    .attr("fill", d => fillFor(d));
                entered.each(function(d) {
                    animateCounter(this, d);
                });
                return entered;
            },
            update => {
                update.transition().duration(duration).ease(easeType)
                    .attr("x", d => posX(d))
                    .attr("y", d => (yScale(d.category) ?? 0) + yScale.bandwidth() / 2)
                    .attr("text-anchor", d => anchor(d))
                    .attr("fill", d => fillFor(d))
                    .style("font-family", fontFamily)
                    .style("font-size", `${fontSize}px`);
                update.each(function(d) {
                    animateCounter(this, d);
                });
                return update;
            },
            exit => exit.each(d => {
                const previousFrame = this.labelCounterFrameByCategory[d.category];
                if (previousFrame !== undefined) {
                    window.cancelAnimationFrame(previousFrame);
                    delete this.labelCounterFrameByCategory[d.category];
                }
                delete this.labelValueByCategory[d.category];
            }).remove()
        );
        filtered.forEach(d => {
            this.labelValueByCategory[d.category] = d.value;
        });
        // Ensure labels are drawn above bars
        (merged as any).raise?.();
    }

    private relLuminance(colorStr: string): number {
        try {
            const c = (d3.color(colorStr) as any);
            if (!c) return 1;
            const [r, g, b] = [c.r, c.g, c.b].map((v: number) => {
                const s = v / 255;
                return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
            });
            return 0.2126 * r + 0.7152 * g + 0.0722 * b;
        } catch { return 1; }
    }

    // WCAG contrast ratio using relative luminance values
    // Expects luminance in [0,1]
    private contrastRatio(l1: number, l2: number): number {
        const [L1, L2] = l1 >= l2 ? [l1, l2] : [l2, l1];
        return (L1 + 0.05) / (L2 + 0.05);
    }

    private renderCategoryLabels(data: BarDataPoint[], xScale: d3.ScaleLinear<number, number>, yScale: d3.ScaleBand<string>) {
        const enabled = !!this.settings?.labels?.categoryOnBars;
        const fontFamily = this.getCategoryFontFamily();
        const userColor = this.settings?.labels?.color || "#333";
        const fontSize = Math.max(8, this.settings?.labels?.categoryFontSize ?? this.settings?.labels?.fontSize ?? 12);
        const duration = this.getAnimationDuration();
        const easeType = this.getEaseFunction();
        const palette: any = this.colorPalette as any;
        const isHC = !!palette?.isHighContrast;
        const fg = palette?.foreground?.value || "#000";
        const bg = palette?.background?.value || "#fff";
        const labelData = enabled ? data.filter(d => !!this.getDisplayLabel(d)) : [];
        const labels = this.barContainer
            .selectAll<SVGTextElement, BarDataPoint>("text.cat-label")
            .data(labelData, (d: any) => d.category);

        // Determine position and color. If bar is too short to hold the text,
        // place the label just outside the bar with readable color.
        const approxTextWidth = (s: string) => s ? Math.max(24, s.length * (fontSize * 0.6)) : 0;
        const posX = (d: BarDataPoint) => {
            const w = xScale(d.value);
            const label = this.getDisplayLabel(d);
            const needOutside = w < (approxTextWidth(label) + 8);
            return needOutside ? (w + 6) : 6; // inside-left margin = 6
        };
        const fillFor = (d: BarDataPoint) => {
            if (isHC) return fg;
            const w = xScale(d.value);
            const label = this.getDisplayLabel(d);
            const needOutside = w < (approxTextWidth(label) + 8);
            if (!needOutside) {
                return this.getInsideLabelColor(d.color);
            }
            const lumUser = this.relLuminance(userColor);
            const lumBg = this.relLuminance(bg);
            const ratio = this.contrastRatio(lumUser, lumBg);
            if (ratio < 3) {
                const auto = (lumBg < 0.5) ? "#fff" : "#000";
                return fg || auto;
            }
            return userColor;
        };

        const merged = labels.join(
            enter => enter.append("text")
                .attr("class", "cat-label")
                .attr("x", d => posX(d))
                .attr("y", d => (yScale(d.category) ?? 0) + yScale.bandwidth() / 2)
                .attr("dominant-baseline", "middle")
                .attr("text-anchor", "start")
                .style("font-family", fontFamily)
                .style("font-size", `${fontSize}px`)
                .attr("fill", d => fillFor(d))
                .text(d => this.getDisplayLabel(d))
                .call(enter => enter.transition().duration(duration).ease(easeType)
                    .attr("x", d => posX(d))
                    .attr("y", d => (yScale(d.category) ?? 0) + yScale.bandwidth() / 2)
                ),
            update => update
                .attr("x", d => posX(d))
                .style("font-family", fontFamily)
                .style("font-size", `${fontSize}px`)
                .attr("fill", d => fillFor(d))
                .call(update => update.transition().duration(duration).ease(easeType)
                    .attr("x", d => posX(d))
                    .attr("y", d => (yScale(d.category) ?? 0) + yScale.bandwidth() / 2)
                )
                .text(d => this.getDisplayLabel(d)),
            exit => exit.remove()
        );
        (merged as any).raise?.();
    }

    private renderCategoryIcons(data: BarDataPoint[], xScale: d3.ScaleLinear<number, number>, yScale: d3.ScaleBand<string>, formatter: any) {
        const iconData = data.filter(d => !!d.imageUrl);
        const duration = this.getAnimationDuration();
        const easeType = this.getEaseFunction();
        const placeEnd = !!this.settings?.labels?.imageInsideEnd;
        const showLabels = this.settings?.animation ? (this.settings.labels?.show ?? true) : true;
        const labelsInside = !!this.settings?.labels?.labelsInside;
        const fontSize = Math.max(8, this.settings?.labels?.fontSize ?? 12);
        const imagePadding = Math.max(0, Math.min(24, this.settings?.labels?.imagePadding ?? 2));
        const iconOutline = this.settings?.labels?.iconOutline ? "1px solid rgba(0,0,0,0.35)" : "none";
        const iconRadius = this.settings?.labels?.iconOutline ? "3px" : "0px";
        this.barContainer.selectAll("image.cat-icon").remove();
        const icons = this.barContainer
            .selectAll<SVGForeignObjectElement, BarDataPoint>("foreignObject.cat-icon")
            .data(iconData, (d: any) => d.category);

        const iconSize = Math.max(0, yScale.bandwidth() - imagePadding * 2);
        const approxTextWidth = (s: string) => s ? Math.max(0, s.length * (fontSize * 0.6)) : 0;
        const threshold = Math.max(16, fontSize * 2 + 8);
        const iconX = (d: BarDataPoint) => {
            const w = xScale(d.value);
            const maxInside = Math.max(0, w - iconSize - imagePadding);
            if (placeEnd) {
                let offset = 0;
                if (showLabels && labelsInside && w > threshold) {
                    const text = formatter?.format ? formatter.format(d.value) : String(d.value);
                    offset = approxTextWidth(text) + 6;
                }
                const desired = w - iconSize - imagePadding - offset;
                return Math.max(0, Math.min(desired, maxInside));
            }
            return Math.min(imagePadding, maxInside);
        };

        const merged = icons.join(
            enter => {
                const fo = enter.append("foreignObject")
                    .attr("class", "cat-icon")
                    .attr("x", d => placeEnd ? 0 : iconX(d))
                    .attr("y", d => (yScale(d.category) ?? 0) + imagePadding)
                    .attr("width", iconSize)
                    .attr("height", iconSize)
                    .style("pointer-events", "none");
                fo.append("xhtml:img")
                    .attr("src", d => d.imageUrl || "")
                    .attr("width", "100%")
                    .attr("height", "100%")
                    .style("object-fit", "contain")
                    .style("border", iconOutline)
                    .style("border-radius", iconRadius);
                return fo.call(enterSel => enterSel.transition().duration(duration).ease(easeType)
                    .attr("x", d => iconX(d))
                    .attr("y", d => (yScale(d.category) ?? 0) + imagePadding)
                    .attr("width", iconSize)
                    .attr("height", iconSize)
                );
            },
            update => {
                update.select("img")
                    .attr("src", d => d.imageUrl || "")
                    .style("border", iconOutline)
                    .style("border-radius", iconRadius);
                return update.call(updateSel => updateSel.transition().duration(duration).ease(easeType)
                    .attr("x", d => iconX(d))
                    .attr("y", d => (yScale(d.category) ?? 0) + imagePadding)
                    .attr("width", iconSize)
                    .attr("height", iconSize)
                );
            },
            exit => exit.remove()
        );
        (merged as any).raise?.();
    }

    private isImageCategoryValue(category: string): boolean {
        return typeof category === "string" && category.trim().toLowerCase().startsWith("data:image/");
    }

    private getDisplayLabel(d: BarDataPoint): string {
        if (d.imageUrl && this.settings?.labels?.showCategoryWithImage === false) return "";
        if (d.displayLabel && d.displayLabel.trim().length) return d.displayLabel.trim();
        return this.isImageCategoryValue(d.category) ? "" : d.category;
    }

    private getFallbackCategoryLabel(category: string): string {
        return this.isImageCategoryValue(category) ? "" : category;
    }

    private getTooltipCategoryLabel(d: BarDataPoint): string {
        if (d.displayLabel && d.displayLabel.trim().length) return d.displayLabel.trim();
        return this.isImageCategoryValue(d.category) ? "Image" : d.category;
    }

    private normalizeDisplayUnits(value: unknown): number {
        const rawValue = typeof value === "object" && value !== null && "value" in value
            ? (value as { value?: unknown }).value
            : value;
        const numericValue = Number(rawValue);
        const validUnits = [0, 1, 1000, 1000000, 1000000000];

        return validUnits.includes(numericValue) ? numericValue : 1;
    }

    private normalizePrecision(value: unknown): number {
        const precision = Number(value);

        return Number.isFinite(precision) ? Math.max(0, Math.min(6, precision)) : 0;
    }

    private getBarFillColor(baseColor: string): string {
        const mode = this.settings?.bars?.colorMode || "auto";
        if (mode === "auto") return baseColor;
        try {
            const c = d3.hsl(baseColor);
            const delta = mode === "light" ? 0.25 : -0.25;
            c.l = Math.max(0.1, Math.min(0.9, c.l + delta));
            return c.formatHex();
        } catch {
            return baseColor;
        }
    }

    private getInsideLabelColor(baseColor: string): string {
        const mode = this.settings?.bars?.colorMode || "auto";
        if (mode === "light") return "#000";
        if (mode === "dark") return "#fff";
        const barColor = this.getBarFillColor(baseColor);
        const lum = this.relLuminance(barColor);
        return lum < 0.5 ? "#fff" : "#000";
    }

    private normalizeImageUrl(value: string): string {
        if (!value) return "";
        const trimmed = value.trim();
        if (!trimmed.toLowerCase().startsWith("data:image/")) return "";
        return trimmed.replace(/\s+/g, "");
    }

    private getImageKey(categoryName: string): string {
        return (categoryName || "").trim().toLowerCase();
    }

    private resolveImageUrlByKey(imageKey: string | undefined, categoryName: string, imageRaw: string): string {
        const key = imageKey || this.getImageKey(categoryName);
        const categoryKey = this.getImageKey(categoryName);
        const fromField = this.normalizeImageUrl(imageRaw);
        const fromCategory = this.normalizeImageUrl(categoryName);
        const chosen = fromField || fromCategory;
        if (chosen) {
            this.imageMap[key] = chosen;
            if (categoryKey && categoryKey !== key) {
                this.imageMap[categoryKey] = chosen;
            }
            return chosen;
        }
        return this.imageMap[key] || (categoryKey ? this.imageMap[categoryKey] : "") || "";
    }

    private cacheImageAlias(key: string | undefined, url: string | undefined) {
        if (!key || !url) return;
        this.imageMap[key] = url;
    }

    // No local formatter now; we use Power BI valueFormatter based on measure format

    private start() {
        this.stop();
        this.playButton.textContent = "⏸";
        const step = () => {
            if (!this.frames.length) return;
            this.currentFrameIndex++;
            if (this.currentFrameIndex >= this.frames.length) {
                if (this.settings.animation.loop) {
                    this.currentFrameIndex = 0;
                } else {
                    this.currentFrameIndex = this.frames.length - 1;
                    this.stop();
                    return;
                }
            }
            const svgWidth = Number(this.svg.attr("width") || 0);
            const svgHeight = Number(this.svg.attr("height") || 0);
            const innerWidth = svgWidth - this.margin.left - this.margin.right;
            const innerHeight = svgHeight - this.margin.top - this.margin.bottom;
            this.renderFrame(this.frames[this.currentFrameIndex], innerWidth, innerHeight);
        };
        const delay = this.getFrameInterval();
        this.timer = window.setInterval(step, delay);
        this.playButton.textContent = "⏸";
    }

    private stop() {
        if (this.timer !== undefined) {
            clearInterval(this.timer);
            this.timer = undefined;
        }
        this.playButton.textContent = "▶️";
    }

    private togglePlay() {
        if (this.timer === undefined) {
            this.start();
        } else {
            this.stop();
        }
    }

    private reset() {
        this.stop();
        this.currentFrameIndex = 0;
        const svgWidth = Number(this.svg.attr("width") || 0);
        const svgHeight = Number(this.svg.attr("height") || 0);
        const innerWidth = svgWidth - this.margin.left - this.margin.right;
        const innerHeight = svgHeight - this.margin.top - this.margin.bottom;
        if (this.frames.length) {
            this.renderFrame(this.frames[this.currentFrameIndex], innerWidth, innerHeight);
        }
        // Clear selections on reset
        this.selectionManager?.clear?.();
        if (this.progressSlider) this.progressSlider.value = "0";
    }

    private stepForward() {
        this.stop();
        if (!this.frames.length) return;
        this.currentFrameIndex++;
        if (this.currentFrameIndex >= this.frames.length) {
            this.currentFrameIndex = this.settings.animation.loop ? 0 : this.frames.length - 1;
        }
        const svgWidth = Number(this.svg.attr("width") || 0);
        const svgHeight = Number(this.svg.attr("height") || 0);
        const lay = this.computeLayout(svgWidth, svgHeight);
        this.layout = lay;
        this.renderFrame(this.frames[this.currentFrameIndex], lay.innerWidth, lay.innerHeight);
        if (this.progressSlider) this.progressSlider.value = String(this.currentFrameIndex);
    }

    private getAnimationDuration(): number {
        return this.settings?.animation?.reduceMotion ? 0 : (this.settings?.animation?.duration ?? 0);
    }

    private getFrameInterval(): number {
        const duration = this.getAnimationDuration();
        const delay = Math.max(0, this.settings?.animation?.frameDelay ?? 0);
        return Math.max(50, duration + delay);
    }

    private getEaseFunction(): (t: number) => number {
        const easing = this.settings?.animation?.easing || "linear";
        switch (easing) {
            case "easeOut":
                return d3.easeCubicOut;
            case "easeInOut":
                return d3.easeCubicInOut;
            default:
                return d3.easeLinear;
        }
    }

    private renderPlayAxisLabel(label: string, innerWidth: number, innerHeight: number) {
        const show = this.settings?.animation?.showPlayAxisLabel !== false && !!label;
        if (!show) {
            this.playAxisLabelGroup.selectAll("*").remove();
            return;
        }

        const position = this.settings?.animation?.playAxisLabelPosition === "bottom" ? "bottom" : "top";
        const x = Math.max(0, innerWidth - 12);
        const y = position === "bottom" ? Math.max(18, innerHeight - 14) : 24;
        const palette: any = this.colorPalette as any;
        const isHC = !!palette?.isHighContrast;
        const fill = isHC ? (palette?.foreground?.value || "#000") : "rgba(31, 41, 55, 0.72)";
        const stroke = isHC ? (palette?.background?.value || "#fff") : "rgba(255, 255, 255, 0.88)";

        this.playAxisLabelGroup
            .selectAll<SVGTextElement, string>("text.play-axis-current-value")
            .data([label])
            .join(
                enter => enter.append("text")
                    .attr("class", "play-axis-current-value")
                    .attr("text-anchor", "end")
                    .attr("dominant-baseline", "middle")
                    .style("font-family", this.getFontFamily())
                    .style("font-size", "22px")
                    .style("font-weight", "700")
                    .style("paint-order", "stroke")
                    .style("stroke-width", "4px")
                    .style("pointer-events", "none")
                    .text(d => d),
                update => update.text(d => d),
                exit => exit.remove()
            )
            .attr("x", x)
            .attr("y", y)
            .attr("fill", fill)
            .attr("stroke", stroke);

        (this.playAxisLabelGroup as any).raise?.();
    }

    private updateSliderRange() {
        if (!this.progressSlider) return;
        const max = Math.max(0, (this.frames?.length || 1) - 1);
        this.progressSlider.max = String(max);
        this.progressSlider.min = "0";
        this.progressSlider.step = "1";
        this.progressSlider.value = String(Math.min(this.currentFrameIndex, max));
    }

    private handleBarClick(event: MouseEvent, d: BarDataPoint) {
        const isCtrlPressed = !!(event?.ctrlKey || event?.metaKey);
        this.selectionManager?.select(d.selectionId, isCtrlPressed);
    }

    private showCustomTooltip(event: MouseEvent, d: BarDataPoint, frameLabel: string, formatter: any) {
        if (!this.tooltipDiv) return;
        const label = this.getTooltipCategoryLabel(d);
        const value = formatter?.format ? formatter.format(d.value) : String(d.value);
        this.tooltipDiv.textContent = "";
        if (this.settings?.labels?.showImageInTooltip && d.imageUrl) {
            const imgWrap = document.createElement("div");
            imgWrap.style.marginBottom = "6px";
            const img = document.createElement("img");
            img.src = d.imageUrl;
            img.style.width = "64px";
            img.style.height = "64px";
            img.style.objectFit = "contain";
            imgWrap.appendChild(img);
            this.tooltipDiv.appendChild(imgWrap);
        }
        if (frameLabel) {
            const row = document.createElement("div");
            const strong = document.createElement("strong");
            strong.textContent = "Frame:";
            row.appendChild(strong);
            row.appendChild(document.createTextNode(` ${frameLabel}`));
            this.tooltipDiv.appendChild(row);
        }
        const catRow = document.createElement("div");
        const catStrong = document.createElement("strong");
        catStrong.textContent = "Category:";
        catRow.appendChild(catStrong);
        catRow.appendChild(document.createTextNode(` ${label}`));
        this.tooltipDiv.appendChild(catRow);
        const valRow = document.createElement("div");
        const valStrong = document.createElement("strong");
        valStrong.textContent = "Value:";
        valRow.appendChild(valStrong);
        valRow.appendChild(document.createTextNode(` ${value}`));
        this.tooltipDiv.appendChild(valRow);
        this.tooltipDiv.style.display = "block";
        const hostRect = this.rootElement.getBoundingClientRect();
        const ttRect = this.tooltipDiv.getBoundingClientRect();
        let left = event.clientX - hostRect.left + 12;
        let top = event.clientY - hostRect.top + 12;
        if (left + ttRect.width > hostRect.width) left = hostRect.width - ttRect.width - 8;
        if (top + ttRect.height > hostRect.height) top = hostRect.height - ttRect.height - 8;
        if (left < 0) left = 0;
        if (top < 0) top = 0;
        this.tooltipDiv.style.left = `${left}px`;
        this.tooltipDiv.style.top = `${top}px`;
    }

    private hideCustomTooltip() {
        if (this.tooltipDiv) {
            this.tooltipDiv.style.display = "none";
        }
    }

    private getTooltipData(d: BarDataPoint, label: string): powerbi.extensibility.VisualTooltipDataItem[] {
        const units = this.normalizeDisplayUnits(this.settings?.labels?.displayUnits);
        const precision = this.normalizePrecision(this.settings?.labels?.precision);
        const fmt = valueFormatter.create({ format: this.measureFormat, value: units > 0 ? units : Math.abs(d.value), precision });
        const categoryLabel = this.getTooltipCategoryLabel(d);
        const items: powerbi.extensibility.VisualTooltipDataItem[] = [
            { displayName: "Category", value: categoryLabel },
            { displayName: "Value", value: fmt.format(d.value) }
        ];
        if (label) {
            items.unshift({ displayName: "Frame", value: label });
        }
        return items;
    }

    private updateControlsVisibility() {
        const visible = this.settings?.animation?.showControls !== false;
        if (this.controlsRoot) {
            this.controlsRoot.style.display = visible ? "flex" : "none";
        }
    }

    private showLandingPage() {
        if (this.landingPage) return;

        this.stop();
        if (this.controlsRoot) {
            this.controlsRoot.style.display = "none";
        }
        if (this.tooltipDiv) {
            this.tooltipDiv.style.display = "none";
        }

        const page = document.createElement("div");
        page.className = "bar-racing-landing-page";
        page.setAttribute("role", "note");

        const s = page.style;
        s.position = "absolute";
        s.inset = "0";
        s.zIndex = "999";
        s.boxSizing = "border-box";
        s.display = "flex";
        s.flexDirection = "column";
        s.alignItems = "center";
        s.justifyContent = "center";
        s.gap = "14px";
        s.padding = "24px";
        s.background = "#ffffff";
        s.color = "#1f2933";
        s.fontFamily = "Segoe UI, Arial, sans-serif";
        s.textAlign = "center";
        s.overflowY = "auto";

        const icon = document.createElement("img");
        icon.setAttribute("aria-hidden", "true");
        const iconSvg = `
            <svg xmlns="http://www.w3.org/2000/svg" width="132" height="132" viewBox="0 0 512 512">
                <rect x="24" y="24" width="464" height="464" rx="96" fill="#f7be23" stroke="#20252e" stroke-width="18"/>
                <rect x="48" y="48" width="416" height="172" rx="70" fill="#ffd757" opacity=".55"/>
                <rect x="68" y="118" width="110" height="110" rx="24" fill="#fff" stroke="#20252e" stroke-width="10"/>
                <circle cx="110" cy="160" r="18" fill="#1984c4"/>
                <path d="M86 214l38-36 24 26 18-20v30z" fill="#27a763"/>
                <rect x="205" y="128" width="213" height="46" rx="22" fill="#26619c" stroke="#20252e" stroke-width="8"/>
                <rect x="205" y="214" width="169" height="46" rx="22" fill="#1e8758" stroke="#20252e" stroke-width="8"/>
                <rect x="205" y="300" width="113" height="46" rx="22" fill="#c44636" stroke="#20252e" stroke-width="8"/>
                <path d="M92 306v94l82-47z" fill="#20252e"/>
                <path d="M108 330v46l40-23z" fill="#f7be23"/>
                <text x="204" y="422" font-family="Segoe UI, Arial, sans-serif" font-size="52" font-weight="700" fill="#20252e">BR</text>
            </svg>`;
        icon.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(iconSvg)}`;
        icon.alt = "";
        icon.style.width = "132px";
        icon.style.height = "132px";
        page.appendChild(icon);

        const title = document.createElement("h2");
        title.textContent = "Bar Racing Image";
        title.style.margin = "0";
        title.style.color = "#1f5d99";
        title.style.fontSize = "34px";
        title.style.fontWeight = "700";
        page.appendChild(title);

        const subtitle = document.createElement("p");
        subtitle.textContent = "Animated bar race chart with optional category images.";
        subtitle.style.margin = "0";
        subtitle.style.fontSize = "17px";
        subtitle.style.color = "#52616f";
        page.appendChild(subtitle);

        const author = document.createElement("p");
        author.textContent = "Developed by S. Rathinagiri";
        author.style.margin = "0";
        author.style.fontSize = "14px";
        author.style.fontWeight = "600";
        page.appendChild(author);

        const box = document.createElement("div");
        box.style.boxSizing = "border-box";
        box.style.width = "min(560px, 96%)";
        box.style.marginTop = "8px";
        box.style.padding = "18px 20px";
        box.style.border = "1px solid #d7dde5";
        box.style.borderRadius = "6px";
        box.style.background = "#f8fafc";
        box.style.textAlign = "left";

        const boxTitle = document.createElement("div");
        boxTitle.textContent = "To visualize your data:";
        boxTitle.style.fontWeight = "600";
        boxTitle.style.marginBottom = "10px";
        boxTitle.style.fontSize = "15px";
        box.appendChild(boxTitle);

        const list = document.createElement("ol");
        list.style.margin = "0";
        list.style.paddingLeft = "20px";
        list.style.fontSize = "14px";
        list.style.lineHeight = "1.55";

        [
            "Add a category field to Category.",
            "Add a numeric field to Measure.",
            "Add a period, date, or sequence field to Play Axis.",
            "Optionally add data:image values to Image URI for category icons."
        ].forEach(step => {
            const li = document.createElement("li");
            li.textContent = step;
            list.appendChild(li);
        });

        box.appendChild(list);
        page.appendChild(box);

        const note = document.createElement("p");
        note.textContent = "External image URLs are not loaded; use embedded data:image values.";
        note.style.margin = "2px 0 0";
        note.style.fontSize = "12px";
        note.style.color = "#64748b";
        page.appendChild(note);

        this.rootElement.appendChild(page);
        this.landingPage = page;
    }

    private hideLandingPage() {
        if (!this.landingPage) return;
        this.rootElement.removeChild(this.landingPage);
        this.landingPage = null;
    }

    /**
     * Helper function to clear the visual when no data is present
     */
    private clearVisual() {
        this.stop();
        this.barContainer.selectAll("*").remove();
        this.xAxisGroup.selectAll("*").remove();
        this.yAxisGroup.selectAll("*").remove();
        this.xTitleGroup.selectAll("*").remove();
        this.yTitleGroup.selectAll("*").remove();
        this.playAxisLabelGroup.selectAll("*").remove();
        Object.keys(this.labelCounterFrameByCategory).forEach(category => {
            window.cancelAnimationFrame(this.labelCounterFrameByCategory[category]);
        });
        this.labelCounterFrameByCategory = {};
        this.labelValueByCategory = {};
        if (this.frameLabel) {
            this.frameLabel.textContent = "";
        }
        if (this.progressSlider) {
            this.progressSlider.max = "0";
            this.progressSlider.value = "0";
        }
    }
/**
     * This function parses the settings in the formatting pane
     */
    private static parseSettings(options: VisualUpdateOptions): VisualSettings {
        // This line gets the settings from Power BI and uses our defaults (250)
        return VisualSettings.parse(options.dataViews[0]) as VisualSettings;
    }

    /**
     * This function tells Power BI what settings to display in the formatting pane
     */
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
        // Legacy format pane fallback (kept for compatibility)
        const settings = this.settings || new VisualSettings();
        const instances: VisualObjectInstance[] = [];
        if (options.objectName === "animation") {
            instances.push({
                objectName: options.objectName,
                properties: {
                    duration: settings.animation.duration,
                    frameDelay: settings.animation.frameDelay,
                    easing: settings.animation.easing,
                    autoplay: settings.animation.autoplay,
                    loop: settings.animation.loop,
                    showControls: settings.animation.showControls,
                    maxBars: settings.animation.maxBars,
                    showPlayAxisLabel: settings.animation.showPlayAxisLabel,
                    playAxisLabelPosition: settings.animation.playAxisLabelPosition,
                    reduceMotion: settings.animation.reduceMotion
                },
                selector: null
            });
        }
        if (options.objectName === "labels") {
            instances.push({
                objectName: options.objectName,
                properties: {
                    show: settings.labels.show,
                    fontSize: settings.labels.fontSize,
                    color: { solid: { color: settings.labels.color } },
                    displayUnits: settings.labels.displayUnits,
                    precision: settings.labels.precision,
                    categoryOnBars: settings.labels.categoryOnBars,
                    categoryFontFamily: settings.labels.categoryFontFamily,
                    categoryFontSize: settings.labels.categoryFontSize,
                    showCategoryWithImage: settings.labels.showCategoryWithImage,
                    showImageInTooltip: settings.labels.showImageInTooltip,
                    iconOutline: settings.labels.iconOutline,
                    imagePadding: settings.labels.imagePadding,
                    imageInsideEnd: settings.labels.imageInsideEnd
                },
                selector: null
            });
        }
        if (options.objectName === "labels") {
            instances.push({
                objectName: options.objectName,
                properties: {
                    show: settings.labels.show,
                    fontSize: settings.labels.fontSize,
                    color: { solid: { color: settings.labels.color } },
                    displayUnits: settings.labels.displayUnits,
                    precision: settings.labels.precision,
                    categoryOnBars: settings.labels.categoryOnBars,
                    categoryFontFamily: settings.labels.categoryFontFamily,
                    categoryFontSize: settings.labels.categoryFontSize,
                    showCategoryWithImage: settings.labels.showCategoryWithImage,
                    showImageInTooltip: settings.labels.showImageInTooltip,
                    iconOutline: settings.labels.iconOutline,
                    imagePadding: settings.labels.imagePadding,
                    imageInsideEnd: settings.labels.imageInsideEnd,
                    labelsInside: settings.labels.labelsInside
                },
                selector: null
            });
        }
        if (options.objectName === "axes") {
            instances.push({
                objectName: options.objectName,
                properties: {
                    showXAxisLabels: settings.axes.showXAxisLabels,
                    showYAxisLabels: settings.axes.showYAxisLabels,
                    yAxisFontSize: settings.axes.yAxisFontSize,
                    showGridlines: settings.axes.showGridlines,
                    showXAxisTitle: settings.axes.showXAxisTitle,
                    showYAxisTitle: settings.axes.showYAxisTitle
                },
                selector: null
            });
        }
        if (options.objectName === "bars") {
            instances.push({
                objectName: options.objectName,
                properties: {
                    cornerRadius: settings.bars.cornerRadius,
                    colorMode: settings.bars.colorMode
                },
                selector: null
            });
        }
        return instances;
    }    

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingModel);
    }
}
