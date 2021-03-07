/*
*  Power BI Visualizations
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
import powerbi from "powerbi-visuals-api";
import "../style/style.less";
import "@babel/polyfill";

import { valueFormatter, textMeasurementService } from "powerbi-visuals-utils-formattingutils";
import IValueFormatter = valueFormatter.IValueFormatter;

import ValueFormatter = valueFormatter.valueFormatter;
import TextMeasurementService = textMeasurementService.textMeasurementService;
import TextProperties = textMeasurementService.TextProperties;

import { axis } from "powerbi-visuals-utils-chartutils";
import LabelLayoutStrategy = axis.LabelLayoutStrategy;

import { manipulation } from "powerbi-visuals-utils-svgutils";
import translate = manipulation.translate;

import { createLinearColorScale, LinearColorScale, ColorHelper } from "powerbi-visuals-utils-colorutils";

type Selection<T> = d3.Selection<any, T, any, any>;
type Quantile<T> = d3.ScaleQuantile<T>;

import * as d3 from "d3";

import * as _ from "lodash-es";

import { pixelConverter as PixelConverter } from "powerbi-visuals-utils-typeutils";

import IColorPalette = powerbi.extensibility.IColorPalette;
import DataViewObjectPropertyIdentifier = powerbi.DataViewObjectPropertyIdentifier;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IViewport = powerbi.IViewport;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumeration = powerbi.VisualObjectInstanceEnumeration;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import DataViewTable = powerbi.DataViewTable;
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import DataViewMetadata = powerbi.DataViewMetadata;
import ILocalizationManager = powerbi.
extensibility.ILocalizationManager;

import {
    IColorArray,
    IMargin,
    LegendMapChartData,
    LegendMapDataPoint
} from "./dataInterfaces";

import {
    Settings,
    colorbrewer
} from "./settings";

// powerbi.extensibility.utils.tooltip
import {
    TooltipEventArgs,
    ITooltipServiceWrapper,
    TooltipEnabledDataPoint,
    createTooltipServiceWrapper
} from "powerbi-visuals-utils-tooltiputils";

type D3Element =
    Selection<any>;

export class LegendMap implements IVisual {
    private static Properties: any = {
        dataPoint: {
            fill: <DataViewObjectPropertyIdentifier>{
                objectName: "dataPoint",
                propertyName: "fill"
            }
        },
        labels: {
            labelPrecision: <DataViewObjectPropertyIdentifier>{
                objectName: "labels",
                propertyName: "labelPrecision"
            }
        }
    };

    private host: IVisualHost;
    private colorHelper: ColorHelper;
    private localizationManager: ILocalizationManager;

    private tooltipServiceWrapper: ITooltipServiceWrapper;
    private svg: Selection<any>;
    private div: Selection<any>;
    private mainGraphics: Selection<any>;
    private colors: IColorPalette;
    private dataView: DataView;
    private viewport: IViewport;
    private margin: IMargin = { left: 5, right: 10, bottom: 5, top: 5 };

    private animationDuration: number = 1000;

    private static ClsAll: string = "*";
    private static ClsCategoryX: string = "categoryX";
    private static ClsMono: string = "mono";
    public static CLsLegendMapDataLabels = "legendMapDataLabels";
    private static ClsLegend: string = "legend";
    private static ClsBordered: string = "bordered";
    private static ClsNameSvgLegendMap: string = "svgLegendMap";
    private static ClsNameDivLegendMap: string = "divLegendMap";

    private static AttrTransform: string = "transform";
    private static AttrX: string = "x";
    private static AttrY: string = "y";
    private static AttrDX: string = "dx";
    private static AttrDY: string = "dy";
    private static AttrHeight: string = "height";
    private static AttrWidth: string = "width";

    private static HtmlObjTitle: string = "title";
    private static HtmlObjSvg: string = "svg";
    private static HtmlObjDiv: string = "div";
    private static HtmlObjG: string = "g";
    private static HtmlObjText: string = "text";
    private static HtmlObjRect: string = "rect";
    private static HtmlObjTspan: string = "tspan";

    private static StFill: string = "fill";
    private static StOpacity: string = "opacity";

    private static BucketCountMaxLimit: number = 18;
    private static BucketCountMinLimit: number = 1;
    private static ColorbrewerMaxBucketCount: number = 14;

    private static CellHeight: number = 15;

    private static DefaultColorbrewer: string = "Reds";

    private settings: Settings;

    private element: HTMLElement;

    public converter(dataView: DataView, colors: IColorPalette): LegendMapChartData {
        if (!dataView
            || !dataView.categorical
            || !dataView.categorical.categories
            || !dataView.categorical.categories[0]
            || !dataView.categorical.categories[0].values
            || !dataView.categorical.categories[0].values.length
            || !dataView.categorical.values
            || !dataView.categorical.values[0]
            || !dataView.categorical.values[0].values
            || !dataView.categorical.values[0].values.length
        ) {
            return <LegendMapChartData>{
                dataPoints: null
            };
        }

        let categoryValueFormatter: IValueFormatter;
        let valueFormatter: IValueFormatter;
        let dataPoints: LegendMapDataPoint[] = [];

        categoryValueFormatter = ValueFormatter.create({
            format: ValueFormatter.getFormatStringByColumn(dataView.categorical.categories[0].source),
            value: dataView.categorical.categories[0].values[0]
        });

        valueFormatter = ValueFormatter.create({
            format: ValueFormatter.getFormatStringByColumn(dataView.categorical.values[0].source),
            value: dataView.categorical.values[0].values[0]
        });

        // dataView.categorical.categories
        dataView.categorical.categories[0].values.forEach((categoryX, indexX) => {
            dataView.categorical.values.forEach((categoryY) => {
                let categoryYFormatter = ValueFormatter.create({
                    format: categoryY.source.format,
                    value: dataView.categorical.values[0].values[0]
                });
                let value = categoryY.values[indexX];
                dataPoints.push({
                    categoryX: categoryX,
                    categoryY: categoryY.source.displayName,
                    value: value,
                    valueStr: categoryYFormatter.format(value),
                    tooltipInfo: [{
                        displayName: `Category`,
                        value: (categoryX || "").toString()
                    },
                    {
                        displayName: `Y`,
                        value: (categoryY.source.displayName || "").toString()
                    },
                    {
                        displayName: `Value`,
                        value: categoryYFormatter.format(value)
                    }]
                });
            });
        });
        return <LegendMapChartData>{
            dataPoints: dataPoints,
            categoryX: dataView.categorical.categories[0].values.filter((n) => {
                return n !== undefined;
            }),
            categoryY: dataView.categorical.values.map(v => v.source.displayName).filter((n) => {
                return n !== undefined;
            }),
            categoryValueFormatter: categoryValueFormatter,
            valueFormatter: valueFormatter
        };
    }

    constructor({
        host,
        element
    }: VisualConstructorOptions) {
        this.host = host;
        this.element = element;

        this.div = d3.select(element)
            .append(LegendMap.HtmlObjDiv)
            .classed(LegendMap.ClsNameDivLegendMap, true);
        this.svg = this.div
            .append(LegendMap.HtmlObjSvg)
            .classed(LegendMap.ClsNameSvgLegendMap, true);

        this.colorHelper = new ColorHelper(this.host.colorPalette);
        this.localizationManager = host.createLocalizationManager();

        this.tooltipServiceWrapper = createTooltipServiceWrapper(
            this.host.tooltipService,
            element);
    }

    public update(options: VisualUpdateOptions): void {
        if (!options.dataViews || !options.dataViews[0]) {
            return;
        }
        try {
            this.host.eventService.renderingStarted(options);


            this.settings = LegendMap.parseSettings(options.dataViews[0], this.colorHelper);

            this.svg.selectAll(LegendMap.ClsAll).remove();
            this.div.attr("widtht", PixelConverter.toString(options.viewport.width + this.margin.left));
            this.div.style("height", PixelConverter.toString(options.viewport.height + this.margin.left));

            this.svg.attr("width", options.viewport.width);
            this.svg.attr("height", options.viewport.height);

            this.mainGraphics = this.svg.append(LegendMap.HtmlObjG);

            this.setSize(options.viewport);

            this.updateInternal(options, this.settings);
        } catch (ex) {
            this.host.eventService.renderingFailed(options, JSON.stringify(ex));
        }
        this.host.eventService.renderingFinished(options);
    }



    private static parseSettings(dataView: DataView, colorHelper: ColorHelper): Settings {
        let settings: Settings = Settings.parse<Settings>(dataView);
        if (!settings.general.enableColorbrewer) {
            if (settings.general.buckets > LegendMap.BucketCountMaxLimit) {
                settings.general.buckets = LegendMap.BucketCountMaxLimit;
            }
            if (settings.general.buckets < LegendMap.BucketCountMinLimit) {
                settings.general.buckets = LegendMap.BucketCountMinLimit;
            }
        } else {
            if (settings.general.colorbrewer === "") {
                settings.general.colorbrewer = LegendMap.DefaultColorbrewer;
            }
            let colorbrewerArray: IColorArray = colorbrewer[settings.general.colorbrewer];
            let minBucketNum: number = 0;
            let maxBucketNum: number = 0;
            for (let bucketIndex: number = LegendMap.BucketCountMinLimit; bucketIndex < LegendMap.ColorbrewerMaxBucketCount; bucketIndex++) {
                if (minBucketNum === 0 && (colorbrewerArray as Object).hasOwnProperty(bucketIndex.toString())) {
                    minBucketNum = bucketIndex;
                }
                if ((colorbrewerArray as Object).hasOwnProperty(bucketIndex.toString())) {
                    maxBucketNum = bucketIndex;
                }
            }

            if (settings.general.buckets > maxBucketNum) {
                settings.general.buckets = maxBucketNum;
            }
            if (settings.general.buckets < minBucketNum) {
                settings.general.buckets = minBucketNum;
            }
        }

        if (colorHelper.isHighContrast) {
            const foregroundColor: string = colorHelper.getThemeColor("foreground");
            const backgroundColor: string = colorHelper.getThemeColor("background");


            settings.general.enableColorbrewer = false;
            settings.general.gradientStart = backgroundColor;
            settings.general.gradientEnd = backgroundColor;
            settings.general.stroke = foregroundColor;
            settings.general.textColor = foregroundColor;
        }

        return settings;
    }

    private updateInternal(options: VisualUpdateOptions, settings: Settings): void {
        let dataView: DataView = this.dataView = options.dataViews[0];
        let chartData: LegendMapChartData = this.converter(dataView, this.colors);
        let suppressAnimations: boolean = false;
        if (chartData.dataPoints) {
            let minDataValue: number = d3.min(chartData.dataPoints, function (d: LegendMapDataPoint) {
                return d.value as number;
            });
            let maxDataValue: number = d3.max(chartData.dataPoints, function (d: LegendMapDataPoint) {
                return d.value as number;
            });

            let numBuckets: number = settings.general.buckets;
            let colorbrewerScale: string = settings.general.colorbrewer;
            let colorbrewerEnable: boolean = settings.general.enableColorbrewer;
            let colors: Array<string>;


            if (chartData.categoryX.length < numBuckets) {
                numBuckets = chartData.categoryX.length;
            }

            if (colorbrewerEnable) {
                if (colorbrewerScale) {
                    let currentColorbrewer: IColorArray = colorbrewer[colorbrewerScale];
                    colors = (currentColorbrewer ? currentColorbrewer[numBuckets] : colorbrewer.Reds[numBuckets]);
                }
                else {
                    colors = colorbrewer.Reds[numBuckets];	// default color scheme
                }
            } else {
                let startColor: string = settings.general.gradientStart;
                let endColor: string = settings.general.gradientEnd;
                let colorScale: LinearColorScale = createLinearColorScale([0, numBuckets], [startColor, endColor], true);
                colors = [];

                for (let bucketIndex: number = 0; bucketIndex < numBuckets; bucketIndex++) {
                    colors.push(colorScale(bucketIndex));
                }
            }

            let colorScale: Quantile<string> = d3.scaleQuantile<string>()
                .domain([minDataValue, maxDataValue])
                .range(colors);

            
            let gridSizeHeight: number = LegendMap.CellHeight;

            let xOffset: number = this.margin.left;

            const LegendMapCellRaito: number = 19/20;
            let legendElementWidth: number = (this.viewport.width * LegendMapCellRaito - xOffset) / numBuckets;



            let legendMapMain: Selection<LegendMapDataPoint> = this.mainGraphics.selectAll("." + LegendMap.ClsCategoryX);
            let legendMapData = legendMapMain
                .data(chartData.dataPoints);
            let legendMapEntered = legendMapData
                .enter()
                .append(LegendMap.HtmlObjRect);
            let legendMapMerged = legendMapEntered.merge(legendMapMain);

            let elementAnimation: Selection<D3Element> = <Selection<D3Element>>this.getAnimationMode(legendMapMerged, suppressAnimations);
            if (!this.settings.general.fillNullValuesCells) {
                legendMapMerged.style(LegendMap.StOpacity, function (d: any) {
                    return d.value === null ? 0 : 1;
                });
            }
            elementAnimation.style(LegendMap.StFill, function (d: any) {
                return <string>colorScale(d.value);
            });

            this.tooltipServiceWrapper.addTooltip(legendMapMerged, (tooltipEvent: TooltipEventArgs<TooltipEnabledDataPoint>) => {
                return tooltipEvent.data.tooltipInfo;
            });

            // legend
            let legendDataValues = [minDataValue].concat(colorScale.quantiles());
            let legendData = legendDataValues.concat(maxDataValue).map((value, index) => {
                return {
                    value: value,
                    tooltipInfo: [{
                        displayName: `Min value`,
                        value: value && typeof value.toFixed === "function" ? value.toFixed(0) : chartData.categoryValueFormatter.format(value)
                    },
                    {
                        displayName: `Max value`,
                        value: legendDataValues[index + 1] && typeof legendDataValues[index + 1].toFixed === "function" ? legendDataValues[index + 1].toFixed(0) : chartData.categoryValueFormatter.format(maxDataValue)
                    }]
                };
            });

            let legendSelection: Selection<any> = this.mainGraphics.selectAll("." + LegendMap.ClsLegend);
            let legendSelectionData = legendSelection.data(legendData);

            let legendSelectionEntered = legendSelectionData
                .enter()
                .append(LegendMap.HtmlObjG);

            legendSelectionData.exit().remove();

            let legendSelectionMerged = legendSelectionData.merge(legendSelection);
            legendSelectionMerged.classed(LegendMap.ClsLegend, true);

            let legendOffsetCellsY: number = this.margin.top;
            let legendOffsetTextY: number =  gridSizeHeight *2;

            legendSelectionEntered
                .append(LegendMap.HtmlObjRect)
                .attr(LegendMap.AttrX, function (d, i) {
                    return legendElementWidth * i + xOffset;
                })
                .attr(LegendMap.AttrY, legendOffsetCellsY)
                .attr(LegendMap.AttrWidth, legendElementWidth)
                .attr(LegendMap.AttrHeight, gridSizeHeight)
                .style(LegendMap.StFill, function (d, i) {
                    return colors[i];
                })
                .style("stroke", settings.general.stroke)
                .style("stroke-width", 0)
                .style("opacity", (d) => d.value !== maxDataValue ? 1 : 0)
                .classed(LegendMap.ClsBordered, true);

            legendSelectionEntered
                .append(LegendMap.HtmlObjText)
                .classed(LegendMap.ClsMono, true)
                .attr(LegendMap.AttrX, function (d, i) {
                    return legendElementWidth * i ;
                })
                .attr(LegendMap.AttrY, legendOffsetTextY)
                .text(function (d) {
                    return chartData.valueFormatter.format(d.value);
                })
                .style("fill", settings.general.textColor);

                this.tooltipServiceWrapper.addTooltip(
                    legendSelectionEntered,
                    (tooltipEvent: TooltipEventArgs<TooltipEnabledDataPoint>) => {
                        return tooltipEvent.data.tooltipInfo;
                    }
                );
        }
    }

    private static textLimit(text: string, limit: number) {
        if (text.length > limit) {
            return ((text || "").substring(0, limit - 3).trim()) + "…";
        }

        return text;
    }

    private setSize(viewport: IViewport): void {
        let height: number,
            width: number;

        this.svg
            .attr(LegendMap.AttrHeight, Math.max(viewport.height, 0))
            .attr(LegendMap.AttrWidth, Math.max(viewport.width, 0));

        height =
            viewport.height -
            this.margin.top -
            this.margin.bottom;

        width =
            viewport.width -
            this.margin.left -
            this.margin.right;

        this.viewport = {
            height: height,
            width: width
        };

        this.mainGraphics
            .attr(LegendMap.AttrHeight, Math.max(this.viewport.height + this.margin.top, 0))
            .attr(LegendMap.AttrWidth, Math.max(this.viewport.width + this.margin.left, 0));

        this.mainGraphics.attr(LegendMap.AttrTransform, translate(this.margin.left, this.margin.top));
    }

    private truncateTextIfNeeded(text: Selection<any>, width: number): void {
        text.call(LabelLayoutStrategy.clip,
            width,
            TextMeasurementService.svgEllipsis);
    }

    private wrap(text, width): void {
        text.each(function () {
            let text: Selection<D3Element> = d3.select(this);
            let words: string[] = text.text().split(/\s+/).reverse();
            let word: string;
            let line: string[] = [];
            let lineNumber: number = 0;
            let lineHeight: number = 1.1; // ems
            let x: string = text.attr(LegendMap.AttrX);
            let y: string = text.attr(LegendMap.AttrY);
            let dy: number = parseFloat(text.attr(LegendMap.AttrDY));
            let tspan: Selection<any> = text.text(null).append(LegendMap.HtmlObjTspan).attr(LegendMap.AttrX, x).attr(LegendMap.AttrY, y).attr(LegendMap.AttrDY, dy + "em");
            while (word = words.pop()) {
                line.push(word);
                tspan.text(line.join(" "));
                let tspannode: any = tspan.node();  // Fixing Typescript error: Property 'getComputedTextLength' does not exist on type 'Element'.
                if (tspannode.getComputedTextLength() > width) {
                    line.pop();
                    tspan.text(line.join(" "));
                    line = [word];
                    tspan = text.append(LegendMap.HtmlObjTspan).attr(LegendMap.AttrX, x).attr(LegendMap.AttrY, y).attr(LegendMap.AttrDY, ++lineNumber * lineHeight + dy + "em").text(word);
                }
            }
        });
    }

    private getAnimationMode(element: D3Element, suppressAnimations: boolean): D3Element {
        if (suppressAnimations) {
            return element;
        }

        return (<any>element)
            .transition().duration(this.animationDuration);
    }

    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
        const settings: Settings = this.dataView && this.settings
            || Settings.getDefault() as Settings;

        const instanceEnumeration: VisualObjectInstanceEnumeration =
            Settings.enumerateObjectInstances(settings, options);

        return instanceEnumeration || [];
    }
}
