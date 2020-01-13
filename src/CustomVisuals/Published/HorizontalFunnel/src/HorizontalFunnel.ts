/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ''Software''), to deal
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

// tslint:disable-next-line: no-namespace tslint:disable-next-line: no-internal-module
module powerbi.extensibility.visual {
    // color
    import ILegend = powerbi.extensibility.utils.chart.legend.ILegend;
    import IValueFormatter = powerbi.extensibility.utils.formatting.IValueFormatter;
    import valueFormatter = powerbi.extensibility.utils.formatting.valueFormatter;
    // tooltip
    import textMeasurementService = powerbi.extensibility.utils.formatting.textMeasurementService;
    import TextProperties = powerbi.extensibility.utils.formatting.TextProperties;
    import ISelectionManager = powerbi.extensibility.ISelectionManager;
    import IVisual = powerbi.extensibility.visual.IVisual;
    import IVisualHost = powerbi.extensibility.visual.IVisualHost;
    import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
    import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
    import DataViewObjects = powerbi.extensibility.utils.dataview.DataViewObjects;
    import ISelectionId = powerbi.visuals.ISelectionId;
    import tooltip = powerbi.extensibility.utils.tooltip;
    import TooltipEventArgs = powerbi.extensibility.utils.tooltip.TooltipEventArgs;
    import TooltipEnabledDataPoint = powerbi.extensibility.utils.tooltip.TooltipEnabledDataPoint;
    import createTooltipServiceWrapper = powerbi.extensibility.utils.tooltip.createTooltipServiceWrapper;
    // tooltip
    import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
    import ITooltipServiceWrapper = powerbi.extensibility.utils.tooltip.ITooltipServiceWrapper;
    // color
    import IColorPalette = powerbi.extensibility.IColorPalette;
    import ISelectionIdBuilder = powerbi.visuals.ISelectionIdBuilder;
    // label
    import utils = powerbi.extensibility.utils.chart.dataLabel.utils;
    import DataLabelObject = powerbi.extensibility.utils.chart.dataLabel.DataLabelObject;

    interface ITooltipService {
        enabled(): boolean;
        show(options: TooltipShowOptions): void;
        move(options: TooltipMoveOptions): void;
        hide(options: TooltipHideOptions): void;
    }
    interface ITooltipDataItem {
        displayName: string;
        value: string;
    }
    const cardProps: any = {
        categoryLabels: {
            color: <DataViewObjectPropertyIdentifier>{
                objectName: "categoryLabels",
                propertyName: "color"
            },
            fontSize: <DataViewObjectPropertyIdentifier>{
                objectName: "categoryLabels",
                propertyName: "fontSize"
            },
            show: <DataViewObjectPropertyIdentifier>{
                objectName: "categoryLabels",
                propertyName: "show"
            }
        },
        labels: {
            color: <DataViewObjectPropertyIdentifier>{
                objectName: "labels",
                propertyName: "color"
            },
            fontSize: <DataViewObjectPropertyIdentifier>{
                objectName: "labels",
                propertyName: "fontSize"
            },
            labelDisplayUnits: <DataViewObjectPropertyIdentifier>{
                objectName: "labels",
                propertyName: "labelDisplayUnits"
            },
            labelPrecision: <DataViewObjectPropertyIdentifier>{
                objectName: "labels",
                propertyName: "labelPrecision"
            },
        },
        trendlabels: {
            color: <DataViewObjectPropertyIdentifier>{
                objectName: "trendlabels",
                propertyName: "color"
            },
            fontSize: <DataViewObjectPropertyIdentifier>{
                objectName: "trendlabels",
                propertyName: "fontSize"
            },
            labelDisplayUnits: <DataViewObjectPropertyIdentifier>{
                objectName: "trendlabels",
                propertyName: "labelDisplayUnits"
            },
            labelPrecision: <DataViewObjectPropertyIdentifier>{
                objectName: "trendlabels",
                propertyName: "labelPrecision"
            },
        },
        wordWrap: {
            show: <DataViewObjectPropertyIdentifier>{
                objectName: "wordWrap",
                propertyName: "show"
            }
        }
    };
    interface ITooltipDataPoints {
        name: string;
        value: string;
        formatter: string;
    }

    interface ILabelSettings {
        color: string;
        displayUnits: number;
        decimalPlaces: number;
        fontSize: number;
    }

    interface IFunnelTitle {
        show: boolean;
        titleText: string;
        tooltipText: string;
        color: string;
        bkColor: string;
        fontSize: number;
    }

    interface ISortSettings {
        sortBy: string;
        orderBy: string;
    }
    interface IShowConnectorsSettings {
        show: boolean;
    }
    interface IShowLegendSettings {
        show: boolean;
    }
    interface ICardFormatSetting {
        showTitle: boolean;
        textSize: number;
        labelSettings: any; // VisualDataLabelsSettings;
        wordWrap: boolean;
    }
    interface IFunnelViewModel {
        primaryColumn?: string;
        secondaryColumn?: string;
        count?: number;
        toolTipInfo?: ITooltipDataItem[];
        categories?: ICategoryViewModel[];
        values?: IValueViewModel[];
        identity?: ISelectionId;
    }
    interface ICategoryViewModel {
        value: string;
        color: string;
        identity?: ISelectionId;
    }
    interface IValueViewModel {
        // tslint:disable-next-line:no-any
        values: any;
    }
    // tslint:disable-next-line:no-any
    const horizontalFunnelProps: any = {
        LabelSettings: {
            color: <DataViewObjectPropertyIdentifier>{ objectName: "labels", propertyName: "color" },
            fontSize: <DataViewObjectPropertyIdentifier>{ objectName: "labels", propertyName: "fontSize" },
            labelDisplayUnits: <DataViewObjectPropertyIdentifier>{
                objectName: "labels",
                propertyName: "labelDisplayUnits"
            },
            labelPrecision: <DataViewObjectPropertyIdentifier>{
                objectName: "labels",
                propertyName: "labelPrecision"
            }
        },
        ShowConnectors: {
            show: <DataViewObjectPropertyIdentifier>{ objectName: "ShowConnectors", propertyName: "show" }
        },
        ShowLegend: {
            show: <DataViewObjectPropertyIdentifier>{ objectName: "ShowLegend", propertyName: "show" }
        },
        dataPoint: {
            defaultColor: <DataViewObjectPropertyIdentifier>{
                objectName: "dataPoint",
                propertyName: "defaultColor"
            },
            fill: <DataViewObjectPropertyIdentifier>{
                objectName: "dataPoint",
                propertyName: "fill"
            }
        },
        funnelTitle: {
            show: <DataViewObjectPropertyIdentifier>{ objectName: "FunnelTitle", propertyName: "show" },
            titleBackgroundColor: { objectName: "FunnelTitle", propertyName: "backgroundColor" },
            titleFill: { objectName: "FunnelTitle", propertyName: "fill1" },
            titleFontSize: { objectName: "FunnelTitle", propertyName: "fontSize" },
            titleText: <DataViewObjectPropertyIdentifier>{ objectName: "FunnelTitle", propertyName: "titleText" },
            tooltipText: { objectName: "FunnelTitle", propertyName: "tooltipText" }
        },
        show: { objectName: "FunnelTitle", propertyName: "show" },
        sort: {
            OrderBy: <DataViewObjectPropertyIdentifier>{ objectName: "Sort", propertyName: "OrderBy" },
            SortBy: <DataViewObjectPropertyIdentifier>{ objectName: "Sort", propertyName: "SortBy" },
            orderBy: <DataViewObjectPropertyIdentifier>{ objectName: "Sort", propertyName: "OrderBy" },
            sortBy: <DataViewObjectPropertyIdentifier>{ objectName: "Sort", propertyName: "SortBy" },
        },
        titleBackgroundColor: { objectName: "FunnelTitle", propertyName: "backgroundColor" },
        titleFill: { objectName: "FunnelTitle", propertyName: "fill1" },
        titleFontSize: { objectName: "FunnelTitle", propertyName: "fontSize" },
        titleText: { objectName: "FunnelTitle", propertyName: "titleText" },
        tooltipText: { objectName: "FunnelTitle", propertyName: "tooltipText" },

    };
    // tslint:disable-next-line:no-any
    const sortType: any = [
        { value: "Auto", displayName: "Auto" },
        { value: "Series", displayName: "Series" },
        { value: "PrimaryMeasure", displayName: "Primary Measure" },
        { value: "SecondaryMeasure", displayName: "Secondary Measure" }];
    // tslint:disable-next-line:no-any
    const orderType: any = [
        { value: "ascending", displayName: "Ascending", description: "Ascending" },
        { value: "descending", displayName: "Descending", description: "Descending" }];

    function getCategoricalObjectValue<T>(
        category: DataViewCategoryColumn, index: number, objectName: string, propertyName: string, defaultValue: T): T {
        let categoryObjects: DataViewObjects[];
        categoryObjects = category.objects;
        if (categoryObjects) {
            const categoryObject: DataViewObject = categoryObjects[index];
            if (categoryObject) {
                let object: DataViewPropertyValue;
                object = categoryObject[objectName];
                if (object) {
                    let property: T;
                    property = object[propertyName];
                    if (property !== undefined) {
                        return property;
                    }
                }
            }
        }

        return defaultValue;
    }
    /**
     *  It is used to present the visual on screen
     */

    export class HorizontalFunnel implements IVisual {
        public static GETDEFAULTDATA(): IFunnelViewModel {
            return {
                categories: [],
                count: 0,
                identity: null,
                primaryColumn: "",
                secondaryColumn: "",
                toolTipInfo: [],
                values: []
            };
        }
        private static visPrint(categories, series, cat, formatStringProp, categorical, viewModel, xAxisSortedValues, yAxis1SortedValues, catDv, yvalueIndex, yAxis2SortedValues, host, unsortcategoriesvalues, objects, unsortindex, unsortcategories) {
            if (categories && series && categories.length > 0 && series.length > 0) {
                let categorySourceFormatString: string;
                categorySourceFormatString = valueFormatter.getFormatString(cat.source, formatStringProp);
                let toolTipItems: any;
                toolTipItems = [];
                let formattedCategoryValue: any;
                let value: any;
                let categoryColumn: DataViewCategoryColumn;
                categoryColumn = categorical.categories[0];
                let catLength: number;
                catLength = xAxisSortedValues.length;
                for (let iLoop: number = 0; iLoop < catLength; iLoop++) {
                    toolTipItems = [];
                    if (iLoop !== 0) {
                        viewModel.push({ toolTipInfo: [] });
                    }
                    viewModel[0].categories.push({
                        color: "", value: xAxisSortedValues[iLoop]
                    });
                    let decimalPlaces: number = HorizontalFunnel.GETDECIMALPLACECOUNT(yAxis1SortedValues[iLoop]);
                    decimalPlaces = decimalPlaces > 4 ? 4 : decimalPlaces;
                    let primaryFormat: string;
                    primaryFormat = series && series[0] && series[0].source && series[0].source.format ? series[0].source.format : "";
                    let formatter: IValueFormatter;
                    formatter = valueFormatter.create({
                        format: primaryFormat, precision: decimalPlaces, value: 0
                    });
                    formattedCategoryValue = valueFormatter.format(xAxisSortedValues[iLoop], categorySourceFormatString);
                    let tooltipInfo: ITooltipDataItem[] = [];
                    let tooltipItem1: ITooltipDataItem = { displayName: "", value: "" };
                    let tooltipItem2: ITooltipDataItem = { displayName: "", value: "" };
                    tooltipItem1.displayName = catDv.categories["0"].source.displayName;
                    tooltipItem1.value = formattedCategoryValue;
                    tooltipInfo.push(tooltipItem1);
                    tooltipItem2.displayName = catDv.values["0"].source.displayName;
                    let formattedTooltip: string = formatter.format(Math.round(yAxis1SortedValues[iLoop] * 100) / 100);
                    tooltipItem2.value = formattedTooltip;
                    tooltipInfo.push(tooltipItem2);
                    value = Math.round(yAxis1SortedValues[iLoop] * 100) / 100;
                    viewModel[0].values.push({ values: [] });
                    viewModel[0].values[iLoop].values.push(value);
                    if (yvalueIndex !== undefined) {
                        value = yAxis2SortedValues[iLoop];
                        viewModel[0].values[iLoop].values.push(value);
                    }
                    viewModel[iLoop].toolTipInfo = tooltipInfo;
                    let x: any = viewModel[0].values[iLoop];
                    x.toolTipInfo = tooltipInfo;
                }
                let colorPalette: IColorPalette;
                colorPalette = host.colorPalette;
                // create object for colors
                let colorObj: {} = {};
                for (let i: number = 0; i < catLength; i++) {
                    let currentElement: string = categoryColumn.values[i].toString();
                    let defaultColor: Fill;
                    defaultColor = {
                        solid: {
                            color: colorPalette.getColor(currentElement).value
                        }
                    };
                    colorObj[currentElement] = getCategoricalObjectValue<Fill>(
                        categoryColumn, i, "dataPoint", "fill", defaultColor).solid.color;
                }
                for (let iLoop: number = 0; iLoop < catLength; iLoop++) {
                    for (let unLoop: number = 0; unLoop < catLength; unLoop++) {
                        if (unsortcategoriesvalues[unLoop] === xAxisSortedValues[iLoop]) {
                            objects = categoryColumn.objects && categoryColumn.objects[unLoop];
                            let dataPointObject: any;
                            if (objects) {
                                dataPointObject = categoryColumn.objects[unLoop];
                            }
                            let color: any;
                            if (objects && dataPointObject && dataPointObject.dataPoint && dataPointObject.dataPoint.fill && dataPointObject.dataPoint.fill.solid.color) {
                                color = { value: dataPointObject.dataPoint.fill.solid.color };
                            } else {
                                let currentElement: string;
                                const colorPlt: any = colorPalette;
                                currentElement = categoryColumn.values[iLoop].toString();
                                color = colorPlt.colorPalette[currentElement];
                            }
                            unsortindex = unLoop;
                            let categorySelectionId: ISelectionId;
                            categorySelectionId = host.createSelectionIdBuilder()
                                .withCategory(unsortcategories, unsortindex).createSelectionId();
                            viewModel[iLoop].identity = categorySelectionId;
                            viewModel[0].categories[iLoop].identity = categorySelectionId;
                            viewModel[0].categories[iLoop].color = color;
                            break;
                        }
                    }
                }
                viewModel[0].count = catLength;
            }
        }
        private static visualPrinter(caseVar, catVal, arrTextValuesSortIndexes, categorical, arrTempYAxisValues1, arrTempXAxisValues, viewModel, unsortsecondaryvalues, order) {
            if (caseVar.iTotalXAxisNumericValues) {
                for (let iCounter: number = 0; iCounter < caseVar.arrValuesToBeSorted.length; iCounter++) {
                    if (-1 === caseVar.arrIntegerValuesSortIndexes.indexOf(iCounter)
                        && -1 === arrTextValuesSortIndexes.textIndex.indexOf(iCounter)) {
                        caseVar.iSmallestValue = caseVar.arrValuesToBeSorted[iCounter];
                        caseVar.iIndex = iCounter;
                        if (isNaN(parseFloat(caseVar.iSmallestValue))) {
                            arrTextValuesSortIndexes.textValue.push(caseVar.iSmallestValue);
                            arrTextValuesSortIndexes.textIndex.push(iCounter);
                            continue;
                        } else {
                            for (let iInnerCount: number = iCounter + 1;
                                iInnerCount < caseVar.arrValuesToBeSorted.length; iInnerCount++) {
                                if (!isNaN(caseVar.arrValuesToBeSorted[iInnerCount])
                                    && -1 === caseVar.arrIntegerValuesSortIndexes.indexOf(iInnerCount)
                                    && null !== caseVar.arrValuesToBeSorted[iInnerCount]) {
                                    if (caseVar.arrValuesToBeSorted[iInnerCount] < caseVar.iSmallestValue) {
                                        caseVar.iIndex = iInnerCount;
                                        caseVar.iSmallestValue = caseVar.arrValuesToBeSorted[iInnerCount];
                                    }
                                }
                            }
                            caseVar.arrIntegerValuesSortIndexes.push(caseVar.iIndex);
                        }
                        if (-1 === caseVar.arrIntegerValuesSortIndexes.indexOf(iCounter)) {
                            iCounter--;
                        }
                    }
                }
                for (const iLoop of caseVar.arrIntegerValuesSortIndexes) {
                    caseVar.xAxisSortedIntegerValues.push(
                        arrTempXAxisValues[iLoop]);
                    for (let iNumberOfYAxisParameters: number = 0;
                        iNumberOfYAxisParameters < categorical.values.length;
                        iNumberOfYAxisParameters++) {
                        if (0 === iNumberOfYAxisParameters) {
                            caseVar.yAxisAutoSort.push(
                                arrTempYAxisValues1[iLoop]);
                        } else if (1 === iNumberOfYAxisParameters) {
                            caseVar.yAxis1AutoSort.push(
                                caseVar.arrTempYAxisValues2[iLoop]);
                        }
                    }
                }
                for (let iLoop: number = 0; iLoop < arrTextValuesSortIndexes.textValue.length;
                    iLoop++) {
                    caseVar.xAxisSortedIntegerValues.push(
                        arrTempXAxisValues[arrTextValuesSortIndexes.textIndex[iLoop]]);
                    for (let iNumberOfYAxisParameters: number = 0;
                        iNumberOfYAxisParameters < categorical.values.length;
                        iNumberOfYAxisParameters++) {
                        if (0 === iNumberOfYAxisParameters) {
                            caseVar.yAxisAutoSort.push(
                                arrTempYAxisValues1[arrTextValuesSortIndexes.textIndex[iLoop]]);
                        } else if (1 === iNumberOfYAxisParameters) {
                            caseVar.yAxis1AutoSort.push(
                                caseVar.arrTempYAxisValues2[arrTextValuesSortIndexes.textIndex[iLoop]]);
                        }
                    }
                }
            } else {
                caseVar.xAxisSortedIntegerValues = JSON.parse(JSON.stringify(catVal.unsortcategoriesvalues));
                caseVar.yAxisAutoSort = JSON.parse(JSON.stringify(catVal.unsorttargetvalues));
                if (viewModel[0].secondaryColumn) {
                    caseVar.yAxis1AutoSort = JSON.parse(JSON.stringify(unsortsecondaryvalues));
                }
            }
            if (order === "descending") {
                for (let iCount: number = caseVar.xAxisSortedIntegerValues.length - 1;
                    iCount >= 0; iCount--) {
                    catVal.xAxisSortedValues.push(caseVar.xAxisSortedIntegerValues[iCount]);
                    catVal.yAxis1SortedValues.push(caseVar.yAxisAutoSort[iCount]);
                    if (viewModel[0].secondaryColumn) {
                        catVal.yAxis2SortedValues.push(caseVar.yAxis1AutoSort[iCount]);
                    }
                }
            } else {
                catVal.xAxisSortedValues = JSON.parse(JSON.stringify(caseVar.xAxisSortedIntegerValues));
                catVal.yAxis1SortedValues = JSON.parse(JSON.stringify(caseVar.yAxisAutoSort));
                catVal.yAxis2SortedValues = JSON.parse(JSON.stringify(caseVar.yAxis1AutoSort));
            }
        }
        private static loop(categorical, caseVar, value, iValueToBeSorted) {
            for (const iCount of categorical.categories[0].values) {
                value = iCount;
                if (isNaN(value)) {
                    iValueToBeSorted =
                        iCount.toString().match(/-?\d+\.?\d*/);
                    if (null !== iValueToBeSorted) {
                        caseVar.arrValuesToBeSorted.push(parseFloat(iValueToBeSorted[0]));
                        caseVar.iTotalXAxisNumericValues++;
                    } else caseVar.arrValuesToBeSorted.push(value);
                } else {
                    value = iCount;
                    if (isNaN(parseFloat(value))) caseVar.arrValuesToBeSorted.push(value);
                    else caseVar.arrValuesToBeSorted.push(parseFloat(value));
                    caseVar.iTotalXAxisNumericValues++;
                }
            }
        }
        //AK47
        private static caseSeries(order, catVal, categorical, targetvalueIndex, yvalueIndex, viewModel, index) {
            if (order === "ascending") catVal.xAxisSortedValues = categorical.categories[0].values.sort(d3.ascending);
            else catVal.xAxisSortedValues = categorical.categories[0].values.sort(d3.descending);
            for (const iCount of catVal.xAxisSortedValues) {
                const temp: any = iCount;
                for (index = 0; index < catVal.unsortcategoriesvalues.length; index++) {
                    if (temp === catVal.unsortcategoriesvalues[index]) {
                        catVal.yAxis1SortedValues.push(categorical.values[targetvalueIndex].values[index]);
                        if (viewModel[0].secondaryColumn) {
                            catVal.yAxis2SortedValues.push(categorical.values[yvalueIndex].values[index]);
                        }
                        break;
                    }
                }
            }
        }
        private static casePrimaryMeasure(order, index, catVal, categorical, targetvalueIndex, unsortsecondaryvalues, viewModel) {
            if (order === "ascending") catVal.yAxis1SortedValues = catVal.unsorttargetvalues.sort(d3.ascending);
            else catVal.yAxis1SortedValues = catVal.unsorttargetvalues.sort(d3.descending);
            for (const iCount of catVal.yAxis1SortedValues) {
                const temp: any = iCount;
                for (index = 0; index < categorical.values[targetvalueIndex].values.length; index++) {
                    if (temp === categorical.values[targetvalueIndex].values[index]) {
                        if (catVal.xAxisSortedValues.indexOf(catVal.unsortcategoriesvalues[index]) > -1) {
                            continue;
                        } else {
                            catVal.xAxisSortedValues.push(catVal.unsortcategoriesvalues[index]);
                            if (viewModel[0].secondaryColumn) {
                                catVal.yAxis2SortedValues.push(unsortsecondaryvalues[index]);
                            }
                            break;
                        }
                    }
                }
            }
        }
        public static CONVERTER(
            dataView: DataView,
            colors: IColorPalette,
            sort: string,
            order: string,
            host: IVisualHost): IFunnelViewModel[] {
            let viewModel: IFunnelViewModel[];
            viewModel = [];
            viewModel.push(HorizontalFunnel.GETDEFAULTDATA());
            if (dataView) {
                let objects: DataViewObjects = dataView.metadata.objects;
                let targetvalueIndex: number;
                let yvalueIndex: number;
                if (!dataView || !dataView.categorical || !dataView.categorical.values || !dataView.categorical.categories) {
                    viewModel[0].count = -1;
                    return viewModel;
                }
                for (let iLoop: number = 0; iLoop < dataView.categorical.values.length; iLoop++) {
                    if (dataView.categorical.values[iLoop].source.roles && dataView.categorical.values[iLoop].source.roles.hasOwnProperty("primaryMeasure")) {
                        targetvalueIndex = iLoop;
                        viewModel[0].primaryColumn = dataView.categorical.values[iLoop].source.displayName;
                    } else if (dataView.categorical.values[iLoop].source.roles && dataView.categorical.values[iLoop].source.roles.hasOwnProperty("secondaryMeasure")) {
                        yvalueIndex = iLoop;
                        viewModel[0].secondaryColumn = dataView.categorical.values[iLoop].source.displayName;
                    }
                }
                if (targetvalueIndex !== undefined) {
                    let categorical: DataViewCategorical;
                    categorical = dataView.categorical;
                    if (categorical) {
                        let unsortsecondaryvalues: any;
                        let catVal = {
                            unsortcategoriesvalues: JSON.parse(JSON.stringify(categorical.categories[0].values)), unsortcategories: categorical.categories[0], unsorttargetvalues: JSON.parse(JSON.stringify(categorical.values[targetvalueIndex].values)),
                            unsortindex: 0, yAxis1SortedValues: [], yAxis2SortedValues: [], xAxisSortedValues: []
                        }
                        if (viewModel[0].secondaryColumn) unsortsecondaryvalues = JSON.parse(JSON.stringify(categorical.values[yvalueIndex].values));
                        switch (sort) {
                            case "Auto":
                                let caseVar = {
                                    arrValuesToBeSorted: [], iSmallestValue: '', iIndex: 0, arrTempYAxisValues2: [], arrIntegerValuesSortIndexes: [],
                                    arrTextValuesSortIndexes: [], yAxisAutoSort: [], yAxis1AutoSort: [], xAxisSortedIntegerValues: [], iTotalXAxisNumericValues: 0
                                };
                                let iValueToBeSorted: any;
                                let arrTempXAxisValues: PrimitiveValue[] = categorical.categories[0].values;
                                let arrTempYAxisValues1: PrimitiveValue[] = categorical.values[0].values;
                                let arrTextValuesSortIndexes: any = { textIndex: [], textValue: [] };
                                if (2 === categorical.values.length) caseVar.arrTempYAxisValues2 = categorical.values[1].values;
                                // Change of value
                                let value: any;
                                this.loop(categorical, caseVar, value, iValueToBeSorted);
                                this.visualPrinter(caseVar, catVal, arrTextValuesSortIndexes, categorical, arrTempYAxisValues1, arrTempXAxisValues, viewModel, unsortsecondaryvalues, order);
                                break;
                            case "Series":
                                let index: number;
                                this.caseSeries(order, catVal, categorical, targetvalueIndex, yvalueIndex, viewModel, index);//AK47
                                break;
                            case "PrimaryMeasure":
                                this.casePrimaryMeasure(order, index, catVal, categorical, targetvalueIndex, unsortsecondaryvalues, viewModel); //AK47
                                break;
                            case "SecondaryMeasure":
                                if (order === "ascending") catVal.yAxis2SortedValues = unsortsecondaryvalues.sort(d3.ascending);
                                else catVal.yAxis2SortedValues = unsortsecondaryvalues.sort(d3.descending);
                                for (const iCount of catVal.yAxis2SortedValues) {
                                    let temp: any = iCount;
                                    for (index = 0; index < categorical.values[yvalueIndex].values.length; index++) {
                                        if (temp === categorical.values[yvalueIndex].values[index]) {
                                            if (catVal.xAxisSortedValues.indexOf(catVal.unsortcategoriesvalues[index]) > -1) continue;
                                            else {
                                                catVal.xAxisSortedValues.push(catVal.unsortcategoriesvalues[index]);
                                                if (viewModel[0].primaryColumn) catVal.yAxis1SortedValues.push(catVal.unsorttargetvalues[index]);
                                                break;
                                            }
                                        }
                                    }
                                }
                                break;
                            default:
                                break;
                        }
                        let categories: DataViewCategoryColumn[] = categorical.categories;
                        let series: DataViewValueColumns = categorical.values;
                        let catDv: DataViewCategorical = dataView.categorical;
                        let cat: DataViewCategoryColumn = catDv.categories[0];
                        let formatStringProp: DataViewObjectPropertyIdentifier;
                        formatStringProp = <DataViewObjectPropertyIdentifier>{
                            objectName: "general", propertyName: "formatString"
                        };
                        this.visPrint(categories, series, cat, formatStringProp, categorical, viewModel, catVal.xAxisSortedValues, catVal.yAxis1SortedValues, catDv, yvalueIndex, catVal.yAxis2SortedValues,
                            host, catVal.unsortcategoriesvalues, objects, catVal.unsortindex, catVal.unsortcategories);
                    }
                }
            }
            return viewModel;
        }

        public static GETDECIMALPLACECOUNT(value: number): number {
            let decimalPlaces: number = 0;
            let splitArr: string[];
            splitArr = value ? value.toString().split(".") : [];
            if (splitArr[1]) {
                decimalPlaces = splitArr[1].length;
            }

            return decimalPlaces;
        }

        private static minOpacity: number = 0.3;
        private static maxOpacity: number = 1;

        public host: IVisualHost;
        private events: IVisualEventService;
        private tooltipServiceWrapper: ITooltipServiceWrapper;
        private root: d3.Selection<SVGElement>;
        private dataView: DataView;
        private colors: IColorPalette;
        private cardFormatSetting: ICardFormatSetting;
        private durationAnimations: number = 200;
        private selectionManager: ISelectionManager;
        // tslint:disable-next-line:no-any
        private defaultDataPointColor: any = undefined;
        private tooltipInfoValue: string;
        // tslint:disable-next-line:no-any
        private viewModel: any = undefined;

        constructor(options: VisualConstructorOptions) {
            this.events = options.host.eventService;
            this.host = options.host;
            this.root = d3.select(options.element).style("cursor", "default");
            horizontalFunnelProps.cPalette = options.host.colorPalette;
            this.tooltipServiceWrapper = createTooltipServiceWrapper(this.host.tooltipService, options.element);
            this.selectionManager = options.host.createSelectionManager();
            // function call to handle selections on bookmarks
            this.selectionManager.registerOnSelectCallback(() => {
                // tslint:disable-next-line:no-any
                const selection: any = this.root.selectAll(".hf_datapoint");
                this.syncSelectionState(selection, <ISelectionId[]>this.selectionManager.getSelectionIds());
            });
        }
        private printerloop(showLegendProp, showConnectorsProp, element, color, updateVar) {
            for (let i: number = 0; i < (2 * updateVar.catLength - 1); i++) {
                if (!showLegendProp.show) {
                    let constantMultiplier: number = 1;
                    const xDisplacement: number = 10;
                    if (updateVar.catLength > 0) {
                        constantMultiplier = 4 / (5 * updateVar.catLength - 1); // dividing the available space into equal parts
                    }
                    updateVar.width = (updateVar.parentWidth - xDisplacement) * constantMultiplier;
                    // remove 10 from total width as 10 is x displacement
                    if (!showConnectorsProp.show) {
                        updateVar.width = (updateVar.parentWidth - xDisplacement) / updateVar.catLength;
                    }
                }
                element = this.root.select(".hf_parentdiv")
                    .append("div")
                    .style({
                        height: `${updateVar.height}px`
                    })
                    .classed("hf_svg hf_parentElement", true);
                if (i % 2 === 0) {
                    updateVar.classname = `hf_odd${i}`;
                    element
                        .append("div")
                        .style({
                            color,
                            "font-size": `${updateVar.fontsize}px`,
                            "width": `${(0.92 * updateVar.width)}px`
                        })
                        .classed(`hf_legend_item${i} hf_xAxisLabels`, true)
                        .classed("hf_legend", true);
                    element.append("div")
                        .style({
                            color,
                            "font-size": `${updateVar.fontsize}px`,
                            "width": `${0.92 * updateVar.width}px`
                        })
                        .classed(`hf_legend_value1${i}`, true)
                        .classed("hf_legend", true);
                    let displacement: number;
                    displacement = i === 0 ? 0 : 10;
                    element
                        .append("svg")
                        .attr({
                            fill: "#FAFAFA",
                            height: updateVar.height,
                            id: i,
                            overflow: "hidden",
                            width: updateVar.width
                        }).classed(updateVar.classname, true)
                        .append("rect")
                        .attr({
                            height: updateVar.height,
                            width: updateVar.width,
                            x: 10,
                            y: 0
                        });
                    if (this.viewModel.secondaryColumn) {
                        element.append("div")
                            .style({
                                color,
                                width: `${0.92 * updateVar.width}px`
                            })
                            .classed(`hf_legend_value2${i}`, true)
                            .style({ "font-size": `${updateVar.fontsize}px` })
                            .classed("hf_yaxis2", true);
                    }
                } else {
                    updateVar.classname = `hf_even${i}`;
                    let disp: number;
                    disp = 10;
                    if (showConnectorsProp.show) {
                        element
                            .append("svg")
                            .attr({
                                fill: "#FAFAFA",
                                height: updateVar.height,
                                id: i,
                                overflow: "hidden",
                                width: updateVar.width / 4,

                            })
                            .classed(updateVar.classname, true)
                            .append("rect")
                            .attr({
                                height: updateVar.height,
                                width: updateVar.width / 4,
                                x: disp,
                                y: 0
                            });
                    }
                }
            }
        }

        private dataViewMeta(dataViewMetadata, updateVar, dataView,defaultDataPointColor):any {
            if (dataViewMetadata) {
                let objects: DataViewObjects = dataViewMetadata.objects;
                if (objects) {
                    const labelString: string = "labels";
                    updateVar.labelSettings = this.cardFormatSetting.labelSettings;
                    updateVar.labelSettings.labelColor = DataViewObjects.getFillColor(objects, cardProps.labels.color, updateVar.labelSettings.labelColor);
                    updateVar.labelSettings.precision = DataViewObjects.getValue(objects, cardProps.labels.labelPrecision, updateVar.labelSettings.precision);
                    // The precision can't go below 0
                    if (updateVar.labelSettings.precision) {
                        updateVar.labelSettings.precision = (updateVar.labelSettings.precision >= 0) ?
                            ((updateVar.labelSettings.precision <= 4) ? updateVar.labelSettings.precision : 4) : 0;
                        updateVar.precisionValue = updateVar.labelSettings.precision;
                    }
                    defaultDataPointColor = DataViewObjects.getFillColor(objects, horizontalFunnelProps.dataPoint.defaultColor);
                    updateVar.labelSettings.displayUnits = DataViewObjects.getValue(objects, cardProps.labels.labelDisplayUnits, updateVar.labelSettings.displayUnits);
                    let labelsObj: DataLabelObject = <DataLabelObject>dataView.metadata.objects[labelString];
                    utils.updateLabelSettingsFromLabelsObject(labelsObj, updateVar.labelSettings);
                }
            }
            return defaultDataPointColor;
        }

        private caseTilteHeight(viewport, updateVar: any, titleText: any, tooltiptext: any, titlecolor: any, titlebgcolor: any, showLegendProp, element, textProperties, color): d3.Selection<SVGElement> {
            if (updateVar.titleHeight !== 0) {
                let totalWidth: number = viewport.width;
                let textProps: TextProperties = { fontFamily: "Segoe UI", fontSize: `${updateVar.titlefontsize}pt`, text: " (?)" };
                let occupiedWidth: number = textMeasurementService.measureSvgTextWidth(textProps);
                const maxWidth: string = `${totalWidth - occupiedWidth}px`;
                if (!titleText) {
                    let titleDiv: d3.Selection<SVGElement>;
                    titleDiv = this.root.select(".hf_Title_Div_Text");
                    titleDiv.append("div").classed("hf_Title_Div_Text_Span", true).style({ "display": "inline-block", "max-width": maxWidth, "visibility": "hidden" }).text(".");
                    titleDiv.append("span").classed("hf_infoImage hf_icon", true).style({ "display": "inline-block", "margin-left": "3px", "position": "absolute", "width": "2%", }).text("(?)").attr("title", tooltiptext);
                } else {
                    let titleTextDiv: d3.Selection<SVGElement>;
                    titleTextDiv = this.root.select(".hf_Title_Div_Text");
                    titleTextDiv.append("div").classed("hf_Title_Div_Text_Span", true).style({ "display": "inline-block", "max-width": maxWidth, }).text(titleText.toString()).attr("title", titleText.toString());
                    titleTextDiv.append("span").classed("hf_infoImage hf_icon", true).text(" (?)").style({ "display": "inline-block", "margin-left": "3px", "position": "absolute", "width": "2%", }).attr("title", tooltiptext);
                }
                if (!tooltiptext) {
                    this.root.select(".hf_infoImage").style({
                        display: "none"
                    });
                } else {
                    this.root.select(".hf_infoImage").attr("title", tooltiptext);
                }
                if (titlecolor) {
                    this.root.select(".hf_Title_Div").style({ "color": titlecolor, "font-size": `${updateVar.titlefontsize}pt` });
                } else {
                    this.root.select(".hf_Title_Div").style({ "color": "#333333", "font-size": `${updateVar.titlefontsize}pt` });
                }
                if (titlebgcolor) {
                    this.root.select(".hf_Title_Div").style({ "background-color": titlebgcolor });
                    this.root.select(".hf_Title_Div_Text").style({ "background-color": titlebgcolor });
                    this.root.select(".hf_infoImage").style({ "background-color": titlebgcolor });
                } else {
                    this.root.select(".hf_Title_Div").style({ "background-color": "none" });
                }
                if (updateVar.funnelTitleOnOffStatus) {
                    this.root.select(".hf_Title_Div").classed("hf_show_inline", true);
                } else {
                    this.root.select(".hf_Title_Div").classed("hf_hide", true);
                }
            }
            if (showLegendProp.show) {
                element.style({ "vertical-align": "top", "width": `${(updateVar.parentWidth - updateVar.width) / (1.8 * updateVar.catLength)}px` });
                element.append("div").style({ color, "font-size": `${updateVar.fontsize}px`, "visibility": "hidden", "width": `${(updateVar.parentWidth - updateVar.width) / (1.8 * updateVar.catLength)}px` }).classed("hf_legend_item", true).text("s");
                let textSize: number = parseInt(updateVar.fontsize.toString(), 10);
                // ellipses for overflow text
                textProperties = { fontFamily: "sans-serif", fontSize: `${updateVar.fontsize}px`, text: this.viewModel.primaryColumn };
                let availableWidth: number = (updateVar.parentWidth - updateVar.width) / (1.8 * updateVar.catLength);
                let primaryColumn: string = textMeasurementService.getTailoredTextOrDefault(textProperties, availableWidth);
                const maxWidth: number = 60;
                const maxHeight: number = 100;
                if (availableWidth < maxWidth || updateVar.height < maxHeight) {
                    element.append("div").style({ color, "font-size": `${updateVar.fontsize}px`, "margin-left": "0", "overflow": "hidden", "padding-right": "10px", "position": "absolute", "width": `${availableWidth}px`, "word-break": "keep-all" })
                        .attr({ title: this.viewModel.primaryColumn }).text(this.trimString(primaryColumn, availableWidth / textSize)).classed("hf_primary_measure", true);
                } else {
                    element.append("div").style({ color, "font-size": `${updateVar.fontsize}px`, "overflow": "hidden", "padding-right": "10px", "position": "absolute", "width": `${availableWidth}px`, "word-break": "keep-all" })
                        .classed("hf_primary_measure", true).attr({ title: this.viewModel.primaryColumn }).text(this.viewModel.primaryColumn);
                }
                if (updateVar.catLength > 0)
                    element.append("svg").attr({ fill: "white", height: updateVar.height, width: (updateVar.parentWidth - updateVar.width) / (1.8 * updateVar.catLength) });
                if (this.viewModel.secondaryColumn !== "") {
                    // ellipses for overflow text
                    textProperties = { fontFamily: "sans-serif", fontSize: `${updateVar.fontsize}px`, text: this.viewModel.secondaryColumn };
                    const catFactor: number = 1.8;
                    let secondaryColumn: string = textMeasurementService.getTailoredTextOrDefault(textProperties, (updateVar.parentWidth - updateVar.width) / (catFactor * updateVar.catLength));
                    element.append("div").style({ color, "font-size": `${updateVar.fontsize}px`, "visibility": "hidden", "width": `${(updateVar.parentWidth - updateVar.width) / (catFactor * updateVar.catLength)}px`, }).classed("hf_legend_item", true).text("s");
                    textSize = parseInt(updateVar.fontsize.toString(), 10);
                    if ((((updateVar.parentWidth - updateVar.width) / (catFactor * updateVar.catLength)) < maxWidth) || (updateVar.height < maxHeight)) {
                        element.append("div").style({
                            color, "font-size": `${updateVar.fontsize}px`, "margin-left": "0", "margin-top": "5px", "overflow": "hidden", "padding-right": "10px", "position": "absolute", "white-space": "normal",
                            "width": `${(updateVar.parentWidth - updateVar.width) / (1.8 * updateVar.catLength)}px`, "word-break": "keep-all"
                        }).attr({ title: this.viewModel.secondaryColumn }).text(this.trimString(secondaryColumn, ((updateVar.parentWidth - updateVar.width) / (catFactor * updateVar.catLength)) / textSize)).classed("hf_yaxis2", true);
                    } else {
                        element.append("div").style({
                            color, "font-size": `${updateVar.fontsize}px`, "margin-left": "0", "margin-top": "5px", "overflow": "hidden", "padding-right": "10px",
                            "position": "absolute", "white-space": "normal", "width": `${(updateVar.parentWidth - updateVar.width) / (catFactor * updateVar.catLength)}px`, "word-break": "keep-all"
                        }).attr({ title: this.viewModel.secondaryColumn }).text(secondaryColumn).classed("hf_yaxis2", true);
                    }
                }
            }
            return element;
        }

        private visualoop(legendvalue, dataPoints, updateVar, sKMBValueY1Axis, textProperties, sKMBValueY2Axis) {
            for (let i: number = 0; i < this.viewModel.categories.length; i++) {
                const maxLabelValue1: number = 1000000000000;
                const maxLabelValue2: number = 1000000000;
                const maxLabelValue3: number = 1000000;
                const maxLabelValue4: number = 1000;
                if (this.viewModel.values[i].values[0] === null || this.viewModel.values[i].values[0] === 0) {
                    updateVar.percentageVal.push(-1);
                } else {
                    if (updateVar.ymax - this.viewModel.values[i].values[0] > 0) {
                        updateVar.percentageVal.push(((this.viewModel.values[i].values[0] * 100) / updateVar.ymax).toFixed(2));
                    } else {
                        updateVar.percentageVal.push(0);
                    }
                }
                legendvalue = this.root.select(`.hf_legend_item${updateVar.legendpos}`);
                if (this.viewModel.categories[i].value !== null && this.viewModel.categories[i].value !== "") {
                    updateVar.title = dataPoints[i].toolTipInfo[0].value;
                    legendvalue.attr({ title: updateVar.title }).text(updateVar.title);
                } else {
                    legendvalue.attr({ title: "(Blank)" }).text("(Blank)");
                }
                if (this.viewModel.values[i].values[0] !== null) {
                    updateVar.title = String(dataPoints[i].toolTipInfo[1].value);
                    updateVar.precisionValue = updateVar.precisionValue;
                    if (updateVar.displayunitValue === 0) {    // auto option selected then
                        updateVar.maxLabel = 0;
                        for (const jIterator of this.viewModel.values) {
                            if (updateVar.maxLabel < jIterator.values[0]) {
                                updateVar.maxLabel = jIterator.values[0];
                            }
                        }
                        updateVar.displayunitValue = updateVar.maxLabel > maxLabelValue1 ? maxLabelValue1 : updateVar.maxLabel > maxLabelValue2 ? maxLabelValue2 : updateVar.maxLabel > maxLabelValue3 ? maxLabelValue3 : updateVar.maxLabel > maxLabelValue4 ? maxLabelValue4 : 1;
                    }
                    sKMBValueY1Axis = this.format(
                        this.viewModel.values[i].values[0],
                        updateVar.displayunitValue, updateVar.precisionValue, this.dataView.categorical.values[0].source.format);
                    let decimalPlaces: number;
                    decimalPlaces = HorizontalFunnel.GETDECIMALPLACECOUNT(this.viewModel.values[i].values[0]);
                    // tooltip values
                    let sKMBValueY1AxisTooltip: string;
                    sKMBValueY1AxisTooltip = this.format(this.viewModel.values[i].values[0], 1, decimalPlaces, this.dataView.categorical.values[0].source.format);
                    updateVar.displayValue = sKMBValueY1Axis;
                    // ellipses for overflow text
                    textProperties = { fontFamily: "sans-serif", fontSize: `${updateVar.fontsize}px`, text: updateVar.displayValue, };
                    updateVar.displayValue = textMeasurementService.getTailoredTextOrDefault(textProperties, updateVar.width);
                    this.root.select(`.hf_legend_value1${updateVar.legendpos}`).attr({ title: sKMBValueY1AxisTooltip }).text(updateVar.displayValue);
                } else {
                    updateVar.displayValue = "(Blank)";
                    updateVar.title = "(Blank)";
                    this.root.select(`.hf_legend_value1${updateVar.legendpos}`).attr({ title: updateVar.title }).text(this.trimString(updateVar.title, updateVar.width / 10));
                }
                if (this.viewModel.values[i].values.length > 1) {
                    let sKMBValueY2AxisTooltip: string;
                    if (this.viewModel.values[i].values[1] !== null) {
                        if (updateVar.displayunitValue2 === 0) { // auto option selected then
                            updateVar.maxLabel = 0;
                            for (const jIterator of this.viewModel.values)
                                if (updateVar.maxLabel < jIterator.values[1]) updateVar.maxLabel = jIterator.values[1];
                            updateVar.displayunitValue2 = updateVar.maxLabel > maxLabelValue1 ? maxLabelValue1 : updateVar.maxLabel > maxLabelValue2 ? maxLabelValue2 : updateVar.maxLabel > maxLabelValue3 ? maxLabelValue3 : updateVar.maxLabel > maxLabelValue4 ? maxLabelValue4 : 1;
                        }
                        sKMBValueY2Axis = this.format(this.viewModel.values[i].values[1], updateVar.displayunitValue2, updateVar.precisionValue, this.dataView.categorical.values[1].source.format);
                        let decimalPlaces: number;
                        decimalPlaces = HorizontalFunnel.GETDECIMALPLACECOUNT(this.viewModel.values[i].values[1]);
                        sKMBValueY2AxisTooltip = this.format(
                            this.viewModel.values[i].values[1], 1,
                            decimalPlaces, this.dataView.categorical.values[1].source.format);
                        if (dataPoints[i].toolTipInfo.length === 3) { // series,y1 and y2
                            updateVar.title = dataPoints[i].toolTipInfo[2].value;
                            updateVar.precisionValue = updateVar.labelSettings.precision;
                            updateVar.displayValue = sKMBValueY2Axis;
                        } else {
                            updateVar.title = dataPoints[i].toolTipInfo[1].value;
                            updateVar.precisionValue = updateVar.precisionValue;
                            updateVar.displayValue = sKMBValueY2Axis;
                        }
                        textProperties = { fontFamily: "sans-serif", fontSize: `${updateVar.fontsize}px`, text: updateVar.displayValue };
                        updateVar.displayValue = textMeasurementService.getTailoredTextOrDefault(textProperties, updateVar.width);
                        this.root.select(`.hf_legend_value2${updateVar.legendpos}`).attr({ title: sKMBValueY2AxisTooltip }).text(updateVar.displayValue);
                    } else {
                        updateVar.displayValue = "(Blank)";
                        updateVar.title = "(Blank)";
                        this.root.select(`.hf_legend_value2${updateVar.legendpos}`).attr({ title: updateVar.title }).text(this.trimString(updateVar.title, updateVar.width / 10));
                    }
                }
                updateVar.legendpos += 2;
            }
        }
        private visualoop2(updateVar, oddsvg, areafillheight) {
            for (let i: number = 0; i < (2 * updateVar.catLength - 1); i++) {
                if (i % 2 === 0) {
                    updateVar.classname = `hf_odd${i}`;
                    oddsvg = this.root.select(`.${updateVar.classname}`);
                    if (updateVar.percentageVal[updateVar.index] !== 0 && updateVar.percentageVal[updateVar.index] !== -1) {
                        updateVar.percentageVal[updateVar.index] = parseFloat(updateVar.percentageVal[updateVar.index]);
                        updateVar.y = 0;
                        updateVar.y = ((updateVar.height - (updateVar.percentageVal[updateVar.index] * updateVar.height / 100)) / 2);
                        areafillheight.push(updateVar.percentageVal[updateVar.index] * updateVar.height / 100);
                        let disp: number;
                        disp = 10;
                        oddsvg.append("rect")
                            .attr({
                                height: areafillheight[updateVar.index],
                                width: updateVar.width,
                                x: disp,
                                y: updateVar.y
                            }).classed("hf_datapoint hf_dataColor", true);
                    } else {
                        let disp: number;
                        disp = 10;
                        if (updateVar.percentageVal[updateVar.index] === 0) {
                            oddsvg.append("rect")
                                .attr({
                                    height: updateVar.height,
                                    width: updateVar.height,
                                    x: disp,
                                    y: 0,
                                }).classed("hf_datapoint hf_dataColor", true);
                        } else if (updateVar.percentageVal[updateVar.index] === -1) {
                            // showing dotted line if there is no data
                            updateVar.y = ((updateVar.height - (updateVar.percentageVal[updateVar.index] * updateVar.height / 100)) / 2);
                            oddsvg.append("line")
                                .attr({
                                    "stroke-width": 1,
                                    "x1": disp,
                                    "x2": updateVar.width,
                                    "y1": updateVar.y,
                                    "y2": updateVar.y
                                }).classed("hf_datapoint hf_dataColor", true)
                                .style("stroke-dasharray", "1,2");
                        }
                        areafillheight.push(0);
                    }
                    updateVar.index++;
                }
            }
        }
        private sideBar(updateVar, evensvg, areafillheight) {
            for (let i: number = 0; i < updateVar.percentageVal.length; i++) {
                let polygonColor: string;
                if (this.defaultDataPointColor) {
                    polygonColor = this.defaultDataPointColor;
                } else {
                    polygonColor = this.colorLuminance(this.viewModel.categories[i].color.value);
                }
                updateVar.classname = `.hf_even${updateVar.val}`;
                evensvg = this.root.select(updateVar.classname);
                if (updateVar.percentageVal[i] === 0 && updateVar.percentageVal[i + 1] === 0) {
                    evensvg.append("rect")
                        .attr({
                            fill: polygonColor,
                            height: updateVar.height,
                            width: updateVar.width / 4,
                            x: 10,
                            y: 0
                        });
                } else {
                    updateVar.prevyheight = (updateVar.height - areafillheight[i]) / 2;
                    updateVar.nextyheight = (updateVar.height - areafillheight[i + 1]) / 2;
                    let disp: number;
                    disp = 10;
                    if (updateVar.percentageVal[i] && updateVar.percentageVal[i + 1]) {
                        updateVar.dimension = `${disp},${updateVar.prevyheight} ${disp},${areafillheight[i] + updateVar.prevyheight} ` +
                            `${updateVar.width / 4},${areafillheight[i + 1] + updateVar.nextyheight} ${updateVar.width / 4},${updateVar.nextyheight}`;
                    } else if (updateVar.percentageVal[i] && !(updateVar.percentageVal[i + 1])) {
                        updateVar.dimension = `${disp},${updateVar.prevyheight} ${disp},${areafillheight[i] + updateVar.prevyheight} ` +
                            `${updateVar.width / 4},${updateVar.height} ${updateVar.width / 4},0`;
                    } else if (!(updateVar.percentageVal[i]) && updateVar.percentageVal[i + 1]) {
                        updateVar.dimension = `${disp},0 ${disp},${updateVar.height} ${updateVar.width / 4},${areafillheight[i + 1] + updateVar.nextyheight} ` +
                            `${updateVar.width / 4},${updateVar.nextyheight}`;
                    }
                    evensvg
                        .append("polygon")
                        .attr("points", updateVar.dimension)
                        .attr({ fill: polygonColor });
                }
                updateVar.val += 2;
            }
        }
        private endVisual(updateVar, evensvg, areafillheight,dataPoints,selection,options){
            this.sideBar(updateVar, evensvg, areafillheight);
            this.root.selectAll(".fillcolor").style("fill", (d: {}, i: number) => this.colors[i + 1].value);
            this.root.selectAll(".hf_dataColor").style("fill", (d: {}, i: number) => this.viewModel.categories[i].color.value);
            // This is for the dotted line
            this.root.selectAll(".hf_dataColor").style("stroke", (d: {}, i: number) => this.viewModel.categories[i].color.value);
            selection = this.root.selectAll(".hf_datapoint").data(dataPoints, (d: {}, idx: number) => (dataPoints[idx] === 0) ? String(idx) : String(idx + 1));
            let viewModel: any = this.viewModel;
            this.tooltipServiceWrapper.addTooltip(d3.selectAll(".hf_datapoint"), (tooltipEvent: TooltipEventArgs<number>) => { return tooltipEvent.context["cust-tooltip"]; }, (tooltipEvent: TooltipEventArgs<number>) => null, true);
            selection.on("click", (data: IFunnelViewModel) => {
                let ev: any = d3.event; this.selectionManager.select(data.identity, ev.ctrlKey).then((selectionIds: ISelectionId[]) => { this.syncSelectionState(selection, selectionIds); });
                ev.stopPropagation();
            });
            this.root.on("click", () => { this.selectionManager.clear(); this.syncSelectionState(selection, <ISelectionId[]>this.selectionManager.getSelectionIds()); });
            this.syncSelectionState(selection, <ISelectionId[]>this.selectionManager.getSelectionIds());
            d3.selectAll(".hf_parentdiv").on("contextmenu", () => {
                const mouseEvent: MouseEvent = <MouseEvent>d3.event;
                const eventTarget: EventTarget = mouseEvent.target;
                const dataPoint: any = d3.select(eventTarget).datum();
                if (dataPoint !== undefined) {
                    this.selectionManager.showContextMenu(dataPoint.identity, { x: mouseEvent.clientX, y: mouseEvent.clientY });
                    mouseEvent.preventDefault();
                }
            });
        }
        // tslint:disable-next-line:cyclomatic-complexity
        public update(options: VisualUpdateOptions): void {
            this.events.renderingStarted(options);
            let host: IVisualHost;
            host = this.host;
            if (!options.dataViews || (options.dataViews.length < 1) || !options.dataViews[0] || !options.dataViews[0].categorical)
                return;
            let dataView: DataView = this.dataView = options.dataViews[0];
            let ytot: number = 0, yarr: any = [], sKMBValueY1Axis: any, sKMBValueY2Axis: any, color: any, titlecolor: any, titlebgcolor: any, titleText: any, tooltiptext: any, defaultText: d3.Selection<SVGElement>, parentDiv: d3.Selection<SVGElement>;
            let showDefaultText: number, viewport: IViewport, dataPoints: any, parentHeight: number, element: d3.Selection<SVGElement>, legendvalue: d3.Selection<SVGElement>, oddsvg: d3.Selection<SVGElement>, evensvg: d3.Selection<SVGElement>, selection: any;
            let areafillheight: any = [], visualHeight: number, textHeight: number, titlemargin: number;
            this.cardFormatSetting = this.getDefaultFormatSettings();
            defaultText = this.root.select(".hf_defaultText");
            let dataViewMetadata: DataViewMetadata = dataView.metadata, defaultDataPointColor: any;
            let updateVar = {
                labelSettings: null, precisionValue: 0, titleHeight: 0, titlefontsize: 0, funnelTitleOnOffStatus: false, parentWidth: 0, width: 0, catLength: 0, fontsize: 0, height: 0, percentageVal: [], ymax: 0,
                legendpos: 0, title: "", displayunitValue: 0, displayunitValue2: 0, maxLabel: 0, displayValue: "", classname: "", index: 0, y: 0, val: 1, prevyheight: 0, nextyheight: 0, dimension: ""
            };
            defaultDataPointColor=this.dataViewMeta(dataViewMetadata, updateVar, dataView,defaultDataPointColor);
            let showLegendProp: IShowLegendSettings = this.getLegendSettings(this.dataView), funnelTitleSettings: IFunnelTitle = this.getFunnelTitleSettings(this.dataView), showConnectorsProp: IShowConnectorsSettings = this.getConnectorsSettings(this.dataView), dataLabelSettings: ILabelSettings = this.getDataLabelSettings(this.dataView);
            let sortSettings: ISortSettings = this.getSortSettings(this.dataView);
            this.defaultDataPointColor = defaultDataPointColor;
            viewport = options.viewport;
            this.root.selectAll("div").remove();
            dataPoints = HorizontalFunnel.CONVERTER(dataView, horizontalFunnelProps.cPalette, sortSettings.sortBy, sortSettings.orderBy, this.host);
            this.viewModel = dataPoints[0];
            updateVar.catLength = this.viewModel.categories.length;
            updateVar.parentWidth = viewport.width;
            parentHeight = viewport.height;
            updateVar.width = updateVar.parentWidth / (1.4 * updateVar.catLength);
            if (!showConnectorsProp.show) updateVar.width = updateVar.width + (((updateVar.width / 4) * (updateVar.catLength - 1)) / updateVar.catLength);
            visualHeight = parentHeight >= 65 ? parentHeight - 65 + 40 : 65 - parentHeight;
            for (let iLoop: number = 0; iLoop < this.viewModel.categories.length; iLoop++) {
                yarr.push(this.viewModel.values[iLoop].values[0]);
                ytot += this.viewModel.values[iLoop].values[0];
                updateVar.ymax = Math.max.apply(Math, yarr);
            }
            updateVar.funnelTitleOnOffStatus = funnelTitleSettings.show;
            titleText = funnelTitleSettings.titleText;
            tooltiptext = funnelTitleSettings.tooltipText;
            updateVar.titlefontsize = funnelTitleSettings.fontSize;
            if (!updateVar.titlefontsize) updateVar.titlefontsize = 12;
            if (updateVar.funnelTitleOnOffStatus && (titleText || tooltiptext)) {
                let titleTextProperties: TextProperties = { fontFamily: "Segoe UI,wf_segoe-ui_normal,helvetica,arial,sans-serif", fontSize: `${updateVar.titlefontsize}pt`, text: titleText.toString() };
                updateVar.titleHeight = textMeasurementService.measureSvgTextHeight(titleTextProperties);
            } else updateVar.titleHeight = 0;
            updateVar.fontsize = dataLabelSettings.fontSize;
            let textProperties: TextProperties = { fontFamily: "Segoe UI,wf_segoe-ui_normal,helvetica,arial,sans-serif", fontSize: `${updateVar.fontsize}pt`, text: "MAQ Software" };
            textHeight = textMeasurementService.measureSvgTextHeight(textProperties);
            let totalTextHeight: number = updateVar.titleHeight + textHeight * 2;
            updateVar.height = totalTextHeight > visualHeight ? totalTextHeight - visualHeight : visualHeight - totalTextHeight;
            if (!this.viewModel.secondaryColumn) updateVar.height += (1.4 * updateVar.fontsize);
            titlemargin = updateVar.titleHeight;
            if (updateVar.titleHeight !== 0) {
                this.root.append("div").style({ height: `${updateVar.titleHeight}px`, width: "100%" }).classed("hf_Title_Div", true);
                this.root.select(".hf_Title_Div").append("div").style({ width: "100%" }).classed("hf_Title_Div_Text", true);
                this.root.select(".hf_Title_Div_Text").classed("hf_title", true).style({ display: "inline-block" });
                this.root.select(".hf_Title_Div_Text").classed("hf_title", true).style({ display: "inline-block" });
            }
            this.root.append("div").style({ "height": `${updateVar.height + 60}px`, "margin-bottom": "5px", "width": `${updateVar.parentWidth}px`, }).classed("hf_parentdiv", true);
            element = this.root.select(".hf_parentdiv").append("div").classed("hf_svg hf_parentElement", true);
            parentDiv = this.root.select(".hf_parentdiv"); showDefaultText = 1;
            if (dataView && dataView.categorical && dataView.categorical.values) {
                for (const iterator of dataView.categorical.values)
                    if (iterator.source.roles && iterator.source.roles.hasOwnProperty("primaryMeasure")) showDefaultText = 0;
            }
            if (!dataView.categorical.categories || 1 === showDefaultText) {
                let message: string;
                message = 'Please select both "Series" and "Primary Measure" values';
                parentDiv.append("div").text(message).attr("title", message).classed("hf_defaultText", true).style({ top: `${updateVar.height / 2.5}px` });
            }
            if (updateVar.labelSettings !== undefined && updateVar.labelSettings !== null) {
                updateVar.displayunitValue = (updateVar.labelSettings.displayUnits ? updateVar.labelSettings.displayUnits : 0);
                updateVar.displayunitValue2 = (updateVar.labelSettings.displayUnits ? updateVar.labelSettings.displayUnits : 0);
                color = updateVar.labelSettings.labelColor; updateVar.fontsize = updateVar.labelSettings.fontSize;
            }
            titlecolor = funnelTitleSettings.color;
            titlebgcolor = funnelTitleSettings.bkColor;
            element = this.caseTilteHeight(viewport, updateVar, titleText, tooltiptext, titlecolor, titlebgcolor, showLegendProp, element, textProperties, color);
            this.printerloop(showLegendProp, showConnectorsProp, element, color, updateVar);
            this.visualoop(legendvalue, dataPoints, updateVar, sKMBValueY1Axis, textProperties, sKMBValueY2Axis);
            this.visualoop2(updateVar, oddsvg, areafillheight);
            let svgElement: d3.Selection<SVGAElement> = d3.selectAll(".hf_datapoint.hf_dataColor");
            for (let i: number = 0; i < (updateVar.catLength); i++) {
                svgElement[0][i]["cust-tooltip"] = this.viewModel.values[i].toolTipInfo;
            }
            this.endVisual(updateVar, evensvg, areafillheight,dataPoints,selection,options);
            this.events.renderingFinished(options);
        }


        public getDefaultLegendSettings(): IShowLegendSettings {

            return {
                show: false
            };
        }

        public getLegendSettings(dataView: DataView): IShowLegendSettings {
            let objects: DataViewObjects = null;
            let legendSetting: IShowLegendSettings;
            legendSetting = this.getDefaultLegendSettings();

            if (!dataView.metadata || !dataView.metadata.objects) {
                return legendSetting;
            }
            objects = dataView.metadata.objects;
            // tslint:disable-next-line:no-any
            const legendProperties: any = horizontalFunnelProps;
            legendSetting.show =
                DataViewObjects.getValue(objects, legendProperties.ShowLegend.show, legendSetting.show);

            return legendSetting;

        }

        public getDefaultConnectorsSettings(): IShowConnectorsSettings {
            return {
                show: true
            };
        }

        public getConnectorsSettings(dataView: DataView): IShowConnectorsSettings {
            let objects: DataViewObjects = null;
            let connectorsSetting: IShowConnectorsSettings;
            connectorsSetting = this.getDefaultConnectorsSettings();

            if (!dataView.metadata || !dataView.metadata.objects) {
                return connectorsSetting;
            }
            objects = dataView.metadata.objects;
            // tslint:disable-next-line:no-any
            const showConnectorsProps: any = horizontalFunnelProps;
            connectorsSetting.show =
                DataViewObjects.getValue(objects, showConnectorsProps.ShowConnectors.show, connectorsSetting.show);

            return connectorsSetting;

        }

        public getDefaultDataLabelSettings(): ILabelSettings {
            return {
                color: "#333333",
                decimalPlaces: 0,
                displayUnits: 0,
                fontSize: 12
            };
        }

        public getDataLabelSettings(dataView: DataView): ILabelSettings {

            let objects: DataViewObjects = null;
            let dataLabelSetting: ILabelSettings;
            dataLabelSetting = this.getDefaultDataLabelSettings();
            if (!dataView.metadata || !dataView.metadata.objects) {
                return dataLabelSetting;
            }
            objects = dataView.metadata.objects;
            // tslint:disable-next-line:no-any
            const labelProperties: any = horizontalFunnelProps.LabelSettings;
            dataLabelSetting.color =
                DataViewObjects.getFillColor(objects, labelProperties.color, dataLabelSetting.color);
            dataLabelSetting.displayUnits = DataViewObjects.getValue(
                objects, labelProperties.labelDisplayUnits, dataLabelSetting.displayUnits);
            dataLabelSetting.decimalPlaces = DataViewObjects.getValue(
                objects, labelProperties.labelPrecision, dataLabelSetting.decimalPlaces);
            dataLabelSetting.decimalPlaces = dataLabelSetting.decimalPlaces < 0 ?
                0 : dataLabelSetting.decimalPlaces > 4 ? 4 : dataLabelSetting.decimalPlaces;
            dataLabelSetting.fontSize =
                DataViewObjects.getValue(objects, labelProperties.fontSize, dataLabelSetting.fontSize);

            return dataLabelSetting;
        }

        public getDefaultFunnelTitleSettings(dataView: DataView): IFunnelTitle {
            let titleText: string = "";
            if (dataView
                && dataView.categorical
                && dataView.categorical.categories
                && dataView.categorical.categories[0]
                && dataView.categorical.categories[0].source
                && dataView.categorical.categories[0].source.displayName
                && dataView.categorical.values
                && dataView.categorical.values[0]
                && dataView.categorical.values[0].source
                && dataView.categorical.values[0].source.displayName) {
                const measureName: string = dataView.categorical.values[0].source.displayName;
                const catName: string = dataView.categorical.categories[0].source.displayName;
                titleText = `${measureName} by ${catName} `;
            }

            return {
                bkColor: "#fff",
                color: "#333333",
                fontSize: 12,
                show: true,
                titleText,
                tooltipText: "Your tooltip text goes here",
            };
        }

        public getFunnelTitleSettings(dataView: DataView): IFunnelTitle {
            let objects: DataViewObjects = null;
            let fTitleSettings: IFunnelTitle;
            fTitleSettings = this.getDefaultFunnelTitleSettings(dataView);
            if (!dataView.metadata || !dataView.metadata.objects) {
                return fTitleSettings;
            }

            objects = dataView.metadata.objects;
            // tslint:disable-next-line:no-any
            const titleProps: any = horizontalFunnelProps.funnelTitle;
            fTitleSettings.show = DataViewObjects.getValue(objects, titleProps.show, fTitleSettings.show);
            fTitleSettings.titleText =
                DataViewObjects.getValue(objects, titleProps.titleText, fTitleSettings.titleText);
            fTitleSettings.tooltipText =
                DataViewObjects.getValue(objects, titleProps.tooltipText, fTitleSettings.tooltipText);
            fTitleSettings.color =
                DataViewObjects.getFillColor(objects, titleProps.titleFill, fTitleSettings.color);
            fTitleSettings.bkColor =
                DataViewObjects.getFillColor(objects, titleProps.titleBackgroundColor, fTitleSettings.bkColor);
            fTitleSettings.fontSize =
                DataViewObjects.getValue(objects, titleProps.titleFontSize, fTitleSettings.fontSize);

            return fTitleSettings;
        }

        public getDefaultSortSettings(): ISortSettings {
            return {
                orderBy: "ascending",
                sortBy: "Auto"
            };
        }

        public getSortSettings(dataView: DataView): ISortSettings {
            let objects: DataViewObjects = null;
            let sortSettings: ISortSettings;
            sortSettings = this.getDefaultSortSettings();
            if (!dataView.metadata || !dataView.metadata.objects) {
                return sortSettings;
            }
            objects = dataView.metadata.objects;
            // tslint:disable-next-line:no-any
            const sortProps: any = horizontalFunnelProps.sort;
            sortSettings.sortBy = DataViewObjects.getValue(objects, sortProps.sortBy, sortSettings.sortBy);
            // Check if Secondary measure exists before selecting it
            // If exists sort by secondary measure, otherwise sort by 'Auto'
            if (sortSettings.sortBy === "SecondaryMeasure") {
                if (dataView
                    && dataView.categorical
                    && dataView.categorical.values) {
                    let secondaryColumn: boolean = false;
                    // tslint:disable-next-line: prefer-for-of
                    for (let iLoop: number = 0; iLoop < dataView.categorical.values.length; iLoop++) {
                        if (dataView.categorical.values[iLoop].source
                            && dataView.categorical.values[iLoop].source.roles
                            && dataView.categorical.values[iLoop].source.roles.hasOwnProperty("secondaryMeasure")) {
                            secondaryColumn = true;
                        }
                    }
                    sortSettings.sortBy = secondaryColumn ? "SecondaryMeasure" : "Auto";
                }
            }
            sortSettings.orderBy = DataViewObjects.getValue(objects, sortProps.orderBy, sortSettings.orderBy);

            return sortSettings;
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions):
            VisualObjectInstanceEnumeration {
            let enumeration: VisualObjectInstance[];
            enumeration = [];
            let showLegendSettings: IShowLegendSettings;
            showLegendSettings = this.getLegendSettings(this.dataView);
            let showConnectorsSettings: IShowConnectorsSettings;
            showConnectorsSettings = this.getConnectorsSettings(this.dataView);
            let dataLabelSettings: ILabelSettings;
            dataLabelSettings = this.getDataLabelSettings(this.dataView);
            let funnelTitleSettings: IFunnelTitle;
            funnelTitleSettings = this.getFunnelTitleSettings(this.dataView);
            let sortSettings: ISortSettings;
            sortSettings = this.getSortSettings(this.dataView);

            switch (options.objectName) {
                case "FunnelTitle":
                    enumeration.push({
                        displayName: "Funnel title",
                        objectName: "FunnelTitle",
                        properties: {
                            backgroundColor: funnelTitleSettings.bkColor,
                            fill1: funnelTitleSettings.color,
                            fontSize: funnelTitleSettings.fontSize,
                            show: funnelTitleSettings.show,
                            titleText: funnelTitleSettings.titleText,
                            tooltipText: funnelTitleSettings.tooltipText,
                        },
                        selector: null,
                    });
                    break;

                case "Sort":
                    enumeration.push({
                        displayName: "Sort",
                        objectName: "Sort",
                        properties: {
                            OrderBy: sortSettings.orderBy,
                            SortBy: sortSettings.sortBy
                        },
                        selector: null,
                    });
                    break;

                case "labels":
                    enumeration.push({
                        objectName: options.objectName,
                        properties: {
                            color: dataLabelSettings.color,
                            fontSize: dataLabelSettings.fontSize,
                            labelDisplayUnits: dataLabelSettings.displayUnits,
                            labelPrecision: dataLabelSettings.decimalPlaces
                        },
                        selector: null
                    });
                    break;

                case "dataPoint":
                    this.enumerateDataPoints(enumeration);
                    break;

                case "ShowLegend":
                    enumeration.push({
                        displayName: "Show Legend",
                        objectName: "ShowLegend",
                        properties: {
                            show: showLegendSettings.show
                        },
                        selector: null
                    });
                    break;

                case "ShowConnectors":
                    enumeration.push({
                        displayName: "Show Connectors",
                        objectName: "ShowConnectors",
                        properties: {
                            show: showConnectorsSettings.show
                        },
                        selector: null
                    });
                    break;

                default:
                    break;
            }

            return enumeration;
        }

        public getDefaultFormatSettings(): ICardFormatSetting {
            return {
                labelSettings: this.getDefaultLabelSettings(true, "#333333", undefined, undefined),
                showTitle: true,
                textSize: 10,
                wordWrap: false
            };
        }

        // tslint:disable-next-line:no-any
        public getDefaultLabelSettings(show: any, labelColor: any, labelPrecision: any, format: any): {
            // tslint:disable-next-line:no-any
            show: any;
            position: number;
            displayUnits: number;
            precision: number;
            labelColor: {};
            formatterOptions: {};
            fontSize: number;
        } {
            let defaultLabelColor: string;
            defaultLabelColor = "#333333";
            let precision: number = 0;
            if (precision > 4) {
                precision = 4;
            }
            if (show === void 0) { show = false; }
            if (format) {
                let hasDots: boolean;
                hasDots = true; // powerbi.NumberFormat.getCustomFormatMetadata(format).hasDots;
            }

            return {
                displayUnits: 0,
                fontSize: 12,
                formatterOptions: null,
                labelColor: labelColor || defaultLabelColor,
                position: 0 /* Above */,
                precision,
                show
            };
        }

        // This function is to trim numbers if it exceeds number of digits.
        // tslint:disable-next-line:no-any
        public trimString(sValue: any, iNumberOfDigits: number): string {
            if (null === sValue) {
                return "null";
            }
            if (sValue.toString().length < iNumberOfDigits) {
                return sValue;
            } else {
                return (`${sValue.toString().substring(0, iNumberOfDigits)}...`);
            }
        }

        private colorLuminance(hex: string): string {
            let lum: number = 0.50;
            // validate hex string
            hex = String(hex).replace(/[^0-9a-f]/gi, "");
            if (hex.length < 6) {
                hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
            }
            lum = lum || 0;

            // convert to decimal and change luminosity
            let rgb: string = "#";
            // tslint:disable-next-line:no-any
            let c: any;
            let i: number;
            for (i = 0; i < 3; i++) {
                c = parseInt(hex.substr(i * 2, 2), 16);
                c = Math.round(Math.min(Math.max(0, c + (c * lum)), 255)).toString(16);
                rgb += (`00${c}`).substr(c.length);
            }

            return rgb;
        }

        private enumerateDataPoints(enumeration: VisualObjectInstance[]): void {
            // tslint:disable-next-line:no-any
            const data: any = this.viewModel.categories;
            if (!data) {
                return;
            }
            const dataPointsLength: number = data.length;

            const primaryValues: number = this.viewModel.values;
            for (let i: number = 0; i < dataPointsLength; i++) {
                if (primaryValues[i].values[0]) {
                    if (!data[i].color.value) {
                        data[i].color.value = "(Blank)";
                    }
                    enumeration.push({
                        displayName: data[i].value,
                        objectName: "dataPoint",
                        properties: {
                            fill: { solid: { color: data[i].color.value } }
                        },
                        selector: data[i].identity.getSelector(),
                    });
                }
            }
            enumeration.push({
                objectName: "dataPoint",
                properties: {
                    defaultColor: { solid: { color: this.defaultDataPointColor } }
                },
                selector: null,
            });
        }

        // This function is to perform KMB formatting on values.
        private format(d: number, displayunitValue: number, precisionValue: number, format: string): string {

            let displayUnits: number;
            displayUnits = displayunitValue;
            let primaryFormatterVal: number = 0;
            if (displayUnits === 0) {
                let alternateFormatter: number;
                alternateFormatter = d.toString().length;
                if (alternateFormatter > 9) {
                    primaryFormatterVal = 1e9;
                } else if (alternateFormatter <= 9 && alternateFormatter > 6) {
                    primaryFormatterVal = 1e6;
                } else if (alternateFormatter <= 6 && alternateFormatter >= 4) {
                    primaryFormatterVal = 1e3;
                } else {
                    primaryFormatterVal = 10;
                }
            }
            let formatter: IValueFormatter;
            if (format) {
                if (format.indexOf("%") >= 0) {
                    formatter = valueFormatter.create({
                        format,
                        precision: precisionValue
                    });
                } else {
                    formatter = valueFormatter.create({
                        format,
                        precision: precisionValue,
                        value: displayUnits === 0 ? primaryFormatterVal : displayUnits
                    });
                }
            } else {
                formatter = valueFormatter.create({
                    precision: precisionValue,
                    value: displayUnits === 0 ? primaryFormatterVal : displayUnits
                });
            }

            let formattedValue: string;
            formattedValue = formatter.format(d);

            return formattedValue;
        }

        // method to set opacity based on the selections in visual
        // tslint:disable-next-line:no-any
        private syncSelectionState(selection: any, selectionIds?: ISelectionId[]): void {
            const self: this = this;

            if (!selection || !selectionIds) {
                return;
            }

            if (!selectionIds.length) {
                selection.transition()
                    .duration(this.durationAnimations)
                    .style("fill-opacity", HorizontalFunnel.maxOpacity);

                return;
            }

            selection.each(function (dataPoint: IFunnelViewModel): void {
                const isSelected: boolean = self.isSelectionIdInArray(selectionIds, dataPoint.identity);

                d3.select(this).transition()
                    .duration(self.durationAnimations)
                    .style("fill-opacity", isSelected ? HorizontalFunnel.maxOpacity : HorizontalFunnel.minOpacity);
            });
        }

        // method to return boolean based on presence of value in array
        private isSelectionIdInArray(selectionIds: ISelectionId[], selectionId: ISelectionId): boolean {
            if (!selectionIds || !selectionId) {
                return false;
            }

            return selectionIds.some((currentSelectionId: ISelectionId) => {
                return currentSelectionId.includes(selectionId);
            });
        }
    }
}
