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

module powerbi.extensibility.visual {
    /**
         * Interface for ControlChart viewmodel.
         *
         * @interface
         * @property {ChartDataPoint[]} dataPoints  - Set of data points the visual will render.
         * @property {any} minX                     - minimum value of X axis - can be date or number
         * @property {any} maxX                     - maximum value of X axis - can be date or number
         * @property {any} minY                     - minimum value of Y axis - can be number
         * @property {any} maxY                     - maximum value of Y axis - can be number
         * @property {interface} data               - LineData - data point line
         * @property {interface} stageDividerLine   - StatisticsData - contains info on line and labels
         * @property {interface} limitLine          - StatisticsData - contains info on line and labels
         * @property {interface} meanLine           - StatisticsData - contains info on line and labels
         * @property {interface} xAxis              - AxisData - contains info on labels
         * @property {interface} yAxis              - minimum value of X axis - can be date or number
         * @property {boolean} showGridLines        - show gridlines
         * @property {boolean} isDateRange          - is the X axis a date or numeric range?
         * @property {boolean} runRule1             - run rule 1
         * @property {boolean} runRule2             - run rule 2
         * @property {boolean} runRule3             - run rule 3
         */
    interface ControlChartViewModel {
        dataPoints: ChartDataPoint[];
        minX: any;
        maxX: any;
        minY: number;
        maxY: number;
        data: LineData;
        stageDividerLine: StatisticsData;
        limitLine: StatisticsData;
        meanLine: StatisticsData;
        standardDeviations: number;
        xAxis: AxisData;
        yAxis: AxisData;
        showGridLines: boolean;
        isDateRange: boolean;
        runRule1: boolean;
        runRule2: boolean;
        runRule3: boolean;
        ruleColor: string;
        movingRange: number;
        mRError: boolean;
    };


    interface LineData {
        DataColor: string;
        LineColor: string;
        LineStyle: string;
    }

    interface AxisData {
        AxisTitle: string;
        TitleSize: number;
        TitleFont?: string;
        TitleColor: string;
        AxisLabelSize: number;
        AxisLabelFont?: string;
        AxisLabelColor: string;
        AxisFormat: any;
    }

    interface StatisticsData {
        text?: string;
        textSize: number;
        textFont?: any;
        textColor: string;
        lineColor: string;
        lineStyle: string;
        show: boolean;
    }

    /**
     * Interface for ControlChart data points.
     *
     * @interface
     * @property {any} xvalue               - Data value for point. - date or number
     * @property {number} yValue            - y axis value.
     */
    interface ChartDataPoint {
        xValue: any;
        yValue: number;
        MRSum: number;
    };

    /**
         * Interface for ControlChart Stage.
         *
         * @interface
         * @property {number} UCL           - Upper control limit.
         * @property {number} LCL           - Lower control limit.
         * @property {number} Mean          - Mean
         * @property {number} startX        - first point in stage
         * @property {number} endX          - last point in stage
         * @property {number} sum           - sum of data in stage
         * @property {number} count         - number of points in stage
         * @property {string} stage         - Stage name
         * @property {any} stageDividerX    - stage divider x axis value - date or number
         * @property {number} firstId       - first index value in a stage - use this in calculating stats
         * @property {number} lastId        - first index value in a stage - use this in calculating stats
         * @property {boolean} mrError      - if moving range is >= number of data points in each stage
         */
    interface Stage {
        lCL: number;
        uCL: number;
        mean: number;
        startX: any;
        endX: any;
        sum: number;
        count: number;
        stage: string;
        stageDividerX: any;
        firstId: number;
        lastId: number;
        mRError: boolean;
    };

    interface SubGroup {
        mean: number;
        sum: number;
        count: number;
        stage: string;
    };

    /**
         * Function that converts queried data into a view model that will be used by the visual.
         *
         * @function
         * @param {VisualUpdateOptions} options - Contains references to the size of the container
         *                                        and the dataView which contains all the data
         *                                        the visual had queried.
         * @param {IVisualHost} host            - Contains references to the host which contains services
         */
    function visualTransform(options: VisualUpdateOptions, host: IVisualHost): ControlChartViewModel {
        let dataViews = options.dataViews;
        let viewModel: ControlChartViewModel = {
            dataPoints: [],
            minX: null,
            maxX: null,
            minY: 0,
            maxY: 0,
            data: null,
            xAxis: null,
            yAxis: null,
            stageDividerLine: null,
            meanLine: null,
            limitLine: null,
            standardDeviations: 3,
            showGridLines: true,
            isDateRange: true,
            runRule1: false,
            runRule2: false,
            runRule3: false,
            ruleColor: null,
            movingRange: 2,
            mRError: false
        };

        if (!dataViews
            || !dataViews[0]
            || !dataViews[0].categorical
            || !dataViews[0].categorical.categories
            || !dataViews[0].categorical.categories[0].source
            || !dataViews[0].categorical.values)
            return viewModel;

        let categorical = dataViews[0].categorical;
        let category = categorical.categories[0];
        let dataValue = categorical.values[0];
        let ChartDataPoints: ChartDataPoint[] = [];
        var xValues: PrimitiveValue[] = category.values;
        var yValues: PrimitiveValue[] = dataValue.values;

        var isXAxisUsable: boolean = category.values[0] && ((Object.prototype.toString.call(category.values[0]) === '[object Number]') || (Object.prototype.toString.call(category.values[0]) === '[object Date]'));
        var isYAxisNumericData: boolean = (dataValue.values[0] && Object.prototype.toString.call(dataValue.values[0]) === '[object Number]');

        if (isXAxisUsable && isYAxisNumericData) {
            for (let i = 0; i < xValues.length; i++) {
                ChartDataPoints.push({
                    xValue: xValues[i],
                    yValue: <number>yValues[i],
                    MRSum: 0
                });
            }
            var isDateRange: boolean = (Object.prototype.toString.call(ChartDataPoints[0].xValue) === '[object Date]');

            var xAxisFormat: any;
            if (isDateRange)
                xAxisFormat = getValue<string>(dataViews[0].metadata.objects, 'xAxis', 'xAxisFormat', '%d-%b-%y');
            else
                xAxisFormat = getValue<string>(dataViews[0].metadata.objects, 'xAxis', 'xAxisFormat', '.3s')

            let chartData: LineData = {
                DataColor: getFill(dataViews[0], 'chart', 'dataColor', '#FF0000'),
                LineColor: getFill(dataViews[0], 'chart', 'lineColor', '#0000FF'),
                LineStyle: getValue<string>(dataViews[0].metadata.objects, 'chart', 'lineStyle', '')
            };
            let meanLine: StatisticsData = {
                textColor: getFill(dataViews[0], 'statistics', 'meanLabelColor', '#008000'),
                textSize: getValue<number>(dataViews[0].metadata.objects, 'statistics', 'meanLabelSize', 10),
                lineColor: getFill(dataViews[0], 'statistics', 'meanLineColor', '#32CD32'),
                lineStyle: getValue<string>(dataViews[0].metadata.objects, 'statistics', 'meanLineStyle', '6,4'),
                show: getValue<boolean>(dataViews[0].metadata.objects, 'statistics', 'showMean', true)
            };
            let stageDividerLine: StatisticsData = {
                textColor: getFill(dataViews[0], 'statistics', 'stageLabelColor', '#FFD700'),
                textSize: getValue<number>(dataViews[0].metadata.objects, 'statistics', 'stageDividerLabelSize', 12),
                lineColor: getFill(dataViews[0], 'statistics', 'stageDividerColor', '#FFD700'),
                lineStyle: getValue<string>(dataViews[0].metadata.objects, 'statistics', 'stageDividerLineStyle', '10,4'),
                show: getValue<boolean>(dataViews[0].metadata.objects, 'statistics', 'showDividers', true)
            };
            let limitLine: StatisticsData = {
                textColor: getFill(dataViews[0], 'statistics', 'limitLabelColor', '#FFA500'),
                textSize: getValue<number>(dataViews[0].metadata.objects, 'statistics', 'limitLabelSize', 10),
                lineColor: getFill(dataViews[0], 'statistics', 'limitLineColor', '#FFA500'),
                lineStyle: getValue<string>(dataViews[0].metadata.objects, 'statistics', 'limitLineStyle', '6,4'),
                show: getValue<boolean>(dataViews[0].metadata.objects, 'statistics', 'showLimits', true)
            };
            let xAxisData: AxisData = {
                AxisTitle: getValue<string>(dataViews[0].metadata.objects, 'xAxis', 'xAxisTitle', 'Default Value'),
                TitleColor: getFill(dataViews[0], 'xAxis', 'xAxisTitleColor', '#A9A9A9'),
                TitleSize: getValue<number>(dataViews[0].metadata.objects, 'xAxis', 'xAxisTitleSize', 12),
                AxisLabelSize: getValue<number>(dataViews[0].metadata.objects, 'xAxis', 'xAxisLabelSize', 12),
                AxisLabelColor: getFill(dataViews[0], 'xAxis', 'xAxisLabelColor', '#2F4F4F'),
                AxisFormat: xAxisFormat
            };
            let yAxisData: AxisData = {
                AxisTitle: getValue<string>(dataViews[0].metadata.objects, 'yAxis', 'yAxisTitle', 'Default Value'),
                TitleColor: getFill(dataViews[0], 'yAxis', 'yAxisTitleColor', '#A9A9A9'),
                TitleSize: getValue<number>(dataViews[0].metadata.objects, 'yAxis', 'yAxisTitleSize', 12),
                AxisLabelSize: getValue<number>(dataViews[0].metadata.objects, 'yAxis', 'yAxisLabelSize', 12),
                AxisLabelColor: getFill(dataViews[0], 'yAxis', 'yAxisLabelColor', '#2F4F4F'),
                AxisFormat: getValue<string>(dataViews[0].metadata.objects, 'yAxis', 'yAxisFormat', '.3s')
            };

            var mRange: number = getValue<number>(dataViews[0].metadata.objects, 'statistics', 'movingRange', 2);
            if (mRange < 2 || mRange > 50)
                mRange = 2
            else
                mRange = Math.round(mRange);

            return {
                dataPoints: ChartDataPoints,
                minX: null,
                maxX: null,
                minY: 0,
                maxY: 0,
                data: chartData,
                xAxis: xAxisData,
                yAxis: yAxisData,
                stageDividerLine: stageDividerLine,
                limitLine: limitLine,
                meanLine: meanLine,
                standardDeviations: getValue<number>(dataViews[0].metadata.objects, 'statistics', 'standardDeviations', 3),
                showGridLines: getValue<boolean>(dataViews[0].metadata.objects, 'chart', 'showGridLines', true),
                isDateRange: isDateRange,
                runRule1: getValue<boolean>(dataViews[0].metadata.objects, 'rules', 'runRule1', false),
                runRule2: getValue<boolean>(dataViews[0].metadata.objects, 'rules', 'runRule2', false),
                runRule3: getValue<boolean>(dataViews[0].metadata.objects, 'rules', 'runRule3', false),
                ruleColor: getFill(dataViews[0], 'rules', 'ruleColor', '#FFFF00'),
                movingRange: mRange,
                mRError: false
            }
        }
        else {
            return viewModel;
        }
    }

    export class ControlChart implements IVisual {
        private svg: d3.Selection<SVGElement>;
        private host: IVisualHost;
        private Container: d3.Selection<SVGElement>;
        private DataPoints: ChartDataPoint[];
        private dataView: DataView;
        private chartStages: Stage[];
        private controlChartViewModel: ControlChartViewModel;
        private svgRoot: d3.Selection<SVGElementInstance>;
        private svgGroupMain: d3.Selection<SVGElementInstance>;
        private padding: number = 12;
        private plot;
        private xScale;
        private yScale;
        private meanLine = [];
        private uclLines = [];
        private lclLines = [];
        private stageDividers = [];
        private dots;
        private tooltipServiceWrapper: ITooltipServiceWrapper;

        constructor(options: VisualConstructorOptions) {
            this.host = options.host;
            this.svgRoot = d3.select(options.element).append('svg').classed('controlChart', true);
            this.svgGroupMain = this.svgRoot.append("g").classed('Container', true);
            this.tooltipServiceWrapper = createTooltipServiceWrapper(this.host.tooltipService, options.element);
        }

        public update(options: VisualUpdateOptions) {
            var categorical = options.dataViews[0].categorical;           
            // remove all existing SVG elements 
            this.svgGroupMain.selectAll("*").remove();
            this.svgRoot.empty();
            
            if (typeof categorical.categories === "undefined" || typeof categorical.values === "undefined")              
                return;
            
            // get categorical data from visual data view
            this.dataView = options.dataViews[0];
            // convert categorical data into specialized data structure for data binding
            this.controlChartViewModel = visualTransform(options, this.host);
            this.svgRoot
                .attr("width", options.viewport.width)
                .attr("height", options.viewport.height);

            if (this.controlChartViewModel && this.controlChartViewModel.dataPoints[0]) {
                this.GetStages();                                                   //determine stage groups
                this.CalcStats();                                                   //calc mean and sd
                this.CreateAxes(options.viewport.width, options.viewport.height);
                if (this.controlChartViewModel.meanLine.show)
                    this.PlotMean();                                                //mean line
                if (this.controlChartViewModel.stageDividerLine.show)
                    this.DrawStageDividers();                                       //stage changes
                if (this.controlChartViewModel.limitLine.show)
                    this.PlotControlLimits();                                       //lcl and ucl
                this.PlotData();                                                    //plot basic raw data
                this.ApplyRules();
                this.DrawMRWarning();

                //put border around plot area
                var plot = this.plot;
                var borderPath = this.svgGroupMain.append("rect")
                    .attr("x", 0)
                    .attr("y", 0)
                    .attr("height", plot.height)
                    .attr("width", plot.width)
                    .style("stroke", "grey")
                    .style("fill", "none")
                    .style("stroke-width", 1);
            }
        }

        private CreateAxes(viewPortWidth: number, viewPortHeight: number) {
            var xAxisOffset: number = 60;
            var yAxisOffset: number = 54;

            var plot = {
                xAxisOffset: xAxisOffset,
                yAxisOffset: yAxisOffset,
                xOffset: this.padding + xAxisOffset,
                yOffset: this.padding,
                width: viewPortWidth - (this.padding + xAxisOffset) * 2,  //  (this.padding * 2) - (2 * xAxisOffset),
                height: viewPortHeight - (this.padding * 2) - yAxisOffset,
            };
            this.plot = plot;

            this.svgGroupMain.attr({
                height: plot.height,
                width: plot.width,
                transform: 'translate(' + plot.xOffset + ',' + plot.yOffset + ')'
            });

            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let vmXaxis = viewModel.xAxis;
            let vmYaxis = viewModel.yAxis;

            this.GetMinMaxX();
            var xScale;
            var dateFormat;
            dateFormat = d3.time.format(vmXaxis.AxisFormat);
            if (viewModel.isDateRange) {
                xScale = d3.time.scale()
                    .range([0, plot.width])
                    .domain([viewModel.minX, viewModel.maxX]);
            }
            else {
                xScale = d3.scale.linear()
                    .range([0, plot.width])
                    .domain([viewModel.minX, viewModel.maxX]);
                dateFormat = d3.format(vmXaxis.AxisFormat);
            }

            this.xScale = xScale;
            //this.svgRoot.selectAll('.axis').remove();
            // draw x axis
            var xAxis = d3.svg.axis()
                .scale(xScale)
                .orient('bottom')
                .innerTickSize(-plot.height)
                .tickPadding(14)
                .tickFormat(dateFormat);
            if (!viewModel.showGridLines) {
                xAxis.innerTickSize(0);
            }

            this.svgGroupMain
                .append('g')
                .attr('class', 'x axis')
                .style('fill', vmXaxis.AxisLabelColor)
                .attr('transform', 'translate(0,' + plot.height + ')')
                .call(xAxis)
                .style("font-size", vmXaxis.AxisLabelSize + 'px');
            var xGridlineAxis = d3.svg.axis()
                .scale(xScale)
                .orient('bottom')
                .innerTickSize(8)
                .tickPadding(10)
                .outerTickSize(6);
            this.svgGroupMain
                .append('g')
                .attr('class', 'gridLine')
                .attr('transform', 'translate(0,' + plot.height + ')')
                .call(xGridlineAxis)
                .style("font-size", '0px');
            //handle uCL and lCL
            var dataMax = d3.max(viewModel.dataPoints, function (d) { return d.yValue });
            var dataMin = d3.min(viewModel.dataPoints, function (d) { return d.yValue });
            var lclLines = this.lclLines;
            var uclLines = this.uclLines;
            var yMax: number = d3.max(uclLines, function (d) { return d['y1'] });
            var yMin: number = d3.min(lclLines, function (d) { return d['y1'] });
            yMin = Math.min(dataMin, yMin);
            if (yMin < 0)
                yMin = yMin * 1.05;
            else
                yMin = yMin * 0.95;

            yMax = Math.max(dataMax, yMax);
            if (yMax > 0)
                yMax = yMax * 1.05;
            else
                yMax = yMax * 0.95;

            // draw y axis
            var yScale = d3.scale.linear()
                .range([plot.height, 0])
                .domain([yMin, yMax])
                .nice();
            this.yScale = yScale;
            this.controlChartViewModel.minY = yMin;
            this.controlChartViewModel.maxY = yMax;

            var yformatValue = d3.format(vmYaxis.AxisFormat);
            var yAxis = d3.svg.axis()
                .scale(yScale)
                .orient('left')
                .innerTickSize(-plot.width)
                .ticks(8)
                .tickPadding(10)
                .tickFormat(function (d) { return yformatValue(d) });

            if (!viewModel.showGridLines) {
                yAxis
                    .innerTickSize(0)
                    .tickPadding(10);
            }

            this.svgGroupMain
                .append('g')
                .attr('class', 'y axis')
                .style('fill', vmYaxis.AxisLabelColor)
                .style("font-size", vmYaxis.AxisLabelSize + 'px')
                .call(yAxis);
            var yGridlineAxis = d3.svg.axis()
                .scale(yScale)
                .orient('left')
                .innerTickSize(8)
                .tickPadding(10)
                .outerTickSize(6);
            this.svgGroupMain
                .append('g')
                .attr('class', 'gridLine')
                .call(yGridlineAxis)
                .style("font-size", '0px');
            this.svgGroupMain.append("text")
                .attr("transform", "rotate(-90)")
                .attr("y", 0 - xAxisOffset - 10)
                .attr("x", 0 - (plot.height / 2))
                .attr("dy", "1em")
                .style("text-anchor", "middle")
                .style("font-size", vmYaxis.TitleSize + 'px')
                .style("fill", vmYaxis.TitleColor)
                .text(vmYaxis.AxisTitle);
            this.svgGroupMain.append("text")
                .attr("y", plot.height + yAxisOffset)
                .attr("x", (plot.width / 2))
                .style("text-anchor", "middle")
                .style("font-size", vmXaxis.TitleSize + 'px')
                .style("fill", vmXaxis.TitleColor)
                .text(vmXaxis.AxisTitle);
        }

        private GetMinMaxX() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let data = viewModel.dataPoints;
            var maxValue: any;
            var minValue: any;

            if (viewModel.isDateRange) {
                maxValue = new Date(data[0].xValue);
                minValue = new Date(data[0].xValue);
            }
            else {
                maxValue = new Number(data[0].xValue);
                minValue = new Number(data[0].xValue);
            }
            for (var i in data) {
                var dt = data[i].xValue;
                if (maxValue < dt) {
                    maxValue = dt;
                }
                if (minValue > dt) {
                    minValue = dt;
                }
            }
            this.controlChartViewModel.minX = minValue;
            this.controlChartViewModel.maxX = maxValue;
        }

        private PlotData() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let data = viewModel.data;
            var point = [];
            for (let i = 0; i < viewModel.dataPoints.length; i++) {
                var ob = viewModel.dataPoints[i];
                var x = ob.xValue;
                var y = ob.yValue;
                //creating line points                         
                point.push([x, y]);
            }
            // Line      
            var xScale = this.xScale;
            var yScale = this.yScale;
            var d3line2 = d3.svg.line()
                .x(function (d) { return xScale(d[0]) })
                .y(function (d) { return yScale(d[1]) });
            //add line
            this.svgGroupMain.append("svg:path").classed('trend_Line', true)
                .attr("d", d3line2(point))
                .style("stroke-width", '1.5px')
                .style({ "stroke": data.LineColor, "stroke-dasharray": (data.LineStyle) })
                .style("fill", 'none');
            //add dots
            var dots = this.svgGroupMain.attr("id", "groupOfCircles").selectAll("dot")
                .data(point)
                .enter().append("circle")
                .style("fill", data.DataColor)
                .attr("r", 2.5)
                .attr("cx", function (d) { return xScale(d[0]); })
                .attr("cy", function (d) { return yScale(d[1]); });
            this.dots = dots;

            //add tooltip
            var xFormat;
            if (viewModel.isDateRange)
                xFormat = d3.time.format(viewModel.xAxis.AxisFormat);
            else
                xFormat = d3.format(viewModel.xAxis.AxisFormat);
            var yFormat = d3.format(viewModel.yAxis.AxisFormat);
            this.tooltipServiceWrapper.addTooltip(dots,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipData(tooltipEvent.data, data.DataColor, xFormat, yFormat),
                (tooltipEvent: TooltipEventArgs<number>) => null);
        }

        private static getTooltipData(value: any, datacolor: string, xFormat: any, yFormat: any): VisualTooltipDataItem[] {
            return [{
                displayName: xFormat(value[0]).toString(),
                value: yFormat(value[1]).toString(),
                color: datacolor
            }];
        }

        private DrawCrossHairs() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            //     let crossHairLine = viewModel.crossHairLine;
            var xScale = this.xScale;
            var yScale = this.yScale;
            //add focus lines and circle
            var plot = this.plot;
            var point = [];
            for (let i = 0; i < viewModel.dataPoints.length; i++) {
                var ob = viewModel.dataPoints[i];
                var dtDate: any = new Date(ob.xValue);
                var x = ob.xValue;
                var y = ob.yValue;
                //creating line points                         
                point.push([x, y]);
            }
            var focus = this.svgGroupMain.append("g")
                .style("display", "none");
            focus.append("circle")
                .attr("class", "y")
                .style("fill", "none")
                .style("stroke", "red")
                .attr("id", "focuscircle")
                .attr("r", 4);
            focus.append('line')
                .attr('id', 'focusLineX')
                .attr('class', 'focusLine');
            focus.append('line')
                .attr('id', 'focusLineY')
                .attr('class', 'focusLine');
            focus.append("text")
                .attr('id', 'labelText')
                .attr("x", 9)
                .attr("dy", ".35em");
            focus.append("text")
                .attr('id', 'yAxisText')
                .attr("dy", ".35em");
            focus.append("text")
                .attr('id', 'xAxisText')
                .attr("dx", ".15em");
            // append the rectangle to capture mouse
            this.svgGroupMain.append("rect")
                .attr("width", plot.width)
                .attr("height", plot.height)
                .style("fill", "none")
                .style("pointer-events", "all")
                .on("mouseover", function () { focus.style("display", "null"); })
                .on("mouseout", function () { focus.style("display", "none"); })
                .on("mousemove", mousemove);
            var bisectDate = d3.bisector(function (d) { return d[0]; }).left;

            function mousemove() {
                var x0 = xScale.invert(d3.mouse(this)[0]);
                var i = bisectDate(point, x0)
                var d0 = point[i - 1];
                var d1 = point[i];
                var d = x0 - d0[0] > d1[0] - x0 ? d1 : d0;
                var x = xScale(d[0]);
                var y = yScale(d[1]);

                focus.select('#focuscircle')
                    .attr('cx', x)
                    .attr('cy', y);

                var xDomain = [viewModel.minX, viewModel.maxX];
                var yDomain = ([viewModel.minY, viewModel.maxY]);
                var yScale2 = d3.scale.linear().range([plot.height, 0]).domain(yDomain);

                /*        focus.select('#focusLineX')
                            .attr('x1', x).attr('y1', yScale2(yDomain[0]))
                            .attr('x2', x).attr('y2', yScale2(yDomain[1]))
                            .style("stroke", crossHairLine.lineColor);
                        focus.select('#focusLineY')
                            .attr('x1', xScale(xDomain[0])).attr('y1', y)
                            .attr('x2', xScale(xDomain[1])).attr('y2', y)
                            .style("stroke", crossHairLine.lineColor);
        
                        var dateFormat;
                        if (viewModel.isDateRange)
                            dateFormat = d3.time.format(viewModel.xAxis.AxisFormat);
                        else
                            dateFormat = d3.format(viewModel.xAxis.AxisFormat);
        
                        var yAxisFormat = d3.format(viewModel.yAxis.AxisFormat);
                        focus.select('#yAxisText')
                            .text(yAxisFormat(d[1]).toString())
                            .attr("y", y)
                            .style("text-anchor", "end")
                            .attr("dx", "-.15em")
                            .style("font-size", crossHairLine.textSize + 'px')
                            .style("fill", crossHairLine.textColor);
        
                        focus.select('#xAxisText')
                            .text(dateFormat(d[0]).toString())
                            .attr("y", plot.height)
                            .attr("x", x)
                            .style("text-anchor", "middle")
                            .attr("dy", ".95em")
                            .style("font-size", crossHairLine.textSize + 'px')
                            .style("fill", crossHairLine.textColor);*/
            }
        }

        private GetStages() {
            let stage: Stage = {
                uCL: 0,
                lCL: 0,
                mean: 0,
                sum: 0,
                startX: null,
                endX: null,
                stage: '',
                count: 0,
                stageDividerX: null,
                firstId: 0,
                lastId: 0,
                mRError: false
            };
            this.chartStages = [];
            let stages: Stage[] = [];

            var dataView = this.dataView;
            if (!dataView
                || !dataView
                || !dataView.categorical
                || !dataView.categorical.values)
                return [];

            var stageValue: string;
            var categorical;
            var hasStages: boolean;
            if (dataView.categorical.values.length == 1 || !dataView.categorical.values[1]) {
                stageValue = '';
                hasStages = false;
            }
            else {
                categorical = dataView.categorical;
                stageValue = <string>categorical.values[1].values[0];
                hasStages = true;
            }
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            stage.startX = viewModel.dataPoints[0].xValue;
            var currentStage: string = stageValue;
            stage.stage = stageValue;
            for (let i = 0; i < viewModel.dataPoints.length; i++) {
                var obj = viewModel.dataPoints[i];

                if (hasStages)
                    stageValue = <string>categorical.values[1].values[i];

                if (currentStage == stageValue || !hasStages) {
                    stage.sum = stage.sum + obj.yValue;
                    stage.count++;
                }
                else {
                    if (stage.count > 0)
                        stage.mean = stage.sum / stage.count;
                    if (i > 0)
                        stage.endX = viewModel.dataPoints[i - 1].xValue;
                    else
                        stage.endX = obj.xValue;

                    var nextStartDate: any = <any>(obj.xValue);
                    var endDate: any = (viewModel.dataPoints[i - 1].xValue);
                    stage.stageDividerX = ((nextStartDate.valueOf() + endDate.valueOf()) / 2);

                    stages.push({
                        startX: stage.startX,
                        endX: stage.endX,
                        count: stage.count,
                        mean: stage.mean,
                        sum: stage.sum,
                        uCL: 0,
                        lCL: 0,
                        stage: stage.stage,
                        stageDividerX: stage.stageDividerX,
                        firstId: stage.firstId,
                        lastId: i - 1,
                        mRError: false
                    });
                    stage.mean = obj.yValue;
                    stage.count = 1;
                    stage.sum = obj.yValue;
                    stage.startX = obj.xValue;
                    stage.firstId = i;
                    stage.stage = stageValue;
                }
                //set stage to last stage value
                if (hasStages)
                    currentStage = <string>categorical.values[1].values[i];

                //last point
                if (i == (viewModel.dataPoints.length - 1)) {
                    if (stage.count > 0)
                        stage.mean = stage.sum / stage.count;
                    stage.endX = obj.xValue;
                    stage.stage = stageValue;
                    stages.push({
                        startX: stage.startX,
                        endX: stage.endX,
                        count: stage.count,
                        mean: stage.mean,
                        sum: stage.sum,
                        uCL: 0,
                        lCL: 0,
                        stage: stage.stage,
                        stageDividerX: stage.endX,
                        firstId: stage.firstId,
                        lastId: i,
                        mRError: false
                    });
                }
            }
            this.chartStages = stages;
        }

        private PlotMean() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let mLine = viewModel.meanLine;
            var xScale = this.xScale;
            var yScale = this.yScale;
            var meanLine = this.meanLine;

            var mean = this.svgGroupMain.selectAll("meanLine")
                .data(meanLine)
                .enter().append("polyline")
                .attr("points", function (d) { return xScale(d['x1']).toString() + "," + yScale(d['y2']).toString() + "," + xScale(d['x2']).toString() + "," + yScale(d['y2']).toString() })
                .style({ "stroke": mLine.lineColor, "stroke-width": 1.5, "stroke-dasharray": (mLine.lineStyle) });

            // x&#x0304;
            var xbar = 0x0304;
            var yformatValue = d3.format(viewModel.yAxis.AxisFormat);

            var stages = this.chartStages;
            var plot = this.plot;
            var meanText = this.svgGroupMain.selectAll("meanText")
                .data(meanLine)
                .enter().append("text")
                .attr("x", function (d) { if (stages.length == 1) return plot.width; else { return xScale(d['x1']).toString() } })
                .attr("y", function (d) { return yScale(d['y2']).toString() })
                .attr("dx", ".15em")
                .attr("dy", function (d) { if (stages.length == 1) return ".30em"; else return "-.25em" })
                .attr("text-anchor", "start")
                .text(function (d) { return 'x' + String.fromCharCode(xbar) + ' = ' + yformatValue(d['y2']).toString() })
                .style("font-size", mLine.textSize + 'px')
                .style("fill", mLine.textColor);

            //add tooltip
            var xFormat;
            if (viewModel.isDateRange)
                xFormat = d3.time.format(viewModel.xAxis.AxisFormat);
            else
                xFormat = d3.format(viewModel.xAxis.AxisFormat);

            this.tooltipServiceWrapper.addTooltip(mean,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipMeanData(tooltipEvent.data, mLine.textColor, 'Mean', yformatValue),
                (tooltipEvent: TooltipEventArgs<number>) => null);
            this.tooltipServiceWrapper.addTooltip(meanText,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipMeanData(tooltipEvent.data, mLine.textColor, 'Mean', yformatValue),
                (tooltipEvent: TooltipEventArgs<number>) => null);

        }

        private static getTooltipMeanData(value: any, datacolor: string, label: any, yFormat: any): VisualTooltipDataItem[] {
            return [{
                displayName: label,
                value: yFormat(value['y2']).toString(),
                color: datacolor
            }];
        }

        private CalcStats() {
            let stages = this.chartStages;
            var meanLine = [];
            var uclLines = [];
            var lclLines = [];
            var uCL: number;
            var lCL: number;
            var stageDividers = [];
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let data = viewModel.dataPoints;
            var numsds = viewModel.standardDeviations;
            var mr: number = viewModel.movingRange;
            var mrSum: number;
            var d2: number = this.LookUpd2(viewModel.movingRange);
            var mrMin: number;
            var mrMax: number;
            var mrRangeSum: number;
            var rBar: number;
            var lCL: number;
            var uCL: number;

            for (let i = 0; i < stages.length; i++) {
                mrRangeSum = 0;
                if (mr > stages[i].count) {
                    //cannot process stage if moving range is greater than number of items in stage?
                    stages[i].mRError = true;
                    viewModel.mRError = true;
                }
                else {
                    for (let j = stages[i].firstId + mr - 1; j <= stages[i].lastId; j++) {
                        //for each data point in the stage
                        mrMin = data[j - 1].yValue;
                        mrMax = data[j - 1].yValue;
                        //move thru stage data points using mr window 
                        for (let k = j; k >= j - mr + 1; k--) {
                            if (mrMin > data[k].yValue)
                                mrMin = data[k].yValue;
                            if (mrMax < data[k].yValue)
                                mrMax = data[k].yValue;
                        }
                        mrRangeSum = mrRangeSum + Math.abs(mrMax - mrMin);
                    }
                    rBar = mrRangeSum / (stages[i].count - mr + 1);
                    uCL = stages[i].mean + numsds * rBar / d2;
                    lCL = stages[i].mean - numsds * rBar / d2;
                    stages[i].uCL = uCL;
                    stages[i].lCL = lCL;
                }
                if (i > 0) {
                    meanLine.push({ x1: stages[i - 1].stageDividerX, y1: stages[i - 1].mean, x2: stages[i].stageDividerX, y2: stages[i].mean });
                    stageDividers.push({ x1: stages[i].stageDividerX, prevDividerX: stages[i - 1].stageDividerX, stageName: stages[i].stage });
                    if (!stages[i].mRError) {
                        uclLines.push({ x1: stages[i - 1].stageDividerX, y1: uCL, x2: stages[i].stageDividerX, y2: uCL });
                        lclLines.push({ x1: stages[i - 1].stageDividerX, y1: lCL, x2: stages[i].stageDividerX, y2: lCL });
                    }
                }
                else {
                    meanLine.push({ x1: stages[i].startX, y1: stages[i].mean, x2: stages[i].stageDividerX, y2: stages[i].mean });
                    stageDividers.push({ x1: stages[i].stageDividerX, prevDividerX: stages[i].startX, stageName: stages[i].stage });
                    if (!stages[i].mRError) {
                        uclLines.push({ x1: stages[i].startX, y1: uCL, x2: stages[i].stageDividerX, y2: uCL });
                        lclLines.push({ x1: stages[i].startX, y1: lCL, x2: stages[i].stageDividerX, y2: lCL });
                    }
                }
            }
            this.chartStages = stages;
            this.meanLine = meanLine;
            this.lclLines = lclLines;
            this.uclLines = uclLines;
            this.stageDividers = stageDividers;
        }

        private DrawMRWarning() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            var plot = this.plot;
            var yAxisOffset = 54;
            if (viewModel.mRError) {
                this.svgGroupMain.append("text")
                    .attr("y", yAxisOffset)
                    .attr("x", (plot.width / 2))
                    .style("text-anchor", "middle")
                    .style("font-size", '20px')
                    .style("fill", 'red')
                    .text('Selected Moving Range is greater than the number of data points in a stage');
            }
        }

        private DrawStageDividers() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let stageDiv = viewModel.stageDividerLine;
            var plot = this.plot;
            var stageDividers = this.stageDividers;
            var xScale = this.xScale;
            var yScale = this.yScale;
            this.svgGroupMain.selectAll("divider")
                .data(stageDividers)
                .enter().append("polyline")
                .attr("points", function (d) { return xScale(d['x1']).toString() + "," + "0," + xScale(d['x1']).toString() + "," + plot.height.toString() })
                .style({ "stroke": stageDiv.lineColor, "stroke-width": 1.5, "stroke-dasharray": (stageDiv.lineStyle) });
            var dividerText = this.svgGroupMain.selectAll("dividerText")
                .data(stageDividers)
                .enter().append("text")
                .attr("x", function (d) { return xScale(d['prevDividerX']).toString() })
                .attr("y", "0")
                .attr("dx", ".35em")
                .attr("dy", "1em")
                .attr("text-anchor", "start")
                .text(function (d) { return d['stageName'].toString() })
                .style("font-size", stageDiv.textSize + 'px')
                .style("fill", stageDiv.textColor);

            this.tooltipServiceWrapper.addTooltip(dividerText,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipDividerText(tooltipEvent.data, stageDiv.textColor),
                (tooltipEvent: TooltipEventArgs<number>) => null);

        }

        private static getTooltipDividerText(value: any, datacolor: string): VisualTooltipDataItem[] {
            return [{
                displayName: 'Subgroup',
                value: value['stageName'].toString(),
                color: datacolor
            }];
        }

        private PlotControlLimits() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let limitLine = viewModel.limitLine;
            var lclLines = this.lclLines;
            var uclLines = this.uclLines;
            var xScale = this.xScale;
            var yScale = this.yScale;
            var uclLine = this.svgGroupMain.selectAll("uclLine")
                .data(uclLines)
                .enter().append("polyline")
                .attr("points", function (d) { return xScale(d['x1']).toString() + "," + yScale(d['y1']).toString() + "," + xScale(d['x2']).toString() + "," + yScale(d['y2']).toString() })
                .style({ "stroke": limitLine.lineColor, "stroke-width": 1.5, "stroke-dasharray": (limitLine.lineStyle) });
            var lclLine = this.svgGroupMain.selectAll("lclLine")
                .data(lclLines)
                .enter().append("polyline")
                .attr("points", function (d) { return xScale(d['x1']).toString() + "," + yScale(d['y1']).toString() + "," + xScale(d['x2']).toString() + "," + yScale(d['y2']).toString() })
                .style({ "stroke": limitLine.lineColor, "stroke-width": 1.5, "stroke-dasharray": (limitLine.lineStyle) });
            var yformatValue = d3.format(viewModel.yAxis.AxisFormat);

            var stages = this.chartStages;
            var plot = this.plot;

            var uclText = this.svgGroupMain.selectAll("uclText")
                .data(uclLines)
                .enter().append("text")
                .attr("x", function (d) { if (stages.length == 1) return plot.width; else { return xScale(d['x1']).toString() } })
                .attr("y", function (d) { return yScale(d['y1']).toString() })
                .attr("dx", ".15em")
                .attr("dy", function (d) { if (stages.length == 1) return ".30em"; else return "-.25em" })
                .attr("text-anchor", "start")
                .text(function (d) { return 'UCL = ' + yformatValue(d['y1']).toString() })
                .style("font-size", limitLine.textSize + 'px')
                .style("fill", limitLine.textColor);

            var lclText = this.svgGroupMain.selectAll("lclText")
                .data(lclLines)
                .enter().append("text")
                .attr("x", function (d) { if (stages.length == 1) return plot.width; else { return xScale(d['x1']).toString() } })
                .attr("y", function (d) { return yScale(d['y1']).toString() })
                .attr("dx", ".15em")
                .attr("dy", function (d) { if (stages.length == 1) return ".30em"; else return ".95em" })
                .attr("text-anchor", "start")
                .text(function (d) { return 'LCL = ' + yformatValue(d['y1']).toString() })
                .style("font-size", limitLine.textSize + 'px')
                .style("fill", limitLine.textColor);

            //add tooltip
            var yFormat = d3.format(viewModel.yAxis.AxisFormat);
            this.tooltipServiceWrapper.addTooltip(lclLine,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipLimitData(tooltipEvent.data, limitLine.textColor, 'LCL', yFormat),
                (tooltipEvent: TooltipEventArgs<number>) => null);

            this.tooltipServiceWrapper.addTooltip(uclLine,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipLimitData(tooltipEvent.data, limitLine.textColor, 'UCL', yFormat),
                (tooltipEvent: TooltipEventArgs<number>) => null);

            this.tooltipServiceWrapper.addTooltip(lclText,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipLimitData(tooltipEvent.data, limitLine.textColor, 'LCL', yFormat),
                (tooltipEvent: TooltipEventArgs<number>) => null);

            this.tooltipServiceWrapper.addTooltip(uclText,
                (tooltipEvent: TooltipEventArgs<number>) => ControlChart.getTooltipLimitData(tooltipEvent.data, limitLine.textColor, 'UCL', yFormat),
                (tooltipEvent: TooltipEventArgs<number>) => null);

        }

        private static getTooltipLimitData(value: any, datacolor: string, label: any, yFormat: any): VisualTooltipDataItem[] {
            return [{
                displayName: label,
                value: yFormat(value['y1']).toString(),
                color: datacolor
            }];
        }

        private ApplyRules() {
            let chartStages = this.chartStages;
            //rule 1 - highlight over/below UCL/LCL
            let ChartDataPoints = this.controlChartViewModel.dataPoints;
            var datalen = ChartDataPoints.length;
            let dots = this.dots;
            var consecutiveUPoints = [];
            var consecutiveLPoints = [];
            var consecIncPoints = [];
            var consecDecPoints = [];
            var meanPoints = [];
            let stages = this.chartStages;
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let data = viewModel.dataPoints;
            for (let i = 0; i < stages.length; i++) {
                for (let j = stages[i].firstId; j <= stages[i].lastId; j++) {
                    //Rule #1
                    if (viewModel.runRule1)
                        if (stages[i].uCL < data[j].yValue || stages[i].lCL > data[j].yValue)
                            meanPoints.push([j]);

                    //Rule #2 - over 5 incr or decr
                    if (viewModel.runRule2) {
                        if (j > 0) {
                            if (data[j].yValue > data[j - 1].yValue)
                                consecIncPoints.push([j])
                            else {
                                if (consecIncPoints.length > 5)
                                    this.DrawRulePoints(consecIncPoints);
                                consecIncPoints = [];
                                consecIncPoints.push([j]);
                            }
                            if (data[j].yValue < data[j - 1].yValue)
                                consecDecPoints.push([j])
                            else {
                                if (consecDecPoints.length > 5)
                                    this.DrawRulePoints(consecDecPoints);
                                consecDecPoints = [];
                                consecDecPoints.push([j]);
                            }
                        }
                        else {
                            consecIncPoints.push([j]);
                            consecDecPoints.push([j]);
                        }
                    }

                    //Rule #3
                    if (viewModel.runRule3)
                        if (data[j].yValue > stages[i].mean) {
                            if (consecutiveLPoints.length > 8)
                                this.DrawRulePoints(consecutiveLPoints);
                            consecutiveLPoints = [];
                            consecutiveUPoints.push([j]);
                        }
                        else
                            if (data[j].yValue < stages[i].mean) {
                                if (consecutiveUPoints.length > 8)
                                    this.DrawRulePoints(consecutiveUPoints);
                                consecutiveUPoints = [];
                                consecutiveLPoints.push([j]);
                            }
                }
                if (meanPoints.length > 0) {
                    this.DrawRulePoints(meanPoints);
                    meanPoints = [];
                }
                if (consecutiveLPoints.length > 8)
                    this.DrawRulePoints(consecutiveLPoints);
                if (consecutiveUPoints.length > 8)
                    this.DrawRulePoints(consecutiveUPoints);
                consecutiveLPoints = [];
                consecutiveUPoints = [];

                if (consecIncPoints.length > 5)
                    this.DrawRulePoints(consecIncPoints);
                if (consecDecPoints.length > 5)
                    this.DrawRulePoints(consecDecPoints);
                consecDecPoints = [];
                consecIncPoints = [];
            }
        }

        private DrawRulePoints(points: any) {
            let dots = this.dots;
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            for (let i = 0; i < points.length; i++) {
                d3.select(dots[0][points[i]]).style("fill", viewModel.ruleColor);
            }
        }

        private LookUpd2(mr: number): number {
            var d2Array = [1,
                1.128,
                1.693,
                2.059,
                2.326,
                2.534,
                2.704,
                2.847,
                2.97,
                3.078,
                3.173,
                3.258,
                3.336,
                3.407,
                3.472,
                3.532,
                3.588,
                3.64,
                3.689,
                3.735,
                3.778,
                3.819,
                3.858,
                3.895,
                3.931,
                3.964,
                3.997,
                4.027,
                4.057,
                4.086,
                4.113,
                4.139,
                4.165,
                4.189,
                4.213,
                4.236,
                4.259,
                4.28,
                4.301,
                4.322,
                4.341,
                4.361,
                4.379,
                4.398,
                4.415,
                4.433,
                4.45,
                4.466,
                4.482,
                4.498];
            if (mr > 0)
                return d2Array[mr - 1];
            else
                return 1;
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] {
            var instances: VisualObjectInstance[] = [];
            var viewModel = this.controlChartViewModel;
            var objectName = options.objectName;

            switch (objectName) {
                case 'chart':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            dataColor: viewModel.data.DataColor,
                            lineColor: viewModel.data.LineColor,
                            lineStyle: viewModel.data.LineStyle,
                            showGridLines: viewModel.showGridLines
                        }
                    };
                    instances.push(config);
                    break;
                case 'statistics':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            showDividers: viewModel.stageDividerLine.show,
                            stageDividerColor: viewModel.stageDividerLine.lineColor,
                            stageLabelColor: viewModel.stageDividerLine.textColor,
                            stageDividerLineStyle: viewModel.stageDividerLine.lineStyle,
                            stageDividerLabelSize: viewModel.stageDividerLine.textSize,
                            showLimits: viewModel.limitLine.show,
                            limitLineColor: viewModel.limitLine.lineColor,
                            limitLabelColor: viewModel.limitLine.textColor,
                            limitLabelSize: viewModel.limitLine.textSize,
                            limitLineStyle: viewModel.limitLine.lineStyle,
                            showMean: viewModel.meanLine.show,
                            meanLineColor: viewModel.meanLine.lineColor,
                            meanLabelColor: viewModel.meanLine.textColor,
                            meanLabelSize: viewModel.meanLine.textSize,
                            meanLineStyle: viewModel.meanLine.lineStyle,
                            standardDeviations: viewModel.standardDeviations,
                            movingRange: viewModel.movingRange
                        }
                    };
                    instances.push(config);
                    break;
                case 'xAxis':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            xAxisTitle: viewModel.xAxis.AxisTitle,
                            xAxisTitleColor: viewModel.xAxis.TitleColor,
                            xAxisTitleSize: viewModel.xAxis.TitleSize,
                            xAxisLabelColor: viewModel.xAxis.AxisLabelColor,
                            xAxisLabelSize: viewModel.xAxis.AxisLabelSize,
                            xAxisFormat: viewModel.xAxis.AxisFormat
                        }
                    };
                    instances.push(config);
                    break;
                case 'yAxis':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            yAxisTitle: viewModel.yAxis.AxisTitle,
                            yAxisTitleColor: viewModel.yAxis.TitleColor,
                            yAxisTitleSize: viewModel.yAxis.TitleSize,
                            yAxisLabelColor: viewModel.yAxis.AxisLabelColor,
                            yAxisLabelSize: viewModel.yAxis.AxisLabelSize,
                            yAxisFormat: viewModel.yAxis.AxisFormat
                        }
                    };
                    instances.push(config);
                    break;
                case 'rules':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            runRule1: viewModel.runRule1,
                            runRule2: viewModel.runRule2,
                            runRule3: viewModel.runRule3,
                            ruleColor: viewModel.ruleColor
                        }
                    };
                    instances.push(config);
                    break;
            }
            return instances;
        }
    }
}