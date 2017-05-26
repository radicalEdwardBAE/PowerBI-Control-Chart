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
         * @property {any} minY                     - minimum value of Y axis - can be date or number
         * @property {any} maxY                     - maximum value of Y axis - can be date or number
         * @property {interface} data               - LineData - data point line
         * @property {interface} stageDividerLine   - StatisticsData - contains info on line and labels
         * @property {interface} limitLine          - StatisticsData - contains info on line and labels
         * @property {interface} meanLine           - StatisticsData - contains info on line and labels
         * @property {interface} crossHairLine      - StatisticsData - contains info on line and labels
         * @property {interface} xAxis              - AxisData - contains info on labels
         * @property {interface} yAxis              - minimum value of X axis - can be date or number
         * @property {boolean} showGridLines        - show gridlines
         * @property {boolean} isDateRange          - is the X axis a date or numeric range?
         * @property {boolean} runRule1              - run rule 1
         * @property {boolean} runRule2              - run rule 2
         * @property {boolean} runRule3              - run rule 3
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
        standardDeviation: number;
        crossHairLine: StatisticsData;
        xAxis: AxisData;
        yAxis: AxisData;
        showGridLines: boolean;
        isDateRange: boolean;
        runRule1: boolean;
        runRule2: boolean;
        runRule3: boolean;
        ruleColor: string;
        useSDSubGroups: boolean;
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
     * @property {Object} xvalue    - Data value for point. - date or number
     * @property {number} yValue            - y axis value.
     */
    interface ChartDataPoint {
        xValue: Object;
        yValue: number;
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
         * @property {number} sd            - standard deveiation of stage
         * @property {string} stage         - Stage name
         * @property {any} stageDividerX    - stage divider x axis value - date or number
         * @property {number} firstId       - first index value in a stage - use this in calculating stats
         * @property {number} lastId        - first index value in a stage - use this in calculating stats
         */
    interface Stage {
        lCL: number;
        uCL: number;
        mean: number;
        startX: any;
        endX: any;
        sum: number;
        count: number;
        sd: number;
        stage: string;
        stageDividerX: any;
        firstId: number;
        lastId: number;
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
            crossHairLine: null,
            standardDeviation: 3,
            showGridLines: true,
            isDateRange: true,
            runRule1: false,
            runRule2: false,
            runRule3: false,
            ruleColor: null,
            useSDSubGroups: true
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
        for (let i = 0; i < xValues.length; i++) {
            ChartDataPoints.push({
                xValue: xValues[i],
                yValue: <number>yValues[i]
            });
        }
        //ChartDataPoints.sort( (cat1, cat2) => { return cat2.Value - cat1.Value; })
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
            lineStyle: getValue<string>(dataViews[0].metadata.objects, 'statistics', 'meanLineStyle', '10,4'),
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
            lineStyle: getValue<string>(dataViews[0].metadata.objects, 'statistics', 'limitLineStyle', '10,4'),
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
        let crossHairData: StatisticsData = {
            lineColor: getFill(dataViews[0], 'crossHairs', 'crossHairLineColor', '#40E0D0'),
            textSize: getValue<number>(dataViews[0].metadata.objects, 'crossHairs', 'crossHairLabelSize', 10),
            show: getValue<boolean>(dataViews[0].metadata.objects, 'crossHairs', 'showCrossHairs', true),
            textColor: getFill(dataViews[0], 'crossHairs', 'crossHairLineColor', '#ADFF2F'),
            lineStyle: '4,4'
        };

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
            crossHairLine: crossHairData,
            standardDeviation: getValue<number>(dataViews[0].metadata.objects, 'statistics', 'standardDeviation', 3),
            showGridLines: getValue<boolean>(dataViews[0].metadata.objects, 'chart', 'showGridLines', true),
            isDateRange: isDateRange,
            runRule1: getValue<boolean>(dataViews[0].metadata.objects, 'rules', 'runRule1', false),
            runRule2: getValue<boolean>(dataViews[0].metadata.objects, 'rules', 'runRule2', false),
            runRule3: getValue<boolean>(dataViews[0].metadata.objects, 'rules', 'runRule3', false),
            ruleColor: getFill(dataViews[0], 'rules', 'ruleColor', '#FFFF00'),
            useSDSubGroups: getValue<boolean>(dataViews[0].metadata.objects, 'statistics', 'useSDSubGroups', true)
        };
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

        constructor(options: VisualConstructorOptions) {
            this.host = options.host;
            this.svgRoot = d3.select(options.element).append('svg').classed('controlChart', true);
            this.svgGroupMain = this.svgRoot.append("g").classed('Container', true);
        }

        public update(options: VisualUpdateOptions) {
            var categorical = options.dataViews[0].categorical;
            if (typeof categorical.categories === "undefined" || typeof categorical.values === "undefined") {
                // remove all existing SVG elements 
                this.svgGroupMain.empty();
                return;
            }

            // get categorical data from visual data view
            this.dataView = options.dataViews[0];
            // convert categorical data into specialized data structure for data binding
            this.controlChartViewModel = visualTransform(options, this.host);
            this.svgRoot
                .attr("width", options.viewport.width)
                .attr("height", options.viewport.height);
            this.svgGroupMain.selectAll("*").remove();

            if (this.controlChartViewModel && this.controlChartViewModel.dataPoints[0]) {
                this.GetStages();           //determine stage groups
                this.CalcStats(this.controlChartViewModel.standardDeviation);          //calc mean and sd
                this.CreateAxes(options.viewport.width, options.viewport.height);
                this.PlotData();            //plot basic raw data
                if (this.controlChartViewModel.meanLine.show)
                    this.PlotMean();            //mean line
                if (this.controlChartViewModel.stageDividerLine.show)
                    this.DrawStageDividers();   //stage changes
                if (this.controlChartViewModel.limitLine.show)
                    this.PlotControlLimits();   //lcl and ucl
                if (this.controlChartViewModel.crossHairLine.show)
                    this.DrawCrossHairs();
                //run rules
                this.ApplyRules();
            }
        }

        private CreateAxes(viewPortWidth: number, viewPortHeight: number) {
            var xAxisOffset = 54;
            var yAxisOffset = 54
            var plot = {
                xOffset: this.padding + xAxisOffset,
                yOffset: this.padding,
                width: viewPortWidth - (this.padding * 2) - xAxisOffset - 54,
                height: viewPortHeight - (this.padding * 2) - yAxisOffset,
            };
            this.plot = plot;

            this.svgGroupMain.attr({
                height: plot.height,
                width: plot.width,
                transform: 'translate(' + plot.xOffset + ',' + plot.yOffset + ')'
            });

            var borderPath = this.svgGroupMain.append("rect")
                .attr("x", 0)
                .attr("y", 0)
                .attr("height", plot.height)
                .attr("width", plot.width)
                .style("stroke", "grey")
                .style("fill", "none")
                .style("stroke-width", 1);

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
            this.svgRoot.selectAll('.axis').remove();
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
                .attr("y", 0 - xAxisOffset)
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
                //                var dValue = data[i].yValue;
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
        }

        private DrawCrossHairs() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let crossHairLine = viewModel.crossHairLine;
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
                .on("mouseover", function () { focus.style("display", null); })
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

                focus.select('#focusLineX')
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

                focus.select('#yAxisText')
                    .text(d[1].toString())
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
                    .style("fill", crossHairLine.textColor);
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
                sd: 0,
                stageDividerX: null,
                firstId: 0,
                lastId: 0
            };
            this.chartStages = [];
            let stages: Stage[] = [];

            var dataView = this.dataView;
            if (!dataView
                || !dataView
                || !dataView.categorical
                || !dataView.categorical.values)
                return [];

            var stageValue;
            var categorical;
            var category;
            var noStages: boolean;
            if (dataView.categorical.values.length === 1) {
                categorical = '';
                category = '';
                stageValue = '';
                noStages = true;
            }
            else {
                categorical = dataView.categorical;
                category = categorical.categories[0];
                stageValue = categorical.values[1].values[0];
                noStages = false;
            }
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            stage.startX = viewModel.dataPoints[0].xValue;
            stage.firstId = 0;
            let dataValue = viewModel.dataPoints[0].yValue;
            var stageCounter = 0;
            var currentStage = stageValue;
            stage.stage = stageValue.toString();
            for (let i = 0; i < viewModel.dataPoints.length; i++) {
                var obj = viewModel.dataPoints[i];

                if (!noStages)
                    stageValue = categorical.values[1].values[i];
                else
                    stageValue = '';

                if (currentStage === stageValue || noStages) {
                    stage.sum = stage.sum + obj.yValue;
                    stage.count++;
                }
                else {
                    if (stage.count > 0) {
                        stage.mean = stage.sum / stage.count;
                    }
                    if (i > 0)
                        stage.endX = viewModel.dataPoints[i - 1].xValue;
                    else
                        stage.endX = obj.xValue;

                    var nextStartDate = <any>(obj.xValue);
                    var endDate = (viewModel.dataPoints[i - 1].xValue);
                    stage.stageDividerX = ((nextStartDate.valueOf() + endDate.valueOf()) / 2);

                    stages.push({
                        startX: stage.startX,
                        endX: stage.endX,
                        count: stage.count,
                        mean: stage.mean,
                        sum: stage.sum,
                        sd: 0,
                        uCL: 0,
                        lCL: 0,
                        stage: stage.stage,
                        stageDividerX: stage.stageDividerX,
                        firstId: stage.firstId,
                        lastId: i - 1
                    });
                    stage.mean = obj.yValue;
                    stage.count = 1;
                    stage.sum = obj.yValue;
                    stage.startX = obj.xValue;
                    stage.firstId = i;
                    stage.stage = stageValue.toString();
                }
                //set stage to last stage value
                if (!noStages)
                    currentStage = categorical.values[1].values[i];
                else
                    currentStage = '';

                //last point
                if (i == (viewModel.dataPoints.length - 1)) {
                    if (stage.count > 0) {
                        stage.mean = stage.sum / stage.count;
                    }
                    stage.endX = obj.xValue;
                    stage.stage = stageValue.toString();
                    stages.push({
                        startX: stage.startX,
                        endX: stage.endX,
                        count: stage.count,
                        mean: stage.mean,
                        sum: stage.sum,
                        sd: 0,
                        uCL: 0,
                        lCL: 0,
                        stage: stage.stage,
                        stageDividerX: stage.endX,
                        firstId: stage.firstId,
                        lastId: i
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
            this.svgGroupMain.selectAll("meanLine")
                .data(meanLine)
                .enter().append("polyline")
                .attr("points", function (d) { return xScale(d['x1']).toString() + "," + yScale(d['y2']).toString() + "," + xScale(d['x2']).toString() + "," + yScale(d['y2']).toString() })
                .style({ "stroke": mLine.lineColor, "stroke-width": 1.5, "stroke-dasharray": (mLine.lineStyle) });

            // x&#x0304;
            var xbar = 0x0304;
            var yformatValue = d3.format(viewModel.yAxis.AxisFormat);
            this.svgGroupMain.selectAll("meanText")
                .data(meanLine)
                .enter().append("text")
                .attr("x", function (d) { return xScale(d['x1']).toString() })
                .attr("y", function (d) { return yScale(d['y2']).toString() })
                .attr("dx", ".35em")
                .attr("dy", "-.25em")
                .attr("text-anchor", "start")
                .text(function (d) { return 'x' + String.fromCharCode(xbar) + ' = ' + yformatValue(d['y2']).toString() })
                .style("font-size", mLine.textSize + 'px')
                .style("fill", mLine.textColor);
        }

        private findIndexByKeyValue(arraytosearch, key, valuetosearch) {
            for (var i = 0; i < arraytosearch.length; i++) {
                if (arraytosearch[i][key] == valuetosearch) {
                    return i;
                }
            }
            return null;
        }

        //add this to GetStages
        private GetUnique(src: Stage[]): SubGroup[] {
            var unique = {};
            var distinct: SubGroup[] = [];
            for (var i in src) {
                if (typeof (unique[src[i].stage]) == "undefined") {
                    distinct.push({ stage: src[i].stage, sum: src[i].sum, count: src[i].count, mean: src[i].mean })
                }
                else {
                    var index = this.findIndexByKeyValue(distinct, "stage", src[i].stage);
                    distinct[index].sum = <number>distinct[index].sum + src[i].sum;
                    distinct[index].count = <number>distinct[index].count + src[i].count;
                    if (distinct[index].count > 0)
                        distinct[index].mean = distinct[index].sum / distinct[index].count;
                }
                unique[src[i].stage] = '';
            }
            return distinct;
        }

        private CalcStats(numsds: number) {
            let stages = this.chartStages;
            var diff: number = 0;
            var sumdiffsqrd: number = 0;
            var sd: number = 0;
            var meanLine = [];
            var uclLines = [];
            var lclLines = [];
            var uCL: number;
            var lCL: number;
            var stageDividers = [];
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let data = viewModel.dataPoints;
            var sdTotal: boolean = viewModel.useSDSubGroups;
            //unique stages
            var subGroups = [];
            if (sdTotal) {
                subGroups = this.GetUnique(stages);
            }

            for (let i = 0; i < stages.length; i++) {
                var index: number;
                if (sdTotal) {
                    index = this.findIndexByKeyValue(subGroups, "stage", stages[i].stage);
                    stages[i].count = subGroups[index].count;
                    stages[i].sum = subGroups[index].sum;
                    stages[i].mean = subGroups[index].mean;
                }
                sumdiffsqrd = 0;
                for (let j = stages[i].firstId; j <= stages[i].lastId; j++) {
                    diff = data[j].yValue - stages[i].mean;
                    sumdiffsqrd = sumdiffsqrd + (diff * diff);
                    if (sdTotal)
                        subGroups[index].sum = subGroups[index].sum + (diff * diff);
                }
                if (!sdTotal)
                    stages[i].sum = sumdiffsqrd;
            }

            //Using subGroups, calculate sds and UCLs and LCLs for each stage and then create lines
            for (let m = 0; m < stages.length; m++) {
                var sum: number;
                var count: number;
                var mean: number;

                if (sdTotal) {
                    var index = this.findIndexByKeyValue(subGroups, "stage", stages[m].stage);
                    sum = subGroups[index].sum;
                    count = subGroups[index].count;
                    mean = subGroups[index].mean
                }
                else {
                    sum = stages[m].sum;
                    count = stages[m].count;
                    mean = stages[m].mean;
                }
                if (count > 8)
                    sd = Math.sqrt(sum / count);
                else
                    sd = Math.sqrt(sum / (count - 1));

                stages[m].uCL = mean + numsds * sd;
                stages[m].lCL = mean - numsds * sd;
                stages[m].sd = sd;

                if (m > 0) {
                    meanLine.push({ x1: stages[m - 1].stageDividerX, y1: stages[m - 1].mean, x2: stages[m].stageDividerX, y2: stages[m].mean, });
                    uclLines.push({ x1: stages[m - 1].stageDividerX, y1: stages[m].uCL, x2: stages[m].stageDividerX, y2: stages[m].uCL });
                    lclLines.push({ x1: stages[m - 1].stageDividerX, y1: stages[m].lCL, x2: stages[m].stageDividerX, y2: stages[m].lCL });
                    stageDividers.push({ x1: stages[m].stageDividerX, prevDividerX: stages[m - 1].stageDividerX, stageName: stages[m].stage });
                }
                else {
                    meanLine.push({ x1: stages[m].startX, y1: stages[m].mean, x2: stages[m].stageDividerX, y2: stages[m].mean, });
                    uclLines.push({ x1: stages[m].startX, y1: stages[m].uCL, x2: stages[m].stageDividerX, y2: stages[m].uCL });
                    lclLines.push({ x1: stages[m].startX, y1: stages[m].lCL, x2: stages[m].stageDividerX, y2: stages[m].lCL });
                    stageDividers.push({ x1: stages[m].stageDividerX, prevDividerX: stages[m].startX, stageName: stages[m].stage });
                }
            }

            /*for (let j = 0; j < stages.length; j++) {
                if (j > 0) {
                    meanLine.push({ x1: stages[j - 1].stageDividerX, y1: stages[j - 1].mean, x2: stages[j].stageDividerX, y2: stages[j].mean, });
                    uclLines.push({ x1: stages[j - 1].stageDividerX, y1: stages[j].uCL, x2: stages[j].stageDividerX, y2: stages[j].uCL });
                    lclLines.push({ x1: stages[j - 1].stageDividerX, y1: stages[j].lCL, x2: stages[j].stageDividerX, y2: stages[j].lCL });
                    stageDividers.push({ x1: stages[j].stageDividerX, prevDividerX: stages[j - 1].stageDividerX, stageName: stages[j].stage });
                }
                else {
                    meanLine.push({ x1: stages[j].startX, y1: stages[j].mean, x2: stages[j].stageDividerX, y2: stages[j].mean, });
                    uclLines.push({ x1: stages[j].startX, y1: stages[j].uCL, x2: stages[j].stageDividerX, y2: stages[j].uCL });
                    lclLines.push({ x1: stages[j].startX, y1: stages[j].lCL, x2: stages[j].stageDividerX, y2: stages[j].lCL });
                    stageDividers.push({ x1: stages[j].stageDividerX, prevDividerX: stages[j].startX, stageName: stages[j].stage });
                }
            }*/
            this.chartStages = stages;
            this.meanLine = meanLine;
            this.lclLines = lclLines;
            this.uclLines = uclLines;
            this.stageDividers = stageDividers;
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
            this.svgGroupMain.selectAll("dividerText")
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
        }

        private PlotControlLimits() {
            let viewModel: ControlChartViewModel = this.controlChartViewModel;
            let limitLine = viewModel.limitLine;
            var lclLines = this.lclLines;
            var uclLines = this.uclLines;
            var xScale = this.xScale;
            var yScale = this.yScale;
            this.svgGroupMain.selectAll("uclLine")
                .data(uclLines)
                .enter().append("polyline")
                .attr("points", function (d) { return xScale(d['x1']).toString() + "," + yScale(d['y1']).toString() + "," + xScale(d['x2']).toString() + "," + yScale(d['y2']).toString() })
                .style({ "stroke": limitLine.lineColor, "stroke-width": 1.5, "stroke-dasharray": (limitLine.lineStyle) });
            this.svgGroupMain.selectAll("lclLine")
                .data(lclLines)
                .enter().append("polyline")
                .attr("points", function (d) { return xScale(d['x1']).toString() + "," + yScale(d['y1']).toString() + "," + xScale(d['x2']).toString() + "," + yScale(d['y2']).toString() })
                .style({ "stroke": limitLine.lineColor, "stroke-width": 1.5, "stroke-dasharray": (limitLine.lineStyle) });
            var yformatValue = d3.format(viewModel.yAxis.AxisFormat);
            this.svgGroupMain.selectAll("uclText")
                .data(uclLines)
                .enter().append("text")
                .attr("x", function (d) { return xScale(d['x1']).toString() })
                .attr("y", function (d) { return yScale(d['y1']).toString() })
                .attr("dx", ".35em")
                .attr("dy", "-.25em")
                .attr("text-anchor", "start")
                .text(function (d) { return 'UCL = ' + yformatValue(d['y1']).toString() })
                .style("font-size", limitLine.textSize + 'px')
                .style("fill", limitLine.textColor);
            this.svgGroupMain.selectAll("lclText")
                .data(lclLines)
                .enter().append("text")
                .attr("x", function (d) { return xScale(d['x1']).toString() })
                .attr("y", function (d) { return yScale(d['y1']).toString() })
                .attr("dx", ".35em")
                .attr("dy", "-.25em")
                .attr("text-anchor", "start")
                .text(function (d) { return 'LCL = ' + yformatValue(d['y1']).toString() })
                .style("font-size", limitLine.textSize + 'px')
                .style("fill", limitLine.textColor);
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
                            standardDeviation: viewModel.standardDeviation,
                            useSDSubGroups: viewModel.useSDSubGroups
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
                case 'crossHairs':
                    var config: VisualObjectInstance = {
                        objectName: objectName,
                        selector: null,
                        properties: {
                            crossHairLineColor: viewModel.crossHairLine.lineColor,
                            crossHairLabelSize: viewModel.crossHairLine.textSize,
                            showCrossHairs: viewModel.crossHairLine.show
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