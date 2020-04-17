/*
 *  Power BI Visual CLI
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
    'use strict';
    import createLegend = powerbi.extensibility.utils.chart.legend.createLegend;
    import positionChartArea = powerbi.extensibility.utils.chart.legend.positionChartArea;
    import LegendData = powerbi.extensibility.utils.chart.legend.LegendData;
    import ILegend = powerbi.extensibility.utils.chart.legend.ILegend;
    import IColorPalette = powerbi.extensibility.IColorPalette;
    import ISelectionId = powerbi.visuals.ISelectionId;
    import TextProperties = powerbi.extensibility.utils.formatting.TextProperties;
    import textMeasurementService = powerbi.extensibility.utils.formatting.textMeasurementService;
    import ITooltipServiceWrapper = powerbi.extensibility.utils.tooltip.ITooltipServiceWrapper;
    import TooltipEnabledDataPoint = powerbi.extensibility.utils.tooltip.TooltipEnabledDataPoint;
    import createTooltipServiceWrapper = powerbi.extensibility.utils.tooltip.createTooltipServiceWrapper;

    let legendData: LegendData;
    let distinctGroup: PrimitiveValue[] = [];
    let uniqueGroupAndSelection: IUniqueGroupAndSelection[] = [];
    let dataPoint: IVisualDataPoint;
    let visual: d3.Selection<any>;
    let rootContainer: d3.Selection<any>
    let containerHeight: number;
    let containerWidth: number;
    let lineFunction: any;
    let margin: { top: number, bottom: number, sides: number };
    let cardDistance: number;
    let cardWidth: number;
    let cardHeight: number;
    let spaceLiteral: string = ' ';
    let rootNodeCount: number;
    let legendHeight: number;
    let root: any;
    let self;
    let lineType: string;
    let tree: d3.layout.Tree<d3.layout.tree.Node>;
    let selectionManager: ISelectionManager;
    let divisionVal: number = 300;
    let maxDivisionVal: number = 220;
    let mulVal: number = 40;
    let twoVal: number = 2;
    let tenthVal: number = 10;
    let oneNegativeVal: number = -1
    const nodes: INodeHash = {};

    interface IVisualDataPoint extends d3.layout.tree.Node,
        TooltipEnabledDataPoint {
        x: number;
        y: number;
        x0: number;
        y0: number;
        selfId: any;
        parentId: any;
        children: IVisualDataPoint[];
        mainLabel: any;
        subLabel: any;
        group: any;
        level: number;
        selectionId: any;
        selectionId1: ISelectionId;
        color: string;
        tooltipInfo?: VisualTooltipDataItem[];
        imageData: any;
    }

    interface IUniqueGroupAndSelection {
        groupName: PrimitiveValue,
        index: number,
        selectionId: ISelectionId,
        color: string
    }

    interface IVisualViewModel {
        nodeList: INodeHash;
        levelList: ILevelListHash;
        height: number;
    }

    interface INodeHash {
        [details: string]: IVisualDataPoint;
    }

    interface ILevelListHash {
        [details: string]: INodeHash;
    }

    /**
     * Validates base64 image and images ending with jpg/gif/png/jpeg
     * 
     * @param val
     */
    function imageValidator(val) {
        let imageValidator: RegExp;
        let anotherImageValidator: RegExp;
        let httpImageValidator: RegExp;
        anotherImageValidator = new RegExp("data:image\/([a-zA-Z]*)[+]([a-zA-Z]*);base64,([^\"]*)");
        imageValidator = new RegExp("data:image\/([a-zA-Z]*);base64,([^\"]*)");
        httpImageValidator = new RegExp("(http(s?):)([/|.|\w|\s|-])*\.(?:jpg|gif|png|jpeg|)");
        let imgValidate: any;
        let anotherImgValidate: any;
        let httpImgValidate: any;
        imgValidate = imageValidator.test(String(val));
        anotherImgValidate = anotherImageValidator.test(String(val));
        httpImgValidate = httpImageValidator.test(val);
        if (!imgValidate && !anotherImgValidate && !httpImgValidate) {
            displayErrorMessage(errorText);
            this.isImageDataPresent = true;
            return false;
        }
    }

    /**
     * Get all selection ID's
     * 
     * @param dataView 
     * @param host 
     */
    function getSelectionIds(dataView: DataView, host: IVisualHost): ISelectionId[] {
        return dataView.table.identity.map((identity: any) => {
            const categoryColumn: DataViewCategoryColumn = {
                source: dataView.table.columns[zeroVal],
                values: null,
                identity: [identity]
            };
            return host.createSelectionIdBuilder()
                .withCategory(categoryColumn, zeroVal)
                .createSelectionId();
        });
    }

    /**
     * Get Selection Id for card colour
     * 
     * @param groupName 
     */
    function getSelectionIDForCardColor(groupName: string): any {
        return uniqueGroupAndSelection.filter(item => { if (item.groupName == groupName) return item.selectionId; });
    }

    /**
     * This function searches for the required parent node in the list & pushes node as it's child
     * 
     * @param list 
     * @param node 
     */
    function searchAndPush(list: IVisualDataPoint[], node: IVisualDataPoint): boolean {
        if (!list || Object.keys(list).length === zeroVal) {
            return false;
        } else {
            for (const key of list) {
                if (key.selfId.toString() === node.parentId.toString()) {
                    key.children.push(node);
                    return true;
                } else {
                    const retBool: boolean = searchAndPush(key.children, node);
                    if (retBool) {
                        return true;
                    }
                }
            }
            return false;
        }
    }

    /**
     * This function assigns level to tree nodes and returns max level value
     * 
     * @param node 
     * @param parentLevel 
     * @param lvlList 
     */
    function assignLevel(node: IVisualDataPoint, parentLevel: number, lvlList: ILevelListHash): number {
        if (!node) {
            return parentLevel;
        }
        node.level = parentLevel + 1;
        if (Object.keys(lvlList).length === zeroVal || !(lvlList[node.level])) {
            lvlList[node.level] = {};
        }
        lvlList[node.level][node.selfId] = node;
        let childLevel: number = zeroVal;
        for (const key of node.children) {
            if (key) {
                childLevel = assignLevel(key, node.level, lvlList);
            }
        }
        return childLevel + node.level;
    }

    /**
     * Function to check if two arrays have same elements or not
     * 
     * @param prevArr 
     * @param currArr 
     */
    function sameArrays(prevArr: string[], currArr: string[]): boolean {
        const containsAll = (prevArr, currArr) => currArr.every(currArrItem => prevArr.includes(currArrItem));
        const sameMembers = (prevArr, currArr) => containsAll(prevArr, currArr) && containsAll(currArr, prevArr);
        return sameMembers(prevArr, currArr);
    }

    /**
     * 
     * 
     * @param d 
     */
    function toggle(d: any): boolean {
        if (d.children) {
            d._children = d.children;
            d.children = null;
            return false;
        } else {
            if (d._children) {
                d.children = d._children;
                d._children = null;
                return true;
            } else {
                return false;
            }
        }
    }

    /**
     * 
     * 
     * @param d 
     */
    function toggleAll(d: any): void {
        if (d.children) {
            d.children.forEach(toggleAll);
            toggle(d);
        }
    }

    /**
     * Method to insert straight lines for links
     * @param link 
     * @param linkSettings 
     * @param duration 
     */
    function methodForStraightLines(link, linkSettings, duration) {
        if (lineType === classic) { // straight lines
            link.enter().insert('path', 'g').attr('class', linkClass)
            .style({ 
                stroke: `${linkSettings.color}`, 
                'stroke-width': `${linkSettings.width}px`, 
                fill: 'none' 
            }).attr('d', (data: any): any => {
                const uLine: any = ((d1: any): any => { 
                    return [
                        { x: d1.source.x, y: d1.source.y }, 
                        { x: d1.source.x, y: d1.source.y }, 
                        { x: d1.source.x, y: d1.source.y }, 
                        { x: d1.source.x, y: d1.source.y }]; 
                    })(data);
                return lineFunction(uLine);
            });
            // Transition links to their new position.
            link.transition().duration(duration).attr('d', (data: any): any => {
                const uLine: any = ((d1: any): any => { 
                    return [
                        { x: d1.source.x, y: d1.source.y + cardHeight },
                        { x: d1.source.x, y: d1.target.y - cardHeight / twoVal },
                        { x: d1.target.x, y: d1.target.y - cardHeight / twoVal },
                        { x: d1.target.x, y: d1.target.y }]; 
                    })(data);
                return lineFunction(uLine);
            });
            //Transition exiting nodes to the parent's new position.
            link.exit().transition().duration(duration).attr('d', (data: any): any => {
                // This is needed to draw the lines right back to the caller
                const uLine: any = ((d1: any): any => { 
                    return [
                        { x: d1.source.x, y: d1.source.y },
                        { x: d1.source.x, y: d1.source.y },
                        { x: d1.source.x, y: d1.source.y },
                        { x: d1.source.x, y: d1.source.y }];
                    })(data);
                return lineFunction(uLine);
            }).each('end', (): void => { 
                self.visualRootCalled = null; 
            });
            self.root.on('click', () => { 
                self.visualRootCalled = root; 
                updateTree(root); 
                (<Event>d3.event).stopPropagation(); 
            });
        }
    }

    /**
     * Method to update sub labels for cards
     * 
     * @param nodeEnter 
     * @param labelFontSize 
     * @param visualLabelSettings 
     * @param cardSettings 
     * @param visualSubLabelSettings 
     * @param subLabelFontSize 
     * @param nodeImages 
     */
    function methodToUpdateSubLabels(nodeEnter, labelFontSize, visualLabelSettings, cardSettings, visualSubLabelSettings, subLabelFontSize, nodeImages) {
        nodeEnter.append('text').attr('id', (data: any, iterator: number): string => {
            return `label-${iterator}`; })
            .style({ 
                'font-size': `${labelFontSize}px`, 
                'font-family': `${visualLabelSettings.fontFamily}`, 
                fill: `${visualLabelSettings.textFill}` 
            }).attr('transform', (data: any, iterator: number): string => {
                const labelProp: TextProperties = { 
                    text: data.mainLabel,
                    fontFamily: visualLabelSettings.fontFamily,
                    fontSize: `${labelFontSize}px` 
                }, labelWidth: number = textMeasurementService.measureSvgTextWidth(labelProp),
                labelHeight: number = textMeasurementService.measureSvgTextHeight(labelProp) - 10;
                const subLabelProp: TextProperties = {
                    text: data.subLabel, 
                    fontFamily: visualSubLabelSettings.fontFamily,
                    fontSize: `${subLabelFontSize}px` 
                }, subLabelHeight: number = textMeasurementService.measureSvgTextHeight(subLabelProp),
                xStart: number = (cardSettings.dimension > labelWidth) ? (cardSettings.dimension - labelWidth) / 2 : 5;
                let yStart: number; 
                if (data.subLabel === '') {
                    yStart = cardSettings.dimension - subLabelHeight - 5;
                } else if (data.imageData === '') {
                    yStart = ((cardSettings.dimension + (labelHeight / 2)) / 2) - subLabelHeight;
                } else {
                    yStart = cardSettings.dimension - subLabelHeight
                }
                nodeImages.style('height', `${cardSettings.dimension - labelHeight - subLabelHeight}px`); 
                return `translate(${xStart}, ${yStart})`;
            }).text((data: any): any => { 
                const labelProp: TextProperties = {
                    text: data.mainLabel,
                    fontFamily: visualLabelSettings.fontFamily,
                    fontSize: `${labelFontSize}px`
                }; 
                return data.mainLabel ? textMeasurementService.getTailoredTextOrDefault(labelProp, cardSettings.dimension - 5) : '';
            });
        // Append Sublabel
        nodeEnter.append('text').attr('id', (data: any, iterator: number): string => { 
                return `sublabel-${iterator}`; 
            }).style({ 
                'font-size': `${subLabelFontSize}px`, 
                'font-family': `${visualSubLabelSettings.fontFamily}`, 
                fill: `${visualSubLabelSettings.textFill}` 
            }).attr('transform', (data: any): string => {
                const labelProp: TextProperties = { 
                    text: data.mainLabel, fontFamily: visualLabelSettings.fontFamily, 
                    fontSize: `${labelFontSize}px` 
                }, labelWidth: number = textMeasurementService.measureSvgTextWidth(labelProp),
                labelHeight: number = textMeasurementService.measureSvgTextHeight(labelProp);
                const subLabelProp: TextProperties = { 
                    text: data.subLabel, fontFamily: visualSubLabelSettings.fontFamily, 
                    fontSize: `${subLabelFontSize}px` 
                }, subLabelWidth: number = textMeasurementService.measureSvgTextWidth(subLabelProp),
                subLabelHeight: number = textMeasurementService.measureSvgTextHeight(subLabelProp);
                const xStart: number = (cardSettings.dimension > subLabelWidth) ? (cardSettings.dimension - subLabelWidth + 5) / 2 : 5; 
                let yStart: number;
                if (data.imageData === '') { 
                    yStart = (cardSettings.dimension + (subLabelHeight)) / 2; 
                    return `translate(${xStart}, ${yStart})`; 
                }
                else { 
                    yStart = cardSettings.dimension; 
                    return `translate(${xStart}, ${yStart - 5})`; 
                }
            }).text((data: any): any => {
                const subLabelProp: TextProperties = { 
                    text: data.subLabel, 
                    fontFamily: visualSubLabelSettings.fontFamily, 
                    fontSize: `${subLabelFontSize}px` 
                };
                return data.subLabel ? textMeasurementService.getTailoredTextOrDefault(subLabelProp, cardSettings.dimension - 5) : '';
            });
    }

    /**
     * 
     * 
     * @param node 
     * @param pathEle 
     * @param source 
     * @param rootDivSelection 
     */
    function methodToEnterNewNodesAtParentsPreviousPosition(node, pathEle, source, rootDivSelection) : d3.Selection<any> {
        return node.enter().append('g').attr('class', nodeClass)
        .attr('id', (data: any, iterator: number): string => {
            return `node-${iterator}`;
        }).attr('transform', (card: any): any => { 
            return `translate(${card.parent ? (card.parent.x - (self.settings.card.dimension / twoVal)) 
                : (card.x0 - (self.settings.card.dimension / twoVal))}, ${card.parent ? card.parent.y : card.y0})`; 
        }).on('click', (card: any): void => {
            selectionManager.select(card.selectionId).then((ids: ISelectionId[]) => { 
                node.attr('fill-opacity', minOpacity); 
                d3.select(event.currentTarget).attr('fill-opacity', maxOpacity); pathEle.style('opacity', minOpacity); 
            });
            selectionManager[`selectedIds`].push(card.selectionId); (<Event>d3.event).stopPropagation();
        }).on('dblclick', (card: any): void => {
            let d3Event: any = d3.event;
            if (toggle(card)) { 
                d3Event.ctrlKey ? self.visualRootCalled = card : self.visualRootCalled = source; 
                updateTree(self.visualRootCalled); 
            }
            else { 
                d3Event.ctrlKey ? (self.visualRootCalled = card.parent ? card.parent : source) : self.visualRootCalled = source; 
                updateTree(self.visualRootCalled); 
            }
            node.attr('fill-opacity', maxOpacity);
            pathEle.style('opacity', maxOpacity);
        }).on('mouseover', (data: any): void => { 
            const x: number = d3.mouse(rootContainer.node())[0], y: number = d3.mouse(rootContainer.node())[1];
            self.host.tooltipService.show({ 
                dataItems: data.tooltipInfo, 
                coordinates: [x, y], 
                isTouchEvent: false, 
                identities: data.selectionID 
            });
        }).on('mousemove', (data: any): void => { 
            const x: number = d3.mouse(rootContainer.node())[0], y: number = d3.mouse(rootContainer.node())[1]; 
            self.host.tooltipService.move({ dataItems: data.tooltipInfo, coordinates: [x, y], isTouchEvent: false, identities: data.selectionID });
        }).on('mouseout', (): void => { 
            self.host.tooltipService.hide({ immediately: true, isTouchEvent: false }); 
        });
    }

    /**
     * 
     * 
     * @param cardSettings 
     */
    function methodForDiagonalEnterPoints(cardSettings) {
        return d3.svg.diagonal().source((data: any): any => {
            return { 
                x: data.source.x0 ? data.source.x0 : data.source.x, y: data.source.y0 ? data.source.y0 + cardSettings.dimension : data.source.y + cardSettings.dimension 
            };
        }).target((data: any): any => { 
            return { 
                x: data.target.x, y: data.target.y 
            }; 
        });
    }

    /**
     * 
     * 
     * @param cardSettings 
     */
    function methodToReturnPathValuesToDrawDiagonalLines(cardSettings) {
        return d3.svg.diagonal().source((data: any): any => { 
            return { 
                x: data.source.x, y: data.source.y + cardSettings.dimension 
            }; 
        }).target((data: any): any => { 
            return { 
                x: data.target.x, y: data.target.y 
            }; 
        });
    }

    /**
     * Method to add diagonal lines for links
     * 
     * @param link 
     * @param linkSettings 
     * @param getDiagonalEnterPoints 
     * @param duration 
     * @param getDiagonalPoints 
     */
    function methodForDiagonalLines(link, linkSettings, getDiagonalEnterPoints, duration, getDiagonalPoints) {
        link.enter().insert('path', 'g').attr('class', linkClass).style({
            stroke: `${linkSettings.color}`, 'stroke-width': `${linkSettings.width}px`, fill: 'none'
        }).attr('d', (data: any): any => { 
            return getDiagonalEnterPoints(data); 
        }).transition().duration(duration).attr('d', getDiagonalEnterPoints);
        // Transition links to their new position.
        link.transition().duration(duration).attr('d', (data: any): any => { 
            return getDiagonalPoints(data); 
        });
        // Transition exiting nodes to the parent's new position.
        link.exit().transition().duration(duration).attr('d', (data: any): any => { 
            d3.svg.diagonal().source({ x: data.source.x, y: data.source.y })
            .target({ x: data.source.x, y: data.source.y }); 
        }).remove();
    }

    /**
     * Method to update links for all nodes
     * 
     * @param nodes 
     */
    function methodToUpdateLinks(nodes) {
        return visual.selectAll('path.link').data(tree.links(nodes), 
            (data: any): any => { 
                return data.target.selfId; 
            });
    }

    /**
     * Method to update the nodes and the organization tree
     * 
     * @param source 
     */
    function updateTree(source: any): void {
        const duration: number = 1;
        let pathEle: any = d3.selectAll(path), rootDivSelection: any = d3.select('.rootDiv'), currentTargetEle: any = d3.select(event.currentTarget);
        // Compute the new tree layout.
        const nodes: d3.layout.tree.Node[] = tree.nodes(source).reverse();
        // Normalize for fixed-depth.
        nodes.forEach((d: d3.layout.tree.Node): void => { 
            d.y = (d.depth * cardDistance); 
        });
        // Update the node
        const node: d3.selection.Update<any> = visual.selectAll('g.node')
            .data(nodes, (data: any): any => { 
                return data.selfId; 
            });
        // Enter any new nodes at the parent's previous position.
        const nodeEnter : d3.Selection<any> = methodToEnterNewNodesAtParentsPreviousPosition(node, pathEle, source, rootDivSelection);
        rootDivSelection.on('click', (): void => { 
            if (selectionManager[`selectedIds`].length) { 
                selectionManager.clear(); node.attr('fill-opacity', maxOpacity);
                pathEle.style('opacity', maxOpacity); 
            } 
        });
        let cardSettings: any = self.settings.card;
        nodeEnter.append('rect').classed('container', true).attr({
                    width: `${cardSettings.dimension}px`, 
                    height: `${cardSettings.dimension}px`, 
                    rx: `${cardSettings.cornerRadius}px`, 
                    ry: `5px` 
                }).style('fill', (card: any): any => { 
                    return `${card.color !== null ? card.color : cardSettings.backgroundFill}`;
                });
        if (cardSettings.borderShow) { 
            nodeEnter.select('rect').style({ 
                stroke: `${cardSettings.borderColor}`, 'stroke-width': `${cardSettings.borderSize}`
            }); 
        }
        // render images according to base64 data provided
        let nodeImages: any = nodeEnter.append('image')
            .attr('id', (data: any, iterator: number): string => { 
                return `image-${iterator}`; 
            }).attr('xlink:href', (data: any): any => { 
                return `${data.imageData}`; 
            }).attr('width', (data: any): any => { 
                const image: HTMLImageElement = new Image(); 
                image.src = data.imageData; 
                return `${cardSettings.dimension}px`; //this return will set width of box
            }).attr('height', (data: any): any => { 
                const image: HTMLImageElement = new Image(); 
                image.src = data.imageData; 
                return `${cardSettings.dimension - (heightPercent * cardSettings.dimension)}px`; 
            });
        let labelFontSize: number = self.settings.labelSettings.fontSize, visualSubLabelSettings: any = self.settings.subLabelSettings;
        let visualLabelSettings: any = self.settings.labelSettings;
        if (labelFontSize > maxFontSize) { 
            visualLabelSettings.fontSize = maxFontSize; 
            labelFontSize = maxFontSize; 
        } else if (labelFontSize < minFontSize) { 
            visualLabelSettings.fontSize = minFontSize; 
            labelFontSize = minFontSize; 
        }
        let subLabelFontSize: number = visualSubLabelSettings.fontSize;
        if (subLabelFontSize > maxFontSize) {
            visualSubLabelSettings.fontSize = maxFontSize;
            subLabelFontSize = maxFontSize;
        } else if (subLabelFontSize < minFontSize) {
            visualSubLabelSettings.fontSize = minFontSize;
            subLabelFontSize = minFontSize;
        }
        methodToUpdateSubLabels(nodeEnter, labelFontSize, visualLabelSettings, cardSettings, visualSubLabelSettings, subLabelFontSize, nodeImages);
        // Transition nodes to their new position.
        const nodeUpdate: any = node.transition().duration(duration).attr('transform', (data: any): any => { 
            return `translate(${data.x - (cardSettings.dimension / twoVal)}, ${data.y})`; 
        });
        // Transition exiting nodes to the parent's new position.
        const nodeExit: any = node.exit().transition().duration(duration).attr('transform', (data: any): any => { 
            return `translate(${data.parent ? data.parent.x - (cardSettings.dimension / twoVal) : data.x0}, 
            ${data.parent ? data.parent.y : data.y0})`; 
        }).remove();
        nodeExit.select('text').style('fill-opacity', 1e-6);
        // Update the links…
        const link: d3.selection.Update<any> = methodToUpdateLinks(nodes);
        const getDiagonalEnterPoints: any = methodForDiagonalEnterPoints(cardSettings);
        // function returns path values to draw diagonal when entered & transitioned
        const getDiagonalPoints: any = methodToReturnPathValuesToDrawDiagonalLines(cardSettings);
        let linkSettings: any = self.settings.links;
        methodForStraightLines(link, linkSettings, duration);
        if (lineType === modern) { // diagonal lines
            // Enter any new links at the parent's previous position.
            methodForDiagonalLines(link, linkSettings, getDiagonalEnterPoints, duration, getDiagonalPoints);
            // Stash the old positions for transition.
            nodes.forEach((dataPoint: IVisualDataPoint): any => { 
                dataPoint.x0 = dataPoint.x; dataPoint.y0 = dataPoint.y; 
            });
            self.root.on('click', () => { 
                self.visualRootCalled = root; 
                updateTree(root); 
                (<Event>d3.event).stopPropagation(); 
            });
        }
    }

    /**
     * Method to display error message
     * 
     * @param message       - error message to be displayed 
     */
    function displayErrorMessage(message) {
        let errorDiv: any = d3.select('.rootDiv').append(div).classed(errorMessage, true);
        errorDiv.text(message);
    }

    /**
     * This function formats input data to usable data
     * 
     * @param totColumns 
     */
    function mandatoryDataPresent(totColumns: any): boolean {
        let columnsCount: number = totColumns.length;
        let isLabel: boolean = false;
        let isIDPresent: boolean = false;
        let isParentIdPresent: boolean = false;
        for (let iCounter: number = zeroVal; iCounter < columnsCount; iCounter++) {
            if (totColumns[iCounter].roles.mainLabel) {
                isLabel = true;
            }
            if (totColumns[iCounter].roles.id) {
                isIDPresent = true;
            }
            if (totColumns[iCounter].roles.parentId) {
                isParentIdPresent = true;
            }
        }
        if (isLabel && isIDPresent && isParentIdPresent) {
            return true;
        }
        return false;
    }

    /**
     * Method to get list of unique groups
     * 
     * @param dataViews 
     * @param columnsCount 
     * @param allGroupValuesCount 
     * @param groupPosition 
     * @param thisObj 
     * @param indexArray 
     */
    function getUniqueListOfGroups(dataViews, columnsCount, allGroupValuesCount, groupPosition, thisObj, indexArray) {
        for (let iCounter: number = zeroVal; iCounter < columnsCount; iCounter++) {
            if (dataViews[zeroVal].table.columns[iCounter].roles.group === true) {
                groupPosition = iCounter;
                thisObj.isGroupPresent = true;
                for (let jCounter: number = zeroVal; jCounter < allGroupValuesCount; jCounter++) {
                    if (distinctGroup.indexOf(dataViews[zeroVal].table.rows[jCounter][iCounter].toString()) === negOneVal) {
                        distinctGroup.push(dataViews[zeroVal].table.rows[jCounter][iCounter]);
                        indexArray.push(jCounter);
                    }
                }
            }
        }
    }

    /**
     * Method to get selection ids for groups
     * 
     * @param selectionIDListForGroups 
     * @param indexArray 
     * @param selectionIDList 
     * @param colorPalette 
     * @param temp 
     * @param host 
     */
    function getSelectionIdsForDistinctGroups(selectionIDListForGroups, indexArray, selectionIDList, colorPalette, temp, host) {
        for (let iCounter: number = zeroVal; iCounter < indexArray.length; iCounter++) { 
            selectionIDListForGroups.push(selectionIDList[indexArray[iCounter]]); 
        }
        for (let iCounter: number = zeroVal; iCounter < Math.max(selectionIDListForGroups.length, distinctGroup.length); iCounter++) {
            const defaultColor: Fill = { 
                solid: { 
                    color: colorPalette.getColor(distinctGroup[iCounter] + temp).value 
                } 
            };
            uniqueGroupAndSelection.push({
                groupName: distinctGroup[iCounter], index: indexArray[iCounter], selectionId: host.createSelectionIdBuilder().withMeasure(distinctGroup[iCounter].toString()).createSelectionId(),
                color: colorPalette.getColor(distinctGroup[iCounter] + temp).value
            });
        }
    }

    /**
     * Method to update table data
     * 
     * @param table 
     * @param idIndex 
     * @param parentIdIndex 
     * @param mainLabelIndex 
     * @param thisObj 
     * @param subLabelIndex 
     * @param groupIndex 
     * @param imageDataIndex 
     * @param colorPalette 
     * @param temp 
     * @param selectionIDList 
     * @param host 
     * @param cardColour 
     */
    function updateTable(table, idIndex, parentIdIndex, mainLabelIndex, thisObj, subLabelIndex, groupIndex, imageDataIndex, colorPalette, temp, selectionIDList, host, cardColour) {
        table.columns.forEach((column: DataViewMetadataColumn): void => {
            const role: DataViewMetadataColumn['roles'] = column.roles;
            if (role[`id`]) { 
                idIndex = column.index; 
            }
            if (role[`parentId`]) { 
                parentIdIndex = column.index; 
            }
            if (role[`mainLabel`]) { 
                mainLabelIndex = column.index; 
            }
            if (role[`subLabel`]) { 
                thisObj.isSublabelPresent = true; 
                subLabelIndex = column.index; 
            }
            if (role[`group`]) { 
                groupIndex = column.index; 
            }
            if (role[`image`]) { 
                imageDataIndex = column.index; 
            }
        });
        let i: number = zeroVal;
        table.rows.forEach((row: DataViewTableRow): void => {
            const groupName: string = groupIndex !== null ? row[groupIndex].toString() : '';
            //get Default color for each card
            const defaultColor: Fill = { 
                solid: { 
                    color: colorPalette.getColor(groupName + temp).value
                } 
            };
            dataPoint = {
                x: zeroVal, y: zeroVal, x0: zeroVal, y0: zeroVal, selfId: null, parentId: null, children: [], mainLabel: '', subLabel: '', group: '', level: null,
                selectionId: null, selectionId1: null, color: '', tooltipInfo: [], imageData: null
            };
            if ((row[1] === null)) { 
                if ((row[0] < 0) || (row[0] === 0)) { 
                    displayErrorMessage("ID contains inappropriate data"); 
                    return; 
                } 
            }
            dataPoint.selfId = idIndex !== null ? row[idIndex] : '';
            dataPoint.parentId = parentIdIndex !== null ? row[parentIdIndex] : '';
            dataPoint.mainLabel = mainLabelIndex !== null ? row[mainLabelIndex] : '';
            dataPoint.subLabel = subLabelIndex !== null ? row[subLabelIndex] : '';
            dataPoint.group = groupIndex !== null ? row[groupIndex] : '';
            dataPoint.selectionId = selectionIDList[i];
            dataPoint.selectionId1 = getSelectionIDForCardColor(groupIndex !== null ? row[groupIndex].toString() : '');
            dataPoint.color = dataPoint.group.length > zeroVal ? host.colorPalette.getColor(dataPoint.group).value : cardColour;
            dataPoint.imageData = (row[imageDataIndex] !== null && row[imageDataIndex] !== undefined) ? row[imageDataIndex] : '';
            if (idIndex !== null) {
                dataPoint.tooltipInfo.push(<VisualTooltipDataItem>{ 
                    displayName: table.columns[idIndex].displayName, value: dataPoint.selfId ? dataPoint.selfId.toString() : '' 
                });
            }
            if (parentIdIndex !== null) {
                dataPoint.tooltipInfo.push(<VisualTooltipDataItem>{ 
                    displayName: table.columns[parentIdIndex].displayName, value: dataPoint.parentId ? dataPoint.parentId.toString() : '' 
                });
            }
            if (mainLabelIndex !== null) {
                dataPoint.tooltipInfo.push(<VisualTooltipDataItem>{ 
                    displayName: table.columns[mainLabelIndex].displayName, value: dataPoint.mainLabel ? dataPoint.mainLabel.toString() : '' 
                });
            }
            if (subLabelIndex !== null) {
                dataPoint.tooltipInfo.push(<VisualTooltipDataItem>{ 
                    displayName: table.columns[subLabelIndex].displayName, value: dataPoint.subLabel ? dataPoint.subLabel.toString() : '' 
                });
            }
            if (groupIndex !== null) {
                dataPoint.tooltipInfo.push(<VisualTooltipDataItem>{ 
                    displayName: table.columns[groupIndex].displayName, value: dataPoint.group ? dataPoint.group.toString() : '' 
                });
            }
            nodes[dataPoint.selfId] = dataPoint;
            ++i;
        });
        return table;
    }

    /**
     * 
     * 
     * @param options 
     * @param host 
     * @param thisObj 
     * @param cardColour 
     */
    function visualTransform(options: VisualUpdateOptions, host: IVisualHost, thisObj: any, cardColour: string): IVisualViewModel {
        let table: DataViewTable = options.dataViews[zeroVal].table;
        if (table.rows === null || table.rows.length === zeroVal) { 
            return null; 
        }
        //currently this level list is not used anywhere in the visual, can be removed if no further use found
        const levelList: ILevelListHash = {};
        // list of roots in data
        const rootNodes: string[] = [];
        // getting index values to retrieve values from row
        let idIndex: number = null, mainLabelIndex: number = null, subLabelIndex: number = null, groupIndex: number = null, parentIdIndex: number = null;
        let imageDataIndex: number = null, colorPalette: IColorPalette = host.colorPalette;
        //get distinct list of groups
        let dataViews: DataView[] = options.dataViews;
        const columnsCount: number = dataViews[zeroVal].table.columns.length, allGroupValuesCount: number = dataViews[zeroVal].table.rows.length;
        distinctGroup = []; 
        const indexArray: number[] = []; 
        let groupPosition: number, selectionIDList: ISelectionId[];
        const selectionIDListForGroups: ISelectionId[] = [], temp: string = '';
        // Get Unique list of groups
        getUniqueListOfGroups(dataViews, columnsCount, allGroupValuesCount, groupPosition, thisObj, indexArray);
        // Get all selection ID's
        selectionIDList = getSelectionIds(dataViews[zeroVal], host);
        //Get Selection ID's for distinct Groups
        getSelectionIdsForDistinctGroups(selectionIDListForGroups, indexArray, selectionIDList, colorPalette, temp, host);
        table = updateTable(table, idIndex, parentIdIndex, mainLabelIndex, thisObj, subLabelIndex, groupIndex, imageDataIndex, colorPalette, temp, selectionIDList, host, cardColour);
        Object.keys(nodes).forEach((key: string): void => {
            if (nodes[key].parentId) {
                if (nodes[nodes[key].parentId]) { 
                    nodes[nodes[key].parentId].children.push(nodes[key]); delete nodes[key]; 
                }
                else {
                    Object.keys(nodes).forEach((k: string): void => { 
                        if (searchAndPush(nodes[k].children, nodes[key])) { return; } 
                    });
                    delete nodes[key];
                }
            } else { 
                rootNodes.push(key); 
            }
        });
        let treeHeight: number = 0;
        for (const key of rootNodes) { 
            if (rootNodes[key]) { 
                assignLevel(nodes[rootNodes[key]], negOneVal, levelList); 
            } 
        }
        treeHeight = Object.keys(levelList).length;
        return { 
            nodeList: nodes, levelList: levelList, height: treeHeight 
        };
    }

    export class Visual implements IVisual {
        private host: IVisualHost;
        private settings: VisualSettings;
        private viewModel: IVisualViewModel;
        private root: d3.Selection<SVGElement>;
        private selectionManager: ISelectionManager;
        private viewport: IViewport;
        private svg: d3.Selection<HTMLElement>;
        private legend: ILegend;
        // will store actual root of the data provided
        private visualRoot: IVisualDataPoint = null;
        private cardDataPoints: IVisualDataPoint[];
        // will store the node which is called as root
        private visualRootCalled: IVisualDataPoint = null;
        private colorPalette: IColorPalette;
        private previousRoles: string[];
        private minXView: number;
        private width: number;
        private height: number;
        private minYView: number;
        private rootContainer: d3.Selection<any>;
        private isSublabelPresent: boolean;
        private isGroupPresent: boolean;
        private isMandatoryDataPresent: boolean;
        private isImageDataPresent: boolean;
        private events: IVisualEventService;
        private tooltipServiceWrapper: ITooltipServiceWrapper;

        constructor(options: VisualConstructorOptions) {
            this.host = options.host;
            this.root = d3.select(options.element);
            this.selectionManager = options.host.createSelectionManager();
            this.colorPalette = this.host.colorPalette;
            this.previousRoles = [];
            this.legend = createLegend(options.element, false, null, true);
            this.tooltipServiceWrapper = createTooltipServiceWrapper(this.host.tooltipService, options.element);
            this.rootContainer = this.root.append(div).classed(rootDiv, true);
            this.events = options.host.eventService;
        }

        /**
         * Method to change legend postion on change in formatting pane
         * 
         * @param legendOrientation             - contains the legend orientation
         * @param legendPosition                - contains the legend position
         */
        public legendPositionUpdate(
            legendOrientation: string,
            legendPosition: number): number {
            switch (legendOrientation) {
                case 'Top':
                    legendPosition = 0;
                    break;
                case 'Bottom':
                    legendPosition = 1;
                    break;
                case 'TopCenter':
                    legendPosition = 5;
                    break;
                case 'BottomCenter':
                    legendPosition = 6;
                    break;
                default: break;
            }
            return legendPosition;
        }

        /**
         * Method to display error message and update viewport width
         * 
         * @param options           - Contains references to the size of the container and the dataView which contains all the data the visual had queried.
         */
        public methodToDisplayErrorMessageAndUpdateViewPortWidth(options): VisualUpdateOptions {
            if (!this.isMandatoryDataPresent) {
                displayErrorMessage(errorImproperData);
                return;
            }
            if (this.width === undefined || this.width === null) {
                this.width = options.viewport.width;
                this.minXView = this.width / twoVal;
            }
            if (this.height === undefined || this.height === null) {
                this.height = options.viewport.height;
                this.minYView = this.height / twoVal;
            }
            return options;
        }

        /**
         * Method to render new flag
         * 
         * @param currentRoles          - roles
         * @param renderNewFlag
         */
        public updateRenderNewFlag(currentRoles, renderNewFlag): boolean {
            if (this.previousRoles.length !== currentRoles.length) {
                renderNewFlag = true;
            }
            else {
                renderNewFlag = sameArrays(this.previousRoles, currentRoles);
            }
            if (renderNewFlag) {
                this.previousRoles = currentRoles;
                self.visualRoot = null;
                self.visualRootCalled = null;
                self.visualRoot = {
                    x: zeroVal, y: zeroVal, x0: zeroVal, y0: zeroVal, selfId: null,
                    parentId: null, children: [], mainLabel: 'Root', subLabel: '', group: '', level: null,
                    selectionId: null, selectionId1: null, color: null, tooltipInfo: [], imageData: null
                };
                // if multi root nodes add one more node as root of those, as d3.layout.tree only supports single root
                if (rootNodeCount > 1) {
                    Object.keys(self.viewModel.nodeList).forEach((key: string): void => {
                        self.visualRoot.children.push(self.viewModel.nodeList[key]);
                    });
                }
                else {
                    Object.keys(self.viewModel.nodeList).forEach((key: string): void => {
                        self.visualRoot = self.viewModel.nodeList[key];
                    });
                }
            }
            return renderNewFlag;
        }

        /**
         * Method to update card dimensions and width
         * 
         * @param maxCardWidth          - maximum card width
         */
        public updateCardDimensionsAndBorderSize(maxCardWidth) {
            if (cardWidth < this.settings.card.dimension) {
                this.settings.card.dimension = this.settings.card.dimension;
                if (maxCardWidth < this.settings.card.dimension) {
                    this.settings.card.dimension = maxCardWidth;
                }
            }
            else {
                this.settings.card.dimension = cardWidth;
            }
            cardDistance = this.settings.card.dimension * twoVal;
            // setting bounds to link width
            if (this.settings.links.width < 1) {
                this.settings.links.width = 1;
            }
            else if (this.settings.links.width > tenthVal) {
                this.settings.links.width = tenthVal;
            }
            if (this.settings.card.cornerRadius < 1) {
                this.settings.card.cornerRadius = 1;
            }
            else if (this.settings.card.cornerRadius > tenthVal) {
                this.settings.card.cornerRadius = tenthVal;
            }
            // setting bounds to border width
            if (this.settings.card.borderSize < 1) {
                this.settings.card.borderSize = 1;
            }
            else if (this.settings.card.borderSize > maxBorderSize) {
                this.settings.card.borderSize = maxBorderSize;
            }
        }

        /**
         * Method to update svg container
         * 
         * @param thisObj           - refernce variable to 'this'
         * @param svgRootContainer  - the main svg container
         */
        public svgContainerUpdate(thisObj, svgRootContainer) {
            svgRootContainer.on('wheel.zoom', (): void => {
                let wheelDelta: number = (<any>window).d3.event.wheelDeltaY || (<any>window).d3.event.deltaY;
                const tempWidth: number = (thisObj.width) + (wheelDelta * -2);
                const tempHeight: number = (thisObj.height) + (wheelDelta * -2);
                if (tempWidth > 0 && tempHeight > 0) {
                    thisObj.minXView = thisObj.minXView + (wheelDelta);
                    thisObj.minYView = thisObj.minYView + (wheelDelta);
                    thisObj.width = tempWidth; thisObj.height = tempHeight;
                    d3.select('.rootSvg')
                    .attr('viewBox', thisObj.minXView + spaceLiteral + thisObj.minYView + spaceLiteral + thisObj.width + spaceLiteral + thisObj.height);
                }
            });
            svgRootContainer.call(d3.behavior.drag().on('drag', (): void => {
                thisObj.minXView += oneNegativeVal * (<any>window).d3.event.dx;
                thisObj.minYView += oneNegativeVal * (<any>window).d3.event.dy;
                svgRootContainer
                    .attr('viewBox', thisObj.minXView + spaceLiteral + thisObj.minYView + spaceLiteral + thisObj.width + spaceLiteral + thisObj.height);
            }));
        }

        /**
         * Helper method to update options
         * 
         * @param options               - Contains references to the size of the container and the dataView which contains all the data the visual had queried.
         * @param tempIndex 
         */
        public updateOptions(options, tempIndex): VisualUpdateOptions {
            options.dataViews[zeroVal].table.columns.forEach((column: DataViewMetadataColumn): void => {
                Object.keys(column.roles).forEach((role: string): void => {
                    if (role.indexOf('group') !== negOneVal) {
                        tempIndex = column.index; self.settings.legend.titleText = self.settings.legend.titleText !== '' 
                            ? self.settings.legend.titleText : options.dataViews[zeroVal].table.columns[tempIndex].displayName;
                    }
                });
            });
            if (this.settings.legend.showTitle !== true) {
                self.settings.legend.titleText = '';
            }
            legendData = {
                dataPoints: [],
                fontSize: this.settings.legend.fontSize,
                title: this.settings.legend.titleText,
                labelColor: this.settings.legend.labelColor
            };
            uniqueGroupAndSelection.forEach((data: any, iterator: any): void => {
                legendData.dataPoints.push({
                    label: data.groupName === null ? '(Blank)' : data.groupName.toString(), color: data.color.toString(), //name of the label and color of the icon
                    icon: powerbi.extensibility.utils.chart.legend.LegendIcon.Circle, selected: false, identity: data.selectionId //type of the legend icon
                });
            });
            return options;
        }

        /**
         * Method to check if image data is present or not
         */
        public checkIfImageDataIsPresent() {
            if (Object.keys(this.viewModel.nodeList).length === 0) {
                displayErrorMessage(inappropriateColumnData);
                return;
            }
            this.isImageDataPresent = false;
            if (this.viewModel.nodeList[Number(Object.keys(this.viewModel.nodeList))].imageData) {
                this.isImageDataPresent = imageValidator(this.viewModel
                    .nodeList[Number(Object.keys(this.viewModel.nodeList))].imageData);
                if (this.isImageDataPresent) {
                    return;
                }
            }
        }

        /**
         * Method to implement scroll arrow for top center and bottom center
         * 
         * @param legendPosition        - contains legend position
         * @param width                 - legend width
         * @param options               - Contains references to the size of the container and the dataView which contains all the data the visual had queried.
         */
        public scrollArrowForTopAndBottomCenter(legendPosition, width, options): any{
            switch (legendPosition) {
                case 5:  
                    if (d3.select(".navArrow")[0][0] !== null) {
                        const xpos: string = d3.select(".navArrow").attr("transform").substring(10,
                            d3.select(".navArrow")
                                .attr("transform").indexOf(","));
                        if (Number(xpos) > width - 10) {
                            this.legend.changeOrientation(0);
                            this.legend.drawLegend(legendData, options.viewport);
                        }
                    }
                    break;
                case 6:
                    if (d3.select(".navArrow")[0][0] !== null) {
                        const xpos: string = d3.select(".navArrow").attr("transform").substring(10,
                            d3.select(".navArrow")
                                .attr("transform").indexOf(","));
                        if (Number(xpos) > width - 10) {
                            this.legend.changeOrientation(1);
                            this.legend.drawLegend(legendData, options.viewport);
                        }
                        }
                    break;
                default:
                    break;
            }
        }

        /**
         * Method to render and display visual
         * 
         * @param options           - Contains references to the size of the container and the dataView which contains all the data the visual had queried.       
         */
        public update(options: VisualUpdateOptions): void {
            this.clearAll();
            try {
                this.events.renderingStarted(options);
                let dataViews: DataView[] = options.dataViews;
                this.isMandatoryDataPresent = mandatoryDataPresent(dataViews[zeroVal].table.columns);
                options = this.methodToDisplayErrorMessageAndUpdateViewPortWidth(options);
                this.settings = Visual.parseSettings(options && options.dataViews && options.dataViews[zeroVal]);
                lineFunction = d3.svg.line().x((d: any): any => { return d.x; }).y((d: any): any => { return d.y; }).interpolate(linear);
                self = this; 
                const cardColour: string = this.settings.card.backgroundFill; 
                lineType = this.settings.links.view; 
                const thisObject: this = this;
                // Register events
                uniqueGroupAndSelection = [];
                this.isGroupPresent = false;
                this.isSublabelPresent = false;
                this.viewModel = visualTransform(options, self.host, thisObject, cardColour);
                this.checkIfImageDataIsPresent();
                this.viewport = options.viewport;
                // flag to determine if tree data needs to be updated according to roleCount i.e new data added
                let renderNewFlag: boolean = true, tempIndex: number;
                options = this.updateOptions(options, tempIndex);
                const legendOrientation: string = self.settings.legend.position;
                let legendPosition: number;
                legendPosition = this.legendPositionUpdate(legendOrientation, legendPosition);
                self.settings.legend.position = legendOrientation;
                this.legend.changeOrientation(legendPosition);
                if (self.settings.legend.show === true) {
                    this.legend.drawLegend(legendData, options.viewport);
                    positionChartArea(this.rootContainer, this.legend);
                }
                legendHeight = this.legend.getMargins().height;
                let legendWidth: number = this.legend.getMargins().width;
                cardWidth = mulVal * Math.min(options.viewport.width, options.viewport.height) / divisionVal;
                cardHeight = mulVal * Math.min(options.viewport.width, options.viewport.height) / divisionVal;
                let maxCardWidth = mulVal * (options.viewport.height) / maxDivisionVal;
                this.updateCardDimensionsAndBorderSize(maxCardWidth);
                margin = { top: zeroVal, bottom: tenthVal, sides: tenthVal };
                containerHeight = this.viewport.height - legendHeight;
                containerWidth = this.viewport.width - (margin.sides);
                this.minXView = oneNegativeVal * (containerWidth / tenthVal);
                this.minYView = oneNegativeVal * (containerHeight / tenthVal);
                rootContainer = self.rootContainer;
                rootContainer.style('margin-top', '0px');
                tree = d3.layout.tree().size([containerWidth, containerHeight]);
                visual = rootContainer.append('svg').classed('rootSvg', true)
                    .attr('width', containerWidth).attr('height', (containerHeight > 0) ? containerHeight : 0)
                    .attr('viewBox', zeroVal + spaceLiteral + zeroVal + spaceLiteral + containerWidth + spaceLiteral + ((containerHeight > 0)
                        ? containerHeight : 0)).attr('preserveAspectRatio', 'xMidYMid meet').style('cursor', 'context-menu');
                if ((legendPosition === 0) || (legendPosition === 5)) {
                    d3.select('.rootSvg').style('margin-top', maxLegendHeight + 'px')
                }
                else {
                    d3.select('.rootSvg').style('margin-bottom', maxLegendHeight + 'px')
                }
                let width: number = options.viewport.width;
                this.scrollArrowForTopAndBottomCenter(legendPosition, width, options);
                visual.on("contextmenu", () => {
                    const mouseEvent: MouseEvent = <MouseEvent>d3.event;
                    const eventTarget: EventTarget = mouseEvent.target;
                    const dataPoint: any = d3.select(eventTarget).datum();
                    if (dataPoint !== undefined) {
                        this.selectionManager.showContextMenu(dataPoint ? dataPoint.selectionId : {},
                            { x: mouseEvent.clientX, y: mouseEvent.clientY });
                        mouseEvent.preventDefault();
                    }
                });
                rootNodeCount = Object.keys(self.viewModel.nodeList).length;
                // variable to hold number of roles added
                const currentRoles: string[] = [];
                //iterate over metadata columns to get current role count
                options.dataViews[zeroVal].table.columns.forEach((column: DataViewMetadataColumn): void => {
                    Object.keys(column.roles).forEach((role: string): void => {
                        if (currentRoles.indexOf(role) !== negOneVal) { currentRoles.push(role); }
                    });
                });
                renderNewFlag = this.updateRenderNewFlag(currentRoles, renderNewFlag);
                root = self.visualRoot;
                root.x0 = containerWidth / twoVal;
                if (this.visualRootCalled && !renderNewFlag) {
                    updateTree(this.visualRootCalled);
                }
                else {
                    toggleAll(root);
                    toggle(root);
                    root.children.forEach(toggle);
                    updateTree(root);
                }
                selectionManager = this.selectionManager;
                const thisObj: any = this;
                let svgRootContainer: any = rootContainer.select('svg');
                //Register the events
                this.svgContainerUpdate(thisObj, svgRootContainer);
                this.events.renderingFinished(options);
            } catch (exception) {
                this.events.renderingFailed(options, exception);
            }
        }

        /**
         * This function parses the settings values from format pane options
         * @param dataView              - the dataview object, which contains all data needed to render the visual.
         */
        private static parseSettings(dataView: DataView): VisualSettings {
            return <VisualSettings>VisualSettings.parse(dataView);
        }

        /**
         * This function removes all elements from the root selection.
         */
        private clearAll(): void {
            this.rootContainer.selectAll('*').remove();
            d3.selectAll('.errorMessage').remove();
            d3.selectAll('.legendTitle').remove();
            d3.selectAll('.legendItem').remove();
        }

        /**
         * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
         * objects and properties you want to expose to the users in the property pane.
         * 
         * @param options                       - Map of defined objects
         */
        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
            let objectName: string = options.objectName;
            const objectEnumeration: VisualObjectInstance[] = [];
            let colorIndex: number = zeroVal;
            switch (objectName) {
                case 'card':
                    if (this.isGroupPresent) {
                        objectEnumeration.push({
                            objectName: options.objectName,
                            properties: {
                                dimension: this.settings.card.dimension,
                                cornerRadius: this.settings.card.cornerRadius,
                                borderShow: this.settings.card.borderShow
                            },
                            selector: null
                        });
                        if (this.settings.card.borderShow) {
                            objectEnumeration.push({
                                objectName: options.objectName,
                                properties: {
                                    borderColor: this.settings.card.borderColor,
                                    borderSize: this.settings.card.borderSize
                                },
                                selector: null
                            });
                        }
                    } else {
                        objectEnumeration.push({
                            objectName: options.objectName,
                            properties: {
                                dimension: this.settings.card.dimension,
                                backgroundFill: this.settings.card.backgroundFill,
                                cornerRadius: this.settings.card.cornerRadius,
                                borderShow: this.settings.card.borderShow
                            },
                            selector: null
                        });
                        if (this.settings.card.borderShow) {
                            objectEnumeration.push({
                                objectName: options.objectName,
                                properties: {
                                    borderColor: this.settings.card.borderColor,
                                    borderSize: this.settings.card.borderSize
                                },
                                selector: null
                            });
                        }
                    }
                    return objectEnumeration;
                case 'legend':
                    if (this.isGroupPresent) {
                        objectEnumeration.push({
                            objectName: options.objectName,
                            properties: {
                                show: this.settings.legend.show,
                                position: this.settings.legend.position,
                                labelColor: this.settings.legend.labelColor,
                                fontSize: this.settings.legend.fontSize,
                                showTitle: this.settings.legend.showTitle,
                            },
                            selector: null
                        });
                        if (this.settings.legend.showTitle) {
                            objectEnumeration.push({
                                objectName: options.objectName,
                                properties: { titleText: this.settings.legend.titleText },
                                selector: null
                            });
                        }
                        return objectEnumeration;
                    }
                case 'subLabelSettings':
                    if (this.isSublabelPresent) {
                        objectEnumeration.push({
                            objectName,
                            properties: {
                                fontSize: this.settings.subLabelSettings.fontSize,
                                fontFamily: this.settings.subLabelSettings.fontFamily,
                                textFill: this.settings.subLabelSettings.textFill
                            },
                            selector: null
                        });
                    }
                    return objectEnumeration;
                case 'colorSelector':
                    for (const barDataPoint of legendData.dataPoints) {
                        objectEnumeration.push({
                            objectName: options.objectName,
                            displayName: barDataPoint.label,
                            properties: { fill: { solid: { color: barDataPoint.color } } },
                            selector: legendData.dataPoints[colorIndex].identity
                        });
                        colorIndex++;
                    }
                    return objectEnumeration;
                default:
                    return VisualSettings.enumerateObjectInstances(this.settings || VisualSettings.getDefault(), options);
            }
        }
    }
}
