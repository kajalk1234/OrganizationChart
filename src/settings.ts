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
  'use strict';
  import DataViewObjectsParser = powerbi.extensibility.utils.dataview.DataViewObjectsParser;

  export class VisualSettings extends DataViewObjectsParser {
    public links: Links = new Links();
    public card: Card = new Card();
    public legend: Legend = new Legend();
    public colorSelector : Colorselector = new Colorselector();
    public labelSettings: Labelsettings = new Labelsettings();
    public subLabelSettings: Sublabelsettings = new Sublabelsettings();
  }

  export class Links {
    public view: string = 'classic';
    public color: string = '#FD625E';
    public width: number = 2;
  }

  export class Card {
    public dimension: number = 90;
    public backgroundFill: string = '#1595C4';
    public cornerRadius: number = 5;
    public borderShow: boolean = false;
    public borderColor: string = '#222';
    public borderSize: number = 1;
  }

  export class Legend {
    public show: boolean = true;
    public position : string = 'Top';
    public showTitle : boolean = true;
    public titleText : string = '';
    public labelColor : string = '#A9A9A9';
    public fontSize : number = 12;
  }

  export class Colorselector {
    public fill: string = '#FFF';
  }

  export class Labelsettings {
    public fontSize: number = 12;
    public fontFamily: string = 'Segoe UI';
    public textFill: string = '#FFF';
  }

  export class Sublabelsettings {
    public fontSize: number = 12;
    public fontFamily: string = 'Segoe UI';
    public textFill: string = '#FFF';
  }
}
