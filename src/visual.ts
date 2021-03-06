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
"use strict";

import "core-js/stable";
import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import * as AdaptiveCards from "adaptivecards";

import { VisualSettings } from "./settings";
export class Visual implements IVisual {
    private target: HTMLElement;
    private updateCount: number;
    private settings: VisualSettings;
    private textNode: Text;
    private cardLoad: boolean;
    private cardGet: boolean;
    private card: string;

    constructor(options: VisualConstructorOptions) {
        console.log('Visual constructor', options);
        this.target = options.element;
        this.updateCount = 0;
        this.cardLoad = false;
        this.cardGet = false;


        if (document) {
            const new_p: HTMLElement = document.createElement("p");
            new_p.appendChild(document.createTextNode("Update count:"));
            const new_em: HTMLElement = document.createElement("em");
            this.textNode = document.createTextNode(this.updateCount.toString());
            new_em.appendChild(this.textNode);
            new_p.appendChild(new_em);
            this.target.appendChild(new_p);
            this.getCard(this.target, "https://paradigmdownload.blob.core.windows.net/pbiviz/AdaptiveCards/simpletest.json");
            //this.loadCard(this.target,card);

        }


    }

    public update(options: VisualUpdateOptions) {
        this.settings = Visual.parseSettings(options && options.dataViews && options.dataViews[0]);
        console.log('Visual update', options);
        if (this.textNode) {
            this.textNode.textContent = (this.updateCount++).toString();
        }
        if (document && !(this.cardLoad)) {

            console.log("Get web page");
            this.cardLoad = true;
            // this.getCard(this.target,this.settings.sourceUrl.sourceUrl);

        }
    }

    private static parseSettings(dataView: DataView): VisualSettings {
        return <VisualSettings>VisualSettings.parse(dataView);
    }

    /**
     * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
     * objects and properties you want to expose to the users in the property pane.
     *
     */
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        return VisualSettings.enumerateObjectInstances(this.settings || VisualSettings.getDefault(), options);
    }
    public getCard(target: HTMLElement, url: string) {
        // using XMLHttpRequest
        /*
        let xhr = new XMLHttpRequest();
        let loadCard = this.loadCard;
        xhr.open("GET", this.settings.sourceUrl.sourceUrl, true);
        xhr.onload = function () {   
          target.appendChild(document.createTextNode("Load Card"));        
          //  loadCard(target,xhr.responseText);
          console.log("Got the Card")
          loadCard(target,card);
        }
        xhr.onerror = function () {
           // loadCard(target,"Document not loaded, check Url");
        }
        xhr.send();
        */
        var xhttp = new XMLHttpRequest();
        let loadCard = this.loadCard;
        xhttp.onreadystatechange = function () {
            if (this.readyState == 4 && this.status == 200) {
                // Typical action to be performed when the document is ready:
                //document.getElementById("demo").innerHTML = xhttp.responseText;
                target.appendChild(document.createTextNode("Card Loaded")); 
                let card = JSON.parse(xhttp.responseText) ;
                loadCard(target,card);
            }
        };
        xhttp.open("GET", url, true);
        xhttp.send();

        //this.loadCard(target,card);
    }
    public loadCard(target: HTMLElement, card) {

        var adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.hostConfig = new AdaptiveCards.HostConfig({
            fontFamily: "Segoe UI, Helvetica Neue, sans-serif"
            // More host config options
        });
        adaptiveCard.onExecuteAction = function (action) { console.log(adaptiveCard.toJSON()) }
        adaptiveCard.parse(card);

        // Render the card to an HTML element:
        console.log("Render Card")
        var renderedCard = adaptiveCard.render();
        //console.log(renderedCard);

        // And finally insert it somewhere in your page:
        target.appendChild(renderedCard);

    }
}