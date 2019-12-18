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
module powerbi.extensibility.visual {
    import valueFormatter = powerbi.extensibility.utils.formatting.valueFormatter;
    import IValueFormatter = powerbi.extensibility.utils.formatting.IValueFormatter;
    export class Visual implements IVisual {
        // stores the entire data that is selected by the user
        private static oData: DataViewTableRow[];
        // stores the index of data which is to be added next
        private static iCurrentPosition: number = 0;
        // number of KPI that is to be shown at a time. Possible values : 1, 2, 3, 4
        private static iNumberOfKPI: number = 4;
        // stores the dataView of the visual
        private static oDataView: DataView;
        // stores the font size of elements of the visual
        private static iValueFontSize: number;
        private static iNameFontSize: number;
        // stores the font color of elements of the visual
        private static iValueFontColor: string;
        private static iNameFontColor: string;
        // stores the font family of elements of the visual
        private static iNameFontFamily: string;
        private static iValueFontFamily: string;
        // stores the background color of the container
        private static iHorizontalScroll: boolean;
        private static iVerticalScroll: boolean = true;
        private static iHorizontalStack: boolean = true;
        private static iVerticalStack: boolean;
        private static iBackgroundColor: string;
        // stores the value of time for which current data will appear
        private static iDelay: number = 1200;
        // stores the value of time after which next data will appear
        private static iDuration: number = 4000;
        private static iDurationS: number;
        // stores the index of meaure KPI Status
        private static iIndexOfStatus: number;
        // stores the index of category KPI Nane
        private static iIndexOfName: number;
        // stores the index of measure KPI Last Value
        private static iIndexOfLastValue: number;
        // stores the index of measure KPI Current Value
        private static iIndexOfCurrentValue: number;
        // stores the interval value
        private static iInterval: number = -1;
        // stores the timeout value
        private static iTimeout: number = -1;
        // stores the color for positive indicator
        private static iPositiveIndicatorColor: string;
        // stores the color for negative indicator
        private static iNegativeIndicatorColor: string;
        // stores the color for neutral indicator
        private static iNeutralIndicatorColor: string;
        //Give Threshold in terms of percentage
        private static iPositiveThresholdPercentage: number;
        private static iNegativeThresholdPercentage: number;
        //Get Color for  Depending on threshold
        private static iPositiveThresholdIndicatorColor: string;
        private static iNegativeThresholdIndicatorColor: string;
        private static iNeutralThresholdIndicatorColor: string;
        //Incase of Postive or Negative Data Bag
        private static iPositiveThresholdValue: number;
        private static iNegativeThresholdValue: number;
        // stores the flag variable to check if index is exceeding the length of data
        private static bFlag: boolean;
        // stores the index to check if index is exceeding data length
        private static iCheckIndex: number;
        // tells where to continue from in case the index exceeds data length
        private static iFlagIndex: number;
        // stores if change percentage is to be shown or not
        private static iEnableDelta: number;
        // stores the information if the visual is updated or not
        private static iAnimation: number;
        private static bIsUpdated: boolean;
        private static iAnimationStyle: string;
        // stores the dynamic height and width
        private static dynamicWidth: number;
        private static dynamicHeight: number;
        private static iResponsive: boolean;
        // stores the height and width of tiles when resposnive is turned off
        private static iHeightOfTiles: number;
        private static iWidthOfTiles: number;
        // stores the position of text
        private static iNameAlignment: string;
        // stores the limits for height and width in vertical stacking
        private static iMaxDynamicWidthVertical: number = 291;
        private static iMaxDynamicHeightVertical: number = 690;
        private static iMinDynamicWidthVertical: number = 290;
        private static iMinDynamicHeightVertical: number = 320;
        // stores the limits for height and width in horizontal stacking
        private static iMaxDynamicWidthHorizontal: number;
        private static iMaxDynamicHeightHorizontal: number = 80;
        private static iMinDynamicWidthHorizontal: number = 1140;
        private static iMinDynamicHeightHorizontal: number = 80;
        // limits for height and width in vertical stacking
        private static iMaxWidthOfTilesVertical: number = 400;
        private static iMaxHeightOfTilesVertical: number = 120;
        private static iMinWidthOfTilesVertical: number = 260;
        private static iMinHeightOfTilesVertical: number = 80;
        // limits for height and width in horizontal stacking
        private static iMaxWidthOfTilesHorizontal: number = 300;
        private static iMaxHeightOfTilesHorizontal: number = 200;
        private static iMinWidthOfTilesHorizontal: number = 260;
        private static iMinHeightOfTilesHorizontal: number = 80;
        // maintain max/min KPI Count
        private static iMaxKPICount: number = 4;
        private static iMinKPICount: number = 1;
        // stores value to check if animation is On or Off
        private static iShowAnimation: boolean;
        // stores min and max duration for animation
        private static iMinDuration: number = 2;
        private static iMaxDuration: number = 10;
        // stores fade duration
        private static iFadeDuration: number = 2;
        private static iFadeInDuration: number = 1000;
        // stores margins and borders
        private static iMarginForScroll: number = 15;
        private static iMarginForKPIName: number = 11;
        private static iBorderOfContainer: number = 10;
        private static iMarginForCarousel: number = 6;
        private static iMarginForTop: number = 6;
        private static iMarginForLeft: number = 2;
        // stores value to check if carousel is On or Off
        private static iShowCarousel: boolean;
        // stores display units and decimal places
        private static iDisplayUnits: number;
        private static iDecimalPlaces: number;
        private static displayVal: number;
        // stores total no. of tiles
        private static iNoOfTiles: number;
        private static iMaxCurrentValueWidth: number;
        private static iMaxPriceChangeValueWidth: number;
        private static iMaxDeltaWidth: number;

        // rendering events api
        private events: IVisualEventService ;

        /*
        * Creates instance of KPIIndicator. This method is only called once.
        * @param {VisualConstructorOptions} options - Contains references to the element that will
        *                                             contain the visual and a reference to the host
        *                                             which contains services.
        *
        */
        constructor(options: VisualConstructorOptions) {
            // this is to make the parent container
            d3.select(options.element).append('div').attr('id', 'wrapper');
            d3.select(options.element).append('div').attr('id', 'scrollArrows');
            d3.select('body').style('overflow', 'auto');

            this.events = options.host.eventService;

        }
        /*
        * function to updates the state of the visual. Every sequential databinding and resize will call update.
        * @param {VisualUpdateOptions} options - Contains references to the size of the container
        *                                        and the dataView which contains all the data
        *                                        the visual had queried.
        */
        // tslint:disable-next-line:cyclomatic-complexity
        public update(options: VisualUpdateOptions): void {
            this.events.renderingStarted(options);
            const kpiName: string = 'kpiName';
            const kpiCurrentValue: string = 'kpiCurrentValue';
            const kpiLastValue: string = 'kpiLastValue';
            const kpiStatus: string = 'kpiStatus';
            const kpiPositiveThresholdValue: string = 'kpiPositiveThresholdValue';
            const kpiNegativeThresholdValue: string = 'kpiNegativeThresholdValue';
            let padding = 5;
            Visual.dynamicWidth = options.viewport.width - padding;
            Visual.dynamicHeight = options.viewport.height;
            //clear interval and timeout when update is called
            if (Visual.iInterval !== -1) {
                window.clearTimeout(Visual.iInterval);
            }
            if (Visual.iTimeout !== -1) {
                window.clearTimeout(Visual.iTimeout);
            }
            // check if basic requirements are satisfied else return
            if (options.dataViews.length === 0 || !options.dataViews[0].categorical ||
                ((!options.dataViews[0].categorical.values))) {
                Visual.displayBasicRequirement(1);
                return;
            }
            if ((!options.dataViews[0].categorical.categories)) {
                Visual.displayBasicRequirement(4);
                return;
            }
            // initializing Visual.iCurrentPosition to zero
            Visual.iCurrentPosition = 0;

            if(Visual.appendIndex(options,kpiName,kpiCurrentValue,kpiLastValue,kpiStatus,kpiPositiveThresholdValue,kpiNegativeThresholdValue)) {
                return;
            }
            // storing all the data in one variable
            const len: number = Visual.oDataView.categorical.categories[0].values.length;
            const categoriesLength: number = Visual.oDataView.categorical.categories.length;
            const valuesLength: number = Visual.oDataView.categorical.values.length;
            const cLength: number = Visual.oDataView.metadata.columns.length;
            let iRow: number;
            let iColumn: number;
            let kIterator: number = 0;
            let jIterator: number = 0;
            // tslint:disable-next-line:no-any
            const data: any[] = [];
            for (iRow = 0; iRow < len; iRow++) {
                data[iRow] = [];
                kIterator = 0, jIterator = 0;
                for (iColumn = 0; iColumn < cLength; iColumn++) {
                    if (Visual.oDataView.metadata.columns[iColumn].isMeasure === true) {
                        data[iRow][iColumn] = Visual.oDataView.categorical.values[kIterator++].values[iRow];
                    } else {
                        data[iRow][iColumn] = Visual.oDataView.categorical.categories[jIterator++].values[iRow];
                    }
                }
            }

            if(Visual.normalize(data)){
                return;
            }
            //Getting Length of Input Threshold Percentage and splitting Dot Seperated string into different index of array
            
            Visual.splitThresholdString();

            // On and Off Responsive
            Visual.onAndOff();

            // Set duration according to the format pane
            Visual.setDuration();

            // creating wrapper1 initially to start the visual
            Visual.createWrapper(1);

            // change the top of wrapper1 to initially show it on the screen
            Visual.wrapperCss();

            let iDivStart:number = 1;
            // the visual is updated
            Visual.bIsUpdated = true;
            // populating the wrapper1 that was created
            
            Visual.populateWrapper(1, iDivStart);

            // Apply carousel feature if toggle is turned on
            Visual.applyCarousel();
            // change the value of Visual.iCurrentPosition to number of containers
            Visual.iCurrentPosition = Visual.iNumberOfKPI;
            // call add next data in fixed timeout only if some slicer is not applied or
            //the number of data is equal to the number of containers.
            if (!(Visual.iNumberOfKPI === Visual.oData.length) && (Visual.iShowAnimation === true)) {
                Visual.iInterval = window.setTimeout(Visual.addNextData, Visual.iDuration);
            }
            this.events.renderingFinished(options);
        }

        // apply css to top of wrapper1 to initially show it on the screen
        private static wrapperCss() {
            // Inline CSS required here
            $('#wrapper1').css({ top: '0px', left: '0px' });
            // change the height/width of wrapper initially
            if (Visual.iResponsive) {
                if (Visual.iVerticalStack) {
                    $('#wrapper').css({
                        height: `${Visual.dynamicHeight - Visual.iMarginForScroll - 20}px`,
                        width: `${Visual.iMaxDynamicWidthVertical}px`
                    });
                } else {
                    $('#wrapper').css({
                        height: `${Visual.iMaxDynamicHeightHorizontal}px`,
                        width: `${Visual.dynamicWidth - Visual.iMarginForScroll * 2}px`
                    });
                }

            } else {

                if (Visual.iVerticalStack) {

                    $('#wrapper').css({
                        height: `${(Visual.iHeightOfTiles * Visual.iNumberOfKPI)}px`,
                        width: `${Visual.iWidthOfTiles}px`
                    });
                } else {
                    $('#wrapper').css('height', `${Visual.iHeightOfTiles}px`)
                        .css('width', `${(Visual.iWidthOfTiles * Visual.iNumberOfKPI)}px`);
                }
            }
        }

        private static applyCarousel() {
            if (Visual.iNoOfTiles > 1 && Visual.iShowCarousel) {
                if (Visual.iHorizontalScroll) { // if horizontal scrolling is on then previous and next arrows should be there
                    // previously scrolled data should be appeared
                    $('<div>').attr('id', 'prev').addClass('slideArrows').appendTo('#scrollArrows')
                        // tslint:disable-next-line:typedef
                        .on('click', ()=> {
                            // Reset duration and delay to 0 so that there is no animation when clicked on carousel
                            Visual.iDuration = 0;
                            Visual.iDelay = 0;
                            // subtracting 2 times as addNextdata function is adding numberofKPI in eachcase
                            Visual.iCurrentPosition -= (Visual.iNumberOfKPI * 2);
                            clearTimeout(Visual.iInterval);
                            Visual.addNextData();
                        });
                    $('<div>').attr('id', 'next').addClass('slideArrows').appendTo('#scrollArrows')
                        // tslint:disable-next-line:typedef
                        .on('click', ()=> {
                            // Reset duration and delay to 0 so that there is no animation when clicked on carousel
                            Visual.iDuration = 0;
                            Visual.iDelay = 0;
                            clearTimeout(Visual.iInterval);
                            Visual.addNextData();
                        });

                    if (Visual.iResponsive) { // if responsive is turned on change the top,left of arrows according to the stacking
                        if (Visual.iHorizontalStack) {

                            $('.slideArrows').css({
                                top: `${(Visual.iMaxDynamicHeightHorizontal
                                    + Visual.iMarginForScroll) / 2}px`
                            });
                            $('#prev').css('margin-top', `${Visual.iMarginForTop + Visual.iMarginForLeft}px`);
                        } else {

                            $('.slideArrows').css({ top: `${(Visual.dynamicHeight / 2) - Visual.iMarginForTop}px` });
                            $('#prev').css('margin-top', `${Visual.iMarginForTop + Visual.iMarginForLeft}px`);
                            $('#next').css('left', `${Visual.iMaxDynamicWidthVertical + Visual.iMarginForCarousel}px`);
                        }
                    } else {
                        if (Visual.iHorizontalStack) {

                            $('.slideArrows').css({ top: `${(Visual.iHeightOfTiles / 2) + Visual.iMarginForTop}px` });
                            $('#prev').css('margin-top', `${Visual.iMarginForTop + Visual.iMarginForLeft}px`);
                            $('#next').css('left', `${(Visual.iWidthOfTiles * Visual.iNumberOfKPI)
                                + Visual.iMarginForCarousel}px`);
                        } else {
                            $('.slideArrows').css({
                                top: `${((Visual.iHeightOfTiles * Visual.iNumberOfKPI) / 2) + Visual.iBorderOfContainer}px`
                            });
                            $('#prev').css('margin-top', `${Visual.iMarginForTop + Visual.iMarginForLeft}px`);
                            $('#next').css('left', `${Visual.iWidthOfTiles + Visual.iMarginForCarousel}px`);
                        }
                    }
                } else { // if horizontal is off i.e. vertical scrolling is on then top and bottom arrows should be there
                    $('<div>').attr('id', 'top').addClass('slideArrows').appendTo('#scrollArrows')
                        // tslint:disable-next-line:typedef
                        .on('click', ()=> {
                            // Reset duration and delay to 0 so that there is no animation when clicked on carousel
                            Visual.iDuration = 0;
                            Visual.iDelay = 0;
                            clearTimeout(Visual.iInterval);
                            Visual.addNextData();
                        });
                    $('<div>').attr('id', 'bottom').addClass('slideArrows').appendTo('#scrollArrows')
                        // tslint:disable-next-line:typedef
                        .on('click', ()=> {
                            // Reset duration and delay to 0 so that there is no animation when clicked on carousel
                            Visual.iDuration = 0;
                            Visual.iDelay = 0;
                            Visual.iCurrentPosition -= (Visual.iNumberOfKPI * 2);
                            clearTimeout(Visual.iInterval);
                            Visual.addNextData();
                        });
                    if (Visual.iResponsive) {
                        if (Visual.iHorizontalStack) {
                            $('.slideArrows').css({ left: `${(Visual.dynamicWidth - Visual.iBorderOfContainer * 4)}px` });
                        } else {
                            $('.slideArrows').css({ left: `${(Visual.iMaxDynamicWidthVertical - Visual.iBorderOfContainer)}px` });
                        }
                    } else {
                        if (Visual.iHorizontalStack) {
                            $('.slideArrows').css({
                                left: `${((Visual.iWidthOfTiles * Visual.iNumberOfKPI) - Visual.iBorderOfContainer)}px`
                            });
                        } else {
                            $('.slideArrows').css({ left: `${(Visual.iWidthOfTiles - Visual.iBorderOfContainer)}px` });
                        }
                    }
                }
            }
        }

        // Set duration according to the format pane and convert accordingly
        private static setDuration () {
            Visual.iDurationS = Visual.getValue<number>(Visual.oDataView, 'animation', 'duration', 2);
            if (Visual.iDurationS < Visual.iMinDuration) {
                Visual.iDurationS = Visual.iMinDuration;
            } else if (Visual.iDurationS > Visual.iMaxDuration) {
                Visual.iDurationS = Visual.iMaxDuration;
            }
            // convert duration into milliseconds
            Visual.iDuration = Visual.iDurationS * 1000;
            // set the value of delay according to the duration of animation in a particukar ratio
            Visual.iDelay = 3 * (Visual.iDuration / 10);

            if (Visual.iAnimationStyle === 'noAnimation') {
                Visual.iDelay = 0;
            } else if (Visual.iAnimationStyle === 'fade') {
                // set the duration as per the fade animation
                Visual.iDuration = (Visual.iDurationS + Visual.iFadeDuration) * 1000;
            }
        }
        
        // On and Off Responsive
        private static onAndOff() {
            Visual.iResponsive = Visual.getValue(Visual.oDataView, 'responsive', 'makeResponsive', true);

            // When responsive is turned off set minimum/maximum height and width of the tiles
            if (!(Visual.iResponsive)) {
                // Width of tiles when responsive is OFF
                Visual.iWidthOfTiles = Visual.getValue(Visual.oDataView, 'responsive', 'widthOfTiles', 290);
                // Height of tiles when responsive is OFF
                Visual.iHeightOfTiles = Visual.getValue(Visual.oDataView, 'responsive', 'heightOfTiles', 80);

                if (Visual.iVerticalStack) { // set min/max heigth and width in vertical stacking
                    if (Visual.iWidthOfTiles > Visual.iMaxWidthOfTilesVertical) {
                        Visual.iWidthOfTiles = Visual.iMaxWidthOfTilesVertical;
                    } else if (Visual.iWidthOfTiles < Visual.iMinWidthOfTilesVertical) {
                        Visual.iWidthOfTiles = Visual.iMinWidthOfTilesVertical;
                    }
                    if (Visual.iHeightOfTiles > Visual.iMaxHeightOfTilesVertical) {
                        Visual.iHeightOfTiles = Visual.iMaxHeightOfTilesVertical;
                    } else if (Visual.iHeightOfTiles < Visual.iMinHeightOfTilesVertical) {
                        Visual.iHeightOfTiles = Visual.iMinHeightOfTilesVertical;
                    }
                } else { // set min/max heigth and width in horizontal stacking
                    if (Visual.iWidthOfTiles > Visual.iMaxWidthOfTilesHorizontal) {
                        Visual.iWidthOfTiles = Visual.iMaxWidthOfTilesHorizontal;
                    } else if (Visual.iWidthOfTiles < Visual.iMinWidthOfTilesHorizontal) {
                        Visual.iWidthOfTiles = Visual.iMinWidthOfTilesHorizontal;
                    }
                    if (Visual.iHeightOfTiles > Visual.iMaxHeightOfTilesHorizontal) {
                        Visual.iHeightOfTiles = Visual.iMaxHeightOfTilesHorizontal;
                    } else if (Visual.iHeightOfTiles < Visual.iMinHeightOfTilesHorizontal) {
                        Visual.iHeightOfTiles = Visual.iMinHeightOfTilesHorizontal;
                    }
                }
            }
        }

        //Getting Length of Input Threshold Percentage and splitting Dot Seperated string into different index of array
        private static splitThresholdString () {
            const pPercentage: string = String(Visual.iPositiveThresholdPercentage);
            const pLPercentage: number = pPercentage.length;
            //getting Dot "." Position for the string
            const dotIndex: number = pPercentage.indexOf('.');
            //if value with decimal places were not assigned and if the entered percentage length is greater than 4
            if (dotIndex === -1 && pLPercentage > 4) {
                Visual.iPositiveThresholdPercentage = 9999.99;
            } else if (dotIndex !== -1) {
                //if value with decimal places is assigned whatever might be the value entered after dot it trims into 2 Decimal places
                Visual.iPositiveThresholdPercentage = Visual.iPositiveThresholdPercentage * 100;
                Visual.iPositiveThresholdPercentage = Visual.iPositiveThresholdPercentage -
                    Visual.iPositiveThresholdPercentage % 1;
                Visual.iPositiveThresholdPercentage = Visual.iPositiveThresholdPercentage / 100;
            } else if (Visual.iPositiveThresholdPercentage < 0) {
                Visual.iPositiveThresholdPercentage = 0;
            }
            //if value with decimal places were not assigned and if the entered percentage length is greater than 4
            const nPercentage: string = String(Visual.iNegativeThresholdPercentage);
            const nLPercentage: number = nPercentage.length;
            //getting Dot "." Position for the string
            const ndotIndex: number = nPercentage.indexOf('.');
            //if value with decimal places were not assigned and if the entered percentage length is greater than 4
            if (ndotIndex === -1 && nLPercentage > 4) {
                Visual.iNegativeThresholdPercentage = 9999.99;
            } else if (ndotIndex !== -1) {
                //if value with decimal places is assigned whatever might be the value entered after dot it trims into 2 Decimal places
                Visual.iNegativeThresholdPercentage = Visual.iNegativeThresholdPercentage * 100;
                Visual.iNegativeThresholdPercentage = Visual.iNegativeThresholdPercentage -
                    Visual.iNegativeThresholdPercentage % 1;
                Visual.iNegativeThresholdPercentage = Visual.iNegativeThresholdPercentage / 100;
            } else if (Visual.iNegativeThresholdPercentage < 0) {
                Visual.iNegativeThresholdPercentage = 0;
            }
            // The color of positive indicator
            Visual.iPositiveIndicatorColor = Visual.getFill(Visual.oDataView, 'positiveIndicatorColor');
            // The color of negative indicator
            Visual.iNegativeIndicatorColor = Visual.getFill(Visual.oDataView, 'negativeIndicatorColor');
            // The color of neutral indicator
            Visual.iNeutralIndicatorColor = Visual.getFill(Visual.oDataView, 'neutralIndicatorColor');
            // The color of positive threshold indicator
            Visual.iPositiveThresholdIndicatorColor = Visual.getFill(Visual.oDataView, 'positiveThresholdIndicatorColor');
            // The color of negative threshold indicator
            Visual.iNegativeThresholdIndicatorColor = Visual.getFill(Visual.oDataView, 'negativeThresholdIndicatorColor');
            // The color of neutral threshold indicator
            Visual.iNeutralThresholdIndicatorColor = Visual.getFill(Visual.oDataView, 'neutralThresholdIndicatorColor');
            // Style of Animation
            Visual.iAnimationStyle = Visual.getValue(Visual.oDataView, 'animation', 'animationStyle', 'slideAndWait');
            if (Visual.iShowAnimation === false) {
                Visual.iHorizontalScroll = false;
                Visual.iAnimationStyle = 'noAnimation';
            }
        }

        private static normalize(data: any[]): boolean {
            Visual.oData = data;
            // empty the main div when update is called
            $('#wrapper').empty();
            $('#scrollArrows').empty();

            // The number of containers. Possible values 1,2,3,4
            Visual.iNumberOfKPI = Visual.getValue<number>(Visual.oDataView, 'configuration', 'numberOfKPI', 4);
            // convert KPI count to integer if it is in decimal
            if (Visual.iNumberOfKPI % 1 !== 0) {
                Visual.iNumberOfKPI -= Visual.iNumberOfKPI % 1;
            }
            if (Visual.iNumberOfKPI > Visual.iMaxKPICount) {
                Visual.iNumberOfKPI = Visual.iMaxKPICount;
            } else if (Visual.iNumberOfKPI < Visual.iMinKPICount) {
                Visual.iNumberOfKPI = Visual.iMinKPICount;
            }
            if (Visual.oData.length < Visual.iNumberOfKPI) {
                Visual.iNumberOfKPI = Visual.oData.length;
            }
            // if Visual.iNumberOfKPI is still 0 that means there is no data after filters are applied
            if (Visual.iNumberOfKPI === 0) {
                Visual.displayBasicRequirement(0);
                return true;
            }
            // The font size of containers. We are normalizing it to be 15 at max as height is not changeable
            Visual.iNameFontSize = Visual.getValue<number>(Visual.oDataView, 'name', 'fontSize', 14);
            // Restrict max font size to 14
            if (Visual.iNameFontSize > 25) {
                Visual.iNameFontSize = 25;
            }
            Visual.iValueFontSize = Visual.getValue<number>(Visual.oDataView, 'value', 'fontSize', 14);
            // Restrict max font size to 14
            if (Visual.iValueFontSize > 25) {
                Visual.iValueFontSize = 25;
            }
            // Status of show change percentage
            Visual.iEnableDelta = Visual.getValue<number>(Visual.oDataView, 'configuration', 'enableDelta', 0);
            // Carousel feature to be on or off
            Visual.iNoOfTiles = Visual.oDataView.categorical.categories[0].values.length / Visual.iNumberOfKPI;
            if (Visual.iNoOfTiles > 1) {
                Visual.iShowCarousel = Visual.getValue<boolean>(Visual.oDataView, 'carousel', 'show', false);
            } else {
                Visual.iShowCarousel = false;
            }
            // Animation feature to be on or off
            if (Visual.iNoOfTiles > 1) {
                Visual.iShowAnimation = Visual.getValue<boolean>(Visual.oDataView, 'animation', 'show', true);
            } else {
                Visual.iShowAnimation = false;
            }
            // Change the scrolling to horizontal
            Visual.iHorizontalScroll = Visual.getValue<boolean>(Visual.oDataView, 'animation', 'horizontalScroll', false);
            if (Visual.iHorizontalScroll) {
                Visual.iVerticalScroll = false;
            } else {
                Visual.iVerticalScroll = true;
            }
            // Change the stacking to vertical
            Visual.iVerticalStack = Visual.getValue<boolean>(Visual.oDataView, 'animation', 'verticalStack', false);
            if (Visual.iVerticalStack) {
                Visual.iHorizontalStack = false;
            } else {
                Visual.iHorizontalStack = true;
            }
            // Display units and decimal places for values
            Visual.iDisplayUnits = Visual.getValue<number>(Visual.oDataView, 'value', 'displayUnits', 0);
            Visual.iDecimalPlaces = Visual.getValue<number>(Visual.oDataView, 'value', 'decimalPlaces', 0);
            // Allowed decimal places from 0 to 4 only
            if (Visual.iDecimalPlaces > 4) {
                Visual.iDecimalPlaces = 4;
            } else if (Visual.iDecimalPlaces < 0) {
                Visual.iDecimalPlaces = 0;
            } else {
                // tslint:disable-next-line:no-bitwise
                Visual.iDecimalPlaces = ~~Visual.iDecimalPlaces;
            }
            // The font color of Name and Value
            Visual.iValueFontColor = Visual.getFill(Visual.oDataView, 'valueFontColor');
            Visual.iNameFontColor = Visual.getFill(Visual.oDataView, 'nameFontColor');
            // The font family of Name and Value
            Visual.iNameFontFamily = Visual.getValue<string>(Visual.oDataView, 'name', 'fontFamily', 'Segoe UI');
            Visual.iValueFontFamily = Visual.getValue<string>(Visual.oDataView, 'value', 'fontFamily', 'Segoe UI');
            // Alignment of text
            Visual.iNameAlignment = Visual.getValue<string>(Visual.oDataView, 'name', 'alignName', 'left');
            // The background color of containers
            Visual.iBackgroundColor = Visual.getFill(Visual.oDataView, 'backgroundColor');
            //Get Threshold Percentage Value
            Visual.iPositiveThresholdPercentage = Visual.getValue(Visual.oDataView, 'threshold', 'PThresholdPercentage', null);
            Visual.iNegativeThresholdPercentage = Visual.getValue(Visual.oDataView, 'threshold', 'NThresholdPercentage', null);
            return false;
        }

        private static appendIndex(options: VisualUpdateOptions,kpiName: string,kpiCurrentValue: string, 
            kpiLastValue: string,kpiStatus: string,kpiPositiveThresholdValue: string,kpiNegativeThresholdValue: string): boolean {
             // to pass dataView as a parameter when formatting options are choosen
             Visual.oDataView = options.dataViews[0];
             let oDataCategorical: DataViewCategorical;
             oDataCategorical = Visual.oDataView.categorical;
             let iNumberOfValues: number;
             iNumberOfValues = oDataCategorical.values.length;
             let iNumberOfCategories: number;
             iNumberOfCategories = oDataCategorical.categories.length;
             let iIndex: number = 0;
 
             // initializing the Visual.iIndexOfName, Visual.iIndexOfStatus,
             //Visual.iIndexOfLastValue,Visual.iIndexOfCurrentValue to -1 so that
             //if they are not selected by user the value corresponding to them is not displayed
             Visual.iIndexOfName = -1;
             Visual.iIndexOfStatus = -1;
             Visual.iIndexOfLastValue = -1;
             Visual.iIndexOfCurrentValue = -1;
             Visual.iPositiveThresholdValue = -1;
             Visual.iNegativeThresholdValue = -1;
 
             // assigning proper index for category KPI Name
             for (iIndex = 0; iIndex < iNumberOfCategories; iIndex++) {
                 if (oDataCategorical.categories[iIndex].source.roles[kpiName]) {
                     Visual.iIndexOfName = iIndex;
                     break;
                 }
             }
             // assigning proper index for measures
             for (iIndex = 0; iIndex < iNumberOfValues; iIndex++) {
                 // assigning index for measure KPI Current Value
                 if (oDataCategorical.values[iIndex].source.roles[kpiCurrentValue]) {
                     Visual.iIndexOfCurrentValue = iIndex;
                 } else if (oDataCategorical.values[iIndex].source.roles[kpiLastValue]) { // assigning index for measure KPI Last Value
                     Visual.iIndexOfLastValue = iIndex;
                 } else if (oDataCategorical.values[iIndex].source.roles[kpiStatus]) { // assigning index for measure KPI Status
                     Visual.iIndexOfStatus = iIndex;
                     // assigning index for measure KPI Positive Threshold
                 } else if (oDataCategorical.values[iIndex].source.roles[kpiPositiveThresholdValue]) {
                     Visual.iPositiveThresholdValue = iIndex;
                     // assigning index for measure KPI Negative Threshold
                 } else if (oDataCategorical.values[iIndex].source.roles[kpiNegativeThresholdValue]) {
                     Visual.iNegativeThresholdValue = iIndex;
                 }
             }
             // if KPI current value or KPI name is not selected
             if (Visual.iIndexOfCurrentValue === -1 || Visual.iIndexOfName === -1) {
                 Visual.displayBasicRequirement(1);
 
                 return true;
             }
             //if status, Positive Threshold Data bag and Negative Threshold Data Bag were selected
             if (Visual.iIndexOfStatus !== -1 && (Visual.iPositiveThresholdValue !== -1 || Visual.iNegativeThresholdValue !== -1)) {
                 Visual.displayBasicRequirement(3);
 
                 return true;
             }
             // if status column has values other than -1,0 and 1
             if (Visual.iIndexOfStatus !== -1) {
                 let oStatusData: PrimitiveValue[];
                 oStatusData = Visual.oDataView.categorical.values[Visual.iIndexOfStatus].values;
                 let iLengthOfData: number;
                 iLengthOfData = oStatusData.length;
                 for (iIndex = 0; iIndex < iLengthOfData; iIndex++) {
                     if (oStatusData[iIndex] === null || !(oStatusData[iIndex] === 1 ||
                         oStatusData[iIndex] === -1 || oStatusData[iIndex] === 0)) {
                         Visual.displayBasicRequirement(2);
 
                         return true;
                     }
                 }
             }
             return false;
        }
        /*
        * method to display text if basic requirements are not satisfied
        */
        private static displayBasicRequirement(iStatus: number): void {
            $('#wrapper').empty();
            $('#wrapper').css('width', Visual.dynamicWidth);
            $('<p>').attr('id', 'textToDisplay').appendTo('#wrapper');
            $('#textToDisplay').css('width', Visual.dynamicWidth);
            if (iStatus === 1) {
                document.getElementById('textToDisplay').textContent = `Please select 'KPI current value' `;
            } else if (iStatus === 2) { // if appropriate column for status is not selected
                document.getElementById('textToDisplay').textContent = `Please select a column with values -1, 0 or 1 for 'KPI status' `;
            } else if (iStatus === 3) { // if status column and any of the positive or negative threshold data bag were selected
                // tslint:disable-next-line:max-line-length
                document.getElementById('textToDisplay').textContent = `Select either 'KPI status' or any of the 'KPI positive' or 'KPI negative' threshold data bag `;
            } else if (iStatus === 4) {
                document.getElementById('textToDisplay').textContent = `Please select 'KPI name' `;
            } else { // after filters are selected there is no data to display
                document.getElementById('textToDisplay').textContent = `No Data to display `;
            }
        }
        /*
        * method to enumerate through the objects defined in the capabilities and adds the properties to the format pane
        * @param {EnumerateVisualObjectInstancesOptions} options - Map of defined objects
        */
        // tslint:disable-next-line:cyclomatic-complexity
        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {

            let oObjectName: string;
            oObjectName = options.objectName;
            let oObjectEnumeration: VisualObjectInstance[];
            oObjectEnumeration = [];
            let oDataView: DataView;
            oDataView = Visual.oDataView;
            switch (oObjectName) {
                // enumerate containers object from capabilities.json
                case 'carousel':
                    oObjectEnumeration = Visual.enumerateCarousel(oObjectEnumeration);
                    break;
                case 'animation':
                    oObjectEnumeration = Visual.enumerateAnimation(oObjectEnumeration);
                    break;
                case 'responsive':
                    oObjectEnumeration = Visual.enumerateResponsive(oObjectEnumeration);
                    break;
                case 'name':
                    oObjectEnumeration = Visual.enumerateName(oObjectEnumeration);
                    break;
                case 'value':
                    oObjectEnumeration = Visual.enumerateValue(oObjectEnumeration);
                    break;
                case 'configuration':
                    oObjectEnumeration = Visual.enumerateConfiguration(oObjectEnumeration);
                    break;
                // enumerate indicators object from capabilities.json
                case 'indicators':
                    oObjectEnumeration = Visual.enumerateIndicators(oObjectEnumeration)
                    break;
                // enumerate threshold object from capabilities.json
                case 'threshold':
                    // enumerate threshold positive and negative input fields, if positive and negative threshold data bag were not selected
                    oObjectEnumeration = Visual.enumerateThreshold(oObjectEnumeration);
                    break;
                default:
                    break;
            }

            return oObjectEnumeration;
        }
        /*
        * method to enumerate through the Carousel defined in the capabilities and adds the properties to the format pane
        * @param {VisualObjectInstance} oObjectEnumeration - list of object instances
        */
        private static enumerateCarousel(oObjectEnumeration: VisualObjectInstance[]): VisualObjectInstance[] {
            if ((Visual.iNoOfTiles) > 1) {
                let oCarousel: VisualObjectInstance;
                oCarousel = {
                    objectName: 'carousel',
                    displayName: 'Carousel',
                    selector: null,
                    properties: {
                        show: Visual.iShowCarousel
                    }
                };
                oObjectEnumeration.push(oCarousel);
            }
            
            return oObjectEnumeration;
        }
        /*
        * method to enumerate through the Animation defined in the capabilities and adds the properties to the format pane
        * @param {VisualObjectInstance} oObjectEnumeration - list of object instances
        */
        private static enumerateAnimation(oObjectEnumeration: VisualObjectInstance[]): VisualObjectInstance[] {
            let oAnimation: VisualObjectInstance;
                oAnimation = {
                    objectName: 'animation',
                    displayName: 'Animation',
                    selector: null,
                    properties: {
                        show: Visual.iShowAnimation,
                        duration: Visual.iDurationS,
                        horizontalScroll: Visual.iHorizontalScroll,
                        verticalStack: Visual.iVerticalStack,
                        animationStyle: Visual.iAnimationStyle
                    }
                };
                oObjectEnumeration.push(oAnimation);

            return oObjectEnumeration;
        }
        /*
        * method to enumerate through the Responsive defined in the capabilities and adds the properties to the format pane
        * @param {VisualObjectInstance} oObjectEnumeration - list of object instances
        */
        private static enumerateResponsive (oObjectEnumeration: VisualObjectInstance[]): VisualObjectInstance[] {
            let oResponsive: VisualObjectInstance;
                    oResponsive = {
                        objectName: 'responsive',
                        displayName: 'Responsive',
                        selector: null,
                        properties: {
                            makeResponsive: Visual.iResponsive
                        }
                    };
                    if (Visual.iResponsive === false) {
                        oResponsive = {
                            objectName: 'responsive',
                            displayName: 'Responsive',
                            selector: null,
                            properties: {
                                makeResponsive: Visual.iResponsive,
                                widthOfTiles: Visual.iWidthOfTiles,
                                heightOfTiles: Visual.iHeightOfTiles
                            }
                        };
                    }
                    oObjectEnumeration.push(oResponsive);

            return oObjectEnumeration; 
        }
        /*
        * method to enumerate through the Name defined in the capabilities and adds the properties to the format pane
        * @param {VisualObjectInstance} oObjectEnumeration - list of object instances
        */
        private static enumerateName (oObjectEnumeration: VisualObjectInstance[]): VisualObjectInstance[] {
            let oName: VisualObjectInstance;
                    oName = {
                        objectName: 'name',
                        displayName: 'Name',
                        selector: null,
                        properties: {
                            fontSize: Visual.iNameFontSize,
                            nameFontColor: Visual.iNameFontColor,
                            fontFamily: Visual.iNameFontFamily,
                            alignName: Visual.iNameAlignment

                        }
                    };
                    oObjectEnumeration.push(oName);

            return oObjectEnumeration;
        }
        /*
        * method to enumerate through the Value defined in the capabilities and adds the properties to the format pane
        * @param {VisualObjectInstance} oObjectEnumeration - list of object instances
        */
        private static enumerateValue (oObjectEnumeration: VisualObjectInstance[]): VisualObjectInstance[] {
            let oValue: VisualObjectInstance;
                    oValue = {
                        objectName: 'value',
                        displayName: 'Value',
                        selector: null,
                        properties: {
                            fontSize: Visual.iValueFontSize,
                            valueFontColor: Visual.iValueFontColor,
                            fontFamily: Visual.iValueFontFamily,
                            displayUnits: Visual.iDisplayUnits,
                            decimalPlaces: Visual.iDecimalPlaces
                        }
                    };
            oObjectEnumeration.push(oValue);

            return oObjectEnumeration;
        }
        /*
        * method to enumerate through the Configuration defined in the capabilities and adds the properties to the format pane
        * @param {VisualObjectInstance} oObjectEnumeration - list of object instances
        */
        private static enumerateConfiguration (oObjectEnumeration: VisualObjectInstance[]): VisualObjectInstance[] {
            let oConfiguration: VisualObjectInstance;
                    if (Visual.iIndexOfLastValue !== -1) {
                        oConfiguration = {
                            objectName: 'configuration',
                            displayName: 'Formatting',
                            selector: null,
                            properties: {
                                numberOfKPI: Visual.iNumberOfKPI,
                                enableDelta: Visual.iEnableDelta,
                                backgroundColor: Visual.iBackgroundColor
                            }
                        };
                        oObjectEnumeration.push(oConfiguration);
                    } else {
                        oConfiguration = {
                            objectName: 'configuration',
                            displayName: 'Formatting',
                            selector: null,
                            properties: {
                                numberOfKPI: Visual.iNumberOfKPI,
                                backgroundColor: Visual.iBackgroundColor
                            }
                        };
                        oObjectEnumeration.push(oConfiguration);
                    }
            
            return oObjectEnumeration;
        }
        /*
        * method to enumerate through the Indicators defined in the capabilities and adds the properties to the format pane
        * @param {VisualObjectInstance} oObjectEnumeration - list of object instances
        */
        private static enumerateIndicators (oObjectEnumeration: VisualObjectInstance[]): VisualObjectInstance[] {
            if (Visual.iIndexOfStatus !== -1) {
                let oIndicators: VisualObjectInstance;
                oIndicators = {
                    objectName: 'indicators',
                    displayName: 'Indicators',
                    selector: null,
                    properties: {
                        positiveIndicatorColor: Visual.iPositiveIndicatorColor,
                        negativeIndicatorColor: Visual.iNegativeIndicatorColor,
                        neutralIndicatorColor: Visual.iNeutralIndicatorColor
                    }
                };
                oObjectEnumeration.push(oIndicators);
            }
            
            return oObjectEnumeration;
        }
        /*
        * method to enumerate through the Threshold defined in the capabilities and adds the properties to the format pane
        * @param {VisualObjectInstance} oObjectEnumeration - list of object instances
        */
        private static enumerateThreshold (oObjectEnumeration: VisualObjectInstance[]): VisualObjectInstance[] {
            if (Visual.iIndexOfLastValue !== -1 && Visual.iIndexOfStatus === -1
                && Visual.iPositiveThresholdValue === -1
                && Visual.iNegativeThresholdValue === -1) {
                let oThreshold: VisualObjectInstance;
                oThreshold = {
                    objectName: 'threshold',
                    displayName: 'Threshold',
                    selector: null,
                    properties: {
                        positiveThresholdIndicatorColor: Visual.iPositiveThresholdIndicatorColor,
                        PThresholdPercentage: Visual.iPositiveThresholdPercentage,
                        negativeThresholdIndicatorColor: Visual.iNegativeThresholdIndicatorColor,
                        NThresholdPercentage: Visual.iNegativeThresholdPercentage,
                        neutralThresholdIndicatorColor: Visual.iNeutralThresholdIndicatorColor
                    }
                };
                oObjectEnumeration.push(oThreshold);
            }
            // enumerate threshold with negative input field, if only positive threshold data bag is selected
            if (Visual.iIndexOfStatus === -1
                && Visual.iPositiveThresholdValue !== -1
                && Visual.iNegativeThresholdValue === -1) {
                let oThreshold: VisualObjectInstance;
                oThreshold = {
                    objectName: 'threshold',
                    displayName: 'Threshold',
                    selector: null,
                    properties: {
                        positiveThresholdIndicatorColor: Visual.iPositiveThresholdIndicatorColor,
                        negativeThresholdIndicatorColor: Visual.iNegativeThresholdIndicatorColor,
                        NThresholdPercentage: Visual.iNegativeThresholdPercentage,
                        neutralThresholdIndicatorColor: Visual.iNeutralThresholdIndicatorColor
                    }
                };
                oObjectEnumeration.push(oThreshold);
            }
            // enumerate threshold with positive input field, if only negative threshold data bag is selected
            if (Visual.iIndexOfStatus === -1
                && Visual.iNegativeThresholdValue !== -1
                && Visual.iPositiveThresholdValue === -1) {
                let oThreshold: VisualObjectInstance;
                oThreshold = {
                    objectName: 'threshold',
                    displayName: 'Threshold',
                    selector: null,
                    properties: {
                        positiveThresholdIndicatorColor: Visual.iPositiveThresholdIndicatorColor,
                        PThresholdPercentage: Visual.iPositiveThresholdPercentage,
                        negativeThresholdIndicatorColor: Visual.iNegativeThresholdIndicatorColor,
                        neutralThresholdIndicatorColor: Visual.iNeutralThresholdIndicatorColor
                    }
                };
                oObjectEnumeration.push(oThreshold);
            }
            // enumerate threshold with color field, if both positive and negative threshold data bag were selected
            if (Visual.iIndexOfStatus === -1
                && Visual.iPositiveThresholdValue !== -1
                && Visual.iNegativeThresholdValue !== -1) {
                let oThreshold: VisualObjectInstance;
                oThreshold = {
                    objectName: 'threshold',
                    displayName: 'Threshold',
                    selector: null,
                    properties: {
                        positiveThresholdIndicatorColor: Visual.iPositiveThresholdIndicatorColor,
                        negativeThresholdIndicatorColor: Visual.iNegativeThresholdIndicatorColor,
                        neutralThresholdIndicatorColor: Visual.iNeutralThresholdIndicatorColor
                    }
                };
                oObjectEnumeration.push(oThreshold);
            }

            return oObjectEnumeration;
        }

        /*
        * method to get the color of font or background whichever is needed
        * @param {DataView} oDataView - contains the DataView of options
        * @param {string} sKey - name of property whose value is needed
        */
        // tslint:disable-next-line:cyclomatic-complexity
        private static getFill(oDataView: DataView, sKey: string): string {
            const configuration: string = 'configuration';
            const indicators: string = 'indicators';
            const name: string = 'name';
            const value: string = 'value';
            const threshold: string = 'threshold';
            if (oDataView) {
                const oObjects: DataViewObjects = oDataView.metadata.objects;
                if (oObjects) {
                    // return appropriate value as per the formatting options selected
                    const oConfiguration: DataViewObject = oObjects[configuration];
                    if (oConfiguration) {
                        const oFill: Fill = <Fill>oConfiguration[sKey];
                        if (oFill) {
                            return oFill.solid.color;
                        }
                    }
                    const oIndicators: DataViewObject = oObjects[indicators];
                    if (oIndicators) {
                        const oFill: Fill = <Fill>oIndicators[sKey];
                        if (oFill) {
                            return oFill.solid.color;
                        }
                    }
                    // for font color
                    const oName: DataViewObject = oObjects[name];
                    if (oName) {
                        const oFill: Fill = <Fill>oName[sKey];
                        if (oFill) {
                            return oFill.solid.color;
                        }
                    }
                    const oValue: DataViewObject = oObjects[value];
                    if (oValue) {
                        const oFill: Fill = <Fill>oValue[sKey];
                        if (oFill) {
                            return oFill.solid.color;
                        }
                    }
                    const oThreshold: DataViewObject = oObjects[threshold];
                    if (oThreshold) {
                        const oFill: Fill = <Fill>oThreshold[sKey];
                        if (oFill) {
                            return oFill.solid.color;
                        }
                    }
                }
            }
            if ('nameFontColor' === sKey) {
                return '#043c74';
            } else if ('valueFontColor' === sKey) {
                return '#043c74';
            } else if ('backgroundColor' === sKey) {
                return '#efefef';
            } else if ('positiveIndicatorColor' === sKey) {
                return '#009900';
            } else if ('negativeIndicatorColor' === sKey) {
                return '#ff0000';
            } else if ('neutralIndicatorColor' === sKey) {
                return '#0000ff';
            } else if ('positiveThresholdIndicatorColor' === sKey) {  //Assigning default color for positive threshold
                return '#009900';
            } else if ('negativeThresholdIndicatorColor' === sKey) {  //Assigning default color for negative threshold
                return '#ff0000';
            } else if ('neutralThresholdIndicatorColor' === sKey) {  //Assigning default color for neutral threshold
                return '#0000ff';
            }
        }
        private static getValue<T>(oDataView: DataView, sProperty: string, sKey: string, defaultValue: T): T {
            if (oDataView) {
                const oObjects: DataViewObjects = oDataView.metadata.objects;
                if (oObjects) {
                    // return appropriate value as per the formatting options selected
                    const oProperties: DataViewObject = oObjects[sProperty];
                    if (oProperties) {
                        const value: T = <T>oProperties[sKey];
                        // only the key corresponding to that object should be updated.
                        if (value === undefined) {
                            return defaultValue;
                        }

                        return value;
                    }
                }
            }

            return defaultValue;
        }
        //
        // method to decide which class is to be used for what div and append html elements accordingly
        // @param {DataView} oDataView - contains the DataView of options
        // @param {string} sClassNames - class names that are to be applied to the div
        // @param {number} iIndicator - to tell if the value to be displayed is Change Percentage or Change Value
        // @param {number} iIndex - index of the data row whose value is to be populated

        // tslint:disable-next-line:no-any
        // tslint:disable-next-line:cyclomatic-complexity
        private static appendData(oDataView: DataView, sClassNames: string, iIndicator: number, iIndex: number, sDivIdName: string): void {

            // tslint:disable-next-line:no-any
            let sValueDisplayed: any;
            let iCurrentValue: number;
            let iLastValue: number;
            let tStatus: number;
            // this variable stores the percenatge from positive, negative threshold data bag
            let iPThresholdValue: number = 0;
            let iNThresholdValue: number = 0;
            // if iIndicator is 0, the value to be displayed is KPI Change Percentage
            if (iIndicator === 0) {
                // when both current, last data bag were selected and when status, positive, negative threshold data bag were not selected
                if (Visual.iIndexOfCurrentValue !== -1
                    && Visual.iIndexOfLastValue !== -1
                    && Visual.iIndexOfStatus === -1
                    && (this.iPositiveThresholdPercentage !== null
                        || this.iNegativeThresholdPercentage !== null)
                    && Visual.iPositiveThresholdValue === -1
                    && Visual.iNegativeThresholdValue === -1) {
                        Visual.perChangeForCurrLast(oDataView, sClassNames, iIndicator, iIndex, sDivIdName);
                }
                // when both current, last, positive data bag were selected and when status, negative threshold data bag were not selected
                if (Visual.iIndexOfCurrentValue !== -1
                    && Visual.iIndexOfLastValue !== -1
                    && Visual.iIndexOfStatus === -1
                    && Visual.iPositiveThresholdValue !== -1
                    && Visual.iNegativeThresholdValue === -1) {
                        Visual.perChangeForCurrLastPos(oDataView, sClassNames, iIndicator, iIndex, sDivIdName);
                }
                // when both current, last, negative data bag were selected and when status, positive threshold data bag were not selected
                if (Visual.iIndexOfCurrentValue !== -1
                    && Visual.iIndexOfLastValue !== -1
                    && Visual.iIndexOfStatus === -1
                    && Visual.iPositiveThresholdValue === -1
                    && Visual.iNegativeThresholdValue !== -1) {
                        Visual.perChangeCurrLastNeg(oDataView, sClassNames, iIndicator, iIndex, sDivIdName);
                }
                // when both current, last, positive, negative threshold data bag were selected and when status data bag were not selected
                if (Visual.iIndexOfCurrentValue !== -1
                    && Visual.iIndexOfLastValue !== -1
                    && Visual.iIndexOfStatus === -1
                    && Visual.iPositiveThresholdValue !== -1
                    && Visual.iNegativeThresholdValue !== -1) {
                        Visual.perChangeCurrLastPosNeg(oDataView, sClassNames, iIndicator, iIndex, sDivIdName);
                }

                // tslint:disable-next-line:triple-equals
                if (Visual.iIndexOfCurrentValue !== -1 && Visual.iIndexOfLastValue !== -1 && Visual.iEnableDelta == 1) {
                    Visual.perChangeCurrLastDel(oDataView, sClassNames, iIndicator, iIndex, sDivIdName);
                }
                // tslint:disable-next-line:triple-equals
            } else if (iIndicator == 1) {  // if iIndicator is 1, the value to be displayed is KPI Change Value
                // tslint:disable-next-line:triple-equals
                if (Visual.iIndexOfLastValue != -1 && Visual.iIndexOfCurrentValue != -1) {
                    const title: string = 'KPI Change Value: ';
                    iCurrentValue = <number>oDataView.categorical.values[Visual.iIndexOfCurrentValue].values[iIndex];
                    iLastValue = <number>oDataView.categorical.values[Visual.iIndexOfLastValue].values[iIndex];
                    if (iCurrentValue == null) {
                        sValueDisplayed = iLastValue;
                    } else if (iLastValue == null) {
                        sValueDisplayed = iCurrentValue;
                    } else {
                        sValueDisplayed = iCurrentValue - iLastValue;
                    }
                    // If display unit is selected as Auto
                    let displayVal: number = 0;
                    // tslint:disable-next-line:no-any
                    let tempdata: any = sValueDisplayed;
                    tempdata = Math.round(tempdata);
                    tempdata = Math.abs(tempdata);
                    const valLen: number = String(tempdata).length;
                    if (Visual.iDisplayUnits === 0) {
                        if (valLen > 9) {
                            displayVal = 1e9;
                        } else if (valLen <= 9 && valLen > 6) {
                            displayVal = 1e6;
                        } else if (valLen <= 6 && valLen >= 4) {
                            displayVal = 1e3;
                        } else {
                            displayVal = 10;
                        }
                    }
                    // Apply formatting according to the display unit and decimal places
                    const formatter: IValueFormatter = valueFormatter.create({
                        format: Visual.oDataView.categorical.values[Visual.iIndexOfCurrentValue].source.format ?
                            Visual.oDataView.categorical.values[Visual.iIndexOfCurrentValue].source.format :
                            valueFormatter.DefaultNumericFormat,
                        value: Visual.iDisplayUnits === 0 ? displayVal : Visual.iDisplayUnits,
                        precision: Visual.iDecimalPlaces
                    });
                    sValueDisplayed = formatter.format(sValueDisplayed);
                    d3.select(sDivIdName).append('div')
                        .classed(sClassNames, true)
                        .attr('title', title + sValueDisplayed)
                        .text(sValueDisplayed);
                }
            }
        }

        // subfunctionality of appendData method
        // when both current, last data bag were selected and when status, positive, negative threshold data bag were not selected
        private static perChangeForCurrLast (oDataView: DataView,sClassNames: string, iIndicator: number, iIndex: number, sDivIdName: string): void {
            let sValueDisplayed: any;
            let iCurrentValue: number;
            let iLastValue: number;
            let tStatus: number;
            // this variable stores the percenatge from positive, negative threshold data bag
            let iPThresholdValue: number = 0;
            let iNThresholdValue: number = 0;

            iCurrentValue = <number>oDataView.categorical.values[Visual.iIndexOfCurrentValue].values[iIndex];
            iLastValue = <number>oDataView.categorical.values[Visual.iIndexOfLastValue].values[iIndex];
            // if the last KPI value is 0, then the percentage change should be calculated with denominator as 1
            const title: string = 'KPI Change Value: '; // difference value of kpi current value and kpi last value
            // when Negative threshold Input is not given
            if (this.iNegativeThresholdPercentage === null) {
                if (iLastValue == null || iCurrentValue == null) {
                    sValueDisplayed = '-';
                    d3.select(sDivIdName).append('div')
                        .classed(sClassNames, true)
                        .attr('title', title + sValueDisplayed)
                        .text(sValueDisplayed);
                } else {
                    if (iLastValue === 0) {
                        sValueDisplayed = (((iCurrentValue - iLastValue) / 1) * 100).toFixed(2);
                    } else {
                        sValueDisplayed = (((iCurrentValue - iLastValue) / Math.abs(iLastValue)) * 100).toFixed(2);
                    }
                    if (sValueDisplayed === '0.00') { // when svaluedisplayed is equal to zero then neutral sign will applied
                        tStatus = 0;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else if (sValueDisplayed >= this.iPositiveThresholdPercentage && this.iPositiveThresholdPercentage !== 0) {
                        tStatus = 1;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else if (sValueDisplayed > 0 && sValueDisplayed <= this.iPositiveThresholdPercentage) {
                        tStatus = 0;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    }
                }
            } else if (this.iPositiveThresholdPercentage === null) { // when Positive Threshold input field is null
                if (iLastValue == null || iCurrentValue == null) {
                    sValueDisplayed = '-';
                    d3.select(sDivIdName).append('div')
                        .classed(sClassNames, true)
                        .attr('title', title + sValueDisplayed)
                        .text(sValueDisplayed);
                } else {
                    if (iLastValue === 0) {
                        sValueDisplayed = (((iCurrentValue - iLastValue) / 1) * 100).toFixed(2);
                    } else {
                        sValueDisplayed = (((iCurrentValue - iLastValue) / Math.abs(iLastValue)) * 100).toFixed(2);
                    }
                    if (sValueDisplayed === '0.00') { // when svaluedisplayed is equal to zero then neutral sign will applied
                        tStatus = 0;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else if (sValueDisplayed <= (-this.iNegativeThresholdPercentage) && this.iNegativeThresholdPercentage !== 0) {
                        tStatus = -1;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else if (sValueDisplayed < 0 && sValueDisplayed >= (-this.iNegativeThresholdPercentage)) {
                        tStatus = 0;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    }
                }
            } else { // when both the input fields either null or not null
                if (iLastValue == null || iCurrentValue == null) {
                    sValueDisplayed = '-';
                    d3.select(sDivIdName).append('div')
                        .classed(sClassNames, true)
                        .attr('title', title + sValueDisplayed)
                        .text(sValueDisplayed);
                } else {
                    if (iLastValue === 0) {
                        sValueDisplayed = (((iCurrentValue - iLastValue) / 1) * 100).toFixed(2);
                    } else {
                        sValueDisplayed = (((iCurrentValue - iLastValue) / Math.abs(iLastValue)) * 100).toFixed(2);
                    }
                    if (sValueDisplayed === '0.00') { // when svaluedisplayed is equal to zero then neutral sign will applied
                        tStatus = 0;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else if (sValueDisplayed >= this.iPositiveThresholdPercentage && this.iPositiveThresholdPercentage !== 0) {
                        tStatus = 1;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else if (sValueDisplayed <= (-this.iNegativeThresholdPercentage) && this.iNegativeThresholdPercentage !== 0) {
                        tStatus = -1;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else {
                        tStatus = 0;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    }
                }
            }
        }

        // subfunctionality of appendData method
        // when both current, last, positive data bag were selected and when status, negative threshold data bag were not selected
        private static perChangeForCurrLastPos (oDataView: DataView,sClassNames: string, iIndicator: number, iIndex: number, sDivIdName: string): void {

            // tslint:disable-next-line:no-any
            let sValueDisplayed: any;
            let iCurrentValue: number;
            let iLastValue: number;
            let tStatus: number;
            // this variable stores the percenatge from positive, negative threshold data bag
            let iPThresholdValue: number = 0;
            let iNThresholdValue: number = 0;

            iCurrentValue = <number>oDataView.categorical.values[Visual.iIndexOfCurrentValue].values[iIndex];
                    iLastValue = <number>oDataView.categorical.values[Visual.iIndexOfLastValue].values[iIndex];
                    iPThresholdValue = <number>oDataView.categorical.values[Visual.iPositiveThresholdValue].values[iIndex];
                    const title: string = 'KPI Change Value: '; // difference value of kpi current value and kpi last value
                    if (this.iNegativeThresholdPercentage !== null) { // Negative threshold value
                        // if the last KPI value is 0, then the percentage change should be calculated with denominator as 1
                        if (iLastValue == null || iCurrentValue == null) {
                            sValueDisplayed = '-';
                            d3.select(sDivIdName).append('div')
                                .classed(sClassNames, true)
                                .attr('title', title + sValueDisplayed)
                                .text(sValueDisplayed);
                        } else {
                            if (iLastValue === 0) {
                                sValueDisplayed = (((iCurrentValue - iLastValue) / 1) * 100).toFixed(2);
                            } else {
                                sValueDisplayed = (((iCurrentValue - iLastValue) / Math.abs(iLastValue)) * 100).toFixed(2);
                            }
                            if (sValueDisplayed === '0.00') { // when svaluedisplayed is equal to zero then neutral sign will applied
                                tStatus = 0;
                                Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                            } else if (sValueDisplayed >= iPThresholdValue && iPThresholdValue !== 0) {
                                tStatus = 1;
                                Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                            } else if (sValueDisplayed <= (-this.iNegativeThresholdPercentage) && this.iNegativeThresholdPercentage !== 0) {
                                tStatus = -1;
                                Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                            } else {
                                tStatus = 0;
                                Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                            }
                        }
                    } else {
                        // if the last KPI value is 0, then the percentage change should be calculated with denominator as 1
                        if (iLastValue == null || iCurrentValue == null) {
                            sValueDisplayed = '-';
                            d3.select(sDivIdName).append('div')
                                .classed(sClassNames, true)
                                .attr('title', title + sValueDisplayed)
                                .text(sValueDisplayed);
                        } else {
                            if (iLastValue === 0) {
                                sValueDisplayed = (((iCurrentValue - iLastValue) / 1) * 100).toFixed(2);
                            } else {
                                sValueDisplayed = (((iCurrentValue - iLastValue) / Math.abs(iLastValue)) * 100).toFixed(2);
                            }
                            if (sValueDisplayed === '0.00') { // when svaluedisplayed is equal to zero then neutral sign will applied
                                tStatus = 0;
                                Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                            } else if (sValueDisplayed >= iPThresholdValue && iPThresholdValue !== 0) {
                                tStatus = 1;
                                Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                            } else if (sValueDisplayed > 0 && sValueDisplayed <= iPThresholdValue) {
                                tStatus = 0;
                                Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                            }
                        }
                    }
        }

        // subfunctionality of appendData method
        // when both current, last, negative data bag were selected and when status, positive threshold data bag were not selected
        private static perChangeCurrLastNeg (oDataView: DataView,sClassNames: string, iIndicator: number, iIndex: number, sDivIdName: string): void {
            let sValueDisplayed: any;
            let iCurrentValue: number;
            let iLastValue: number;
            let tStatus: number;
            // this variable stores the percenatge from positive, negative threshold data bag
            let iPThresholdValue: number = 0;
            let iNThresholdValue: number = 0;

            iCurrentValue = <number>oDataView.categorical.values[Visual.iIndexOfCurrentValue].values[iIndex];
            iLastValue = <number>oDataView.categorical.values[Visual.iIndexOfLastValue].values[iIndex];
            iNThresholdValue = <number>oDataView.categorical.values[Visual.iNegativeThresholdValue].values[iIndex];
            // if the last KPI value is 0, then the percentage change should be calculated with denominator as 1
            const title: string = 'KPI Change Value: '; // difference value of kpi current value and kpi last value
            if (this.iPositiveThresholdPercentage !== null) {
                if (iLastValue == null || iCurrentValue == null) {
                    sValueDisplayed = '-';
                    d3.select(sDivIdName).append('div')
                        .classed(sClassNames, true)
                        .attr('title', title + sValueDisplayed)
                        .text(sValueDisplayed);
                } else {
                    if (iLastValue === 0) {
                        sValueDisplayed = (((iCurrentValue - iLastValue) / 1) * 100).toFixed(2);
                    } else {
                        sValueDisplayed = (((iCurrentValue - iLastValue) / Math.abs(iLastValue)) * 100).toFixed(2);
                    }
                    if (sValueDisplayed === '0.00') {   // when svaluedisplayed is equal to zero then neutral sign will applied
                        tStatus = 0;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else if (sValueDisplayed >= this.iPositiveThresholdPercentage && this.iPositiveThresholdPercentage !== 0) {
                        tStatus = 1;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else if (sValueDisplayed <= (-iNThresholdValue) && iNThresholdValue !== 0) {
                        tStatus = -1;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else {
                        tStatus = 0;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    }
                }
            } else {
                if (iLastValue == null || iCurrentValue == null) {
                    sValueDisplayed = '-';
                    d3.select(sDivIdName).append('div')
                        .classed(sClassNames, true)
                        .attr('title', title + sValueDisplayed)
                        .text(sValueDisplayed);
                } else {
                    if (iLastValue === 0) {
                        sValueDisplayed = (((iCurrentValue - iLastValue) / 1) * 100).toFixed(2);
                    } else {
                        sValueDisplayed = (((iCurrentValue - iLastValue) / Math.abs(iLastValue)) * 100).toFixed(2);
                    }
                    if (sValueDisplayed === '0.00') {   // when svaluedisplayed is equal to zero then neutral sign will applied
                        tStatus = 0;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else if (sValueDisplayed <= (-iNThresholdValue) && iNThresholdValue !== 0) {
                        tStatus = -1;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    } else if (sValueDisplayed < 0 && sValueDisplayed > (-this.iNegativeThresholdPercentage)) {
                        tStatus = 0;
                        Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                    }
                }
            }
        }

        // subfunctionality of appendData method
        // when both current, last, positive, negative threshold data bag were selected and when status data bag were not selected
        private static perChangeCurrLastPosNeg (oDataView: DataView,sClassNames: string, iIndicator: number, iIndex: number, sDivIdName: string): void {
            // tslint:disable-next-line:no-any
            let sValueDisplayed: any;
            let iCurrentValue: number;
            let iLastValue: number;
            let tStatus: number;
            // this variable stores the percenatge from positive, negative threshold data bag
            let iPThresholdValue: number = 0;
            let iNThresholdValue: number = 0;

            iCurrentValue = <number>oDataView.categorical.values[Visual.iIndexOfCurrentValue].values[iIndex];
            iLastValue = <number>oDataView.categorical.values[Visual.iIndexOfLastValue].values[iIndex];
            iPThresholdValue = <number>oDataView.categorical.values[Visual.iPositiveThresholdValue].values[iIndex];
            iNThresholdValue = <number>oDataView.categorical.values[Visual.iNegativeThresholdValue].values[iIndex];
            // if the last KPI value is 0, then the percentage change should be calculated with denominator as 1
            const title: string = 'KPI Change Value: '; // difference value of kpi current value and kpi last value
            if (iLastValue == null || iCurrentValue == null) {
                sValueDisplayed = '-';
                d3.select(sDivIdName).append('div')
                    .classed(sClassNames, true)
                    .attr('title', title + sValueDisplayed)
                    .text(sValueDisplayed);
            } else {
                if (iLastValue === 0) {
                    sValueDisplayed = (((iCurrentValue - iLastValue) / 1) * 100).toFixed(2);
                } else {
                    sValueDisplayed = (((iCurrentValue - iLastValue) / Math.abs(iLastValue)) * 100).toFixed(2);
                }
                if (sValueDisplayed === '0.00') { // when svaluedisplayed is equal to zero then neutral sign will applied
                    tStatus = 0;
                    Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                } else if (sValueDisplayed >= iPThresholdValue && iPThresholdValue !== 0) {
                    tStatus = 1;
                    Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                } else if (sValueDisplayed <= (-iNThresholdValue) && iNThresholdValue !== 0) {
                    tStatus = -1;
                    Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                } else {
                    tStatus = 0;
                    Visual.hresholdtliChangeImage(Visual.oDataView, tStatus, iIndex, sDivIdName);
                }
            }
        }

        // subfunctionality of appendData method
        // when both current, last, threshold data bag were selected and when Delta is enabled
        private static perChangeCurrLastDel(oDataView: DataView,sClassNames: string, iIndicator: number, iIndex: number, sDivIdName: string): void {

            let sValueDisplayed: any;
            let iCurrentValue: number;
            let iLastValue: number;
            let tStatus: number;
            // this variable stores the percenatge from positive, negative threshold data bag
            let iPThresholdValue: number = 0;
            let iNThresholdValue: number = 0;

            iCurrentValue = <number>oDataView.categorical.values[Visual.iIndexOfCurrentValue].values[iIndex];
            iLastValue = <number>oDataView.categorical.values[Visual.iIndexOfLastValue].values[iIndex];
            // if the last KPI value is 0, then the percentage change should be calculated with denominator as 1
            const title: string = 'KPI Change Percentage: ';
            if (iLastValue == null || iCurrentValue == null) {
                sValueDisplayed = '-';
                d3.select(sDivIdName).append('div')
                    .classed(sClassNames, true)
                    .attr('title', title + sValueDisplayed)
                    .text(sValueDisplayed);
            } else {
                if (iLastValue === 0) {
                    sValueDisplayed = (((iCurrentValue - iLastValue) / 1) * 100).toFixed(2);
                } else {
                    sValueDisplayed = (((iCurrentValue - iLastValue) / Math.abs(iLastValue)) * 100).toFixed(2);
                }
                const openBracket: string = '(';
                const closeBracket: string = ') ';
                const percent: string = '%';
                d3.select(sDivIdName).append('div')
                    .classed(sClassNames, true)
                    .attr('title', title + sValueDisplayed + percent)
                    .text(openBracket + sValueDisplayed + percent + closeBracket);
            }
        }

        /*
        * method to decide what indicator is to be used on the basis of status and display statistics about the kpi
        * @param {DataView} oDataView - DataView of the visual
        * @param {number} iIndex - Index of data to be loaded
        */
        private static tliChangeImage(oDataView: DataView, iIndex: number, sDivIdName: string): void {
            // to store the status of the data that is being populated
            let iTliStatus: number;
            let sKPICurrentValue: string;
            // if KPI Value column is selected populate it
            if (Visual.iIndexOfCurrentValue !== -1) {
                sKPICurrentValue = <string>oDataView.categorical.values[Visual.iIndexOfCurrentValue].values[iIndex];
                // if display unit for current value is selected as Auto
                if (Visual.iDisplayUnits === 0) {
                    Visual.displayVal = 0;
                    // tslint:disable-next-line:no-any
                    let tempdata: any = sKPICurrentValue;
                    tempdata = Math.round(tempdata);
                    tempdata = Math.abs(tempdata);
                    const valLen: number = String(tempdata).length;
                    if (valLen > 9) {
                        Visual.displayVal = 1e9;
                    } else if (valLen <= 9 && valLen > 6) {
                        Visual.displayVal = 1e6;
                    } else if (valLen <= 6 && valLen >= 4) {
                        Visual.displayVal = 1e3;
                    } else {
                        Visual.displayVal = 10;
                    }
                }

                // Apply formatting according to the display unit and decimal places
                const formatter: IValueFormatter = valueFormatter.create({
                    format: Visual.oDataView.categorical.values[Visual.iIndexOfCurrentValue].source.format ?
                        Visual.oDataView.categorical.values[Visual.iIndexOfCurrentValue].source.format :
                        valueFormatter.DefaultNumericFormat,
                    value: Visual.iDisplayUnits === 0 ? Visual.displayVal : Visual.iDisplayUnits,
                    precision: Visual.iDecimalPlaces
                });
                sKPICurrentValue = formatter.format(sKPICurrentValue);
                const title: string = 'KPI Current Value: ';
                if (sKPICurrentValue == null) {
                    d3.select(sDivIdName)
                        .append('div').classed('tliPrice', true).attr('title', title + sKPICurrentValue).text('-');
                } else {
                    d3.select(sDivIdName)
                        .append('div').classed('tliPrice', true).attr('title', title + sKPICurrentValue).text(sKPICurrentValue);
                }
            }
            // populate the other details on the basis of selection of Status column
            if (Visual.iIndexOfStatus !== -1) {
                // storing the value of status of current data to nTliStatus
                iTliStatus = Number(oDataView.categorical.values[Visual.iIndexOfStatus].values[iIndex]);
                switch (iTliStatus) {

                    // when nTliStatus is 0 that is no change therefore neutral value
                    case 0:
                        if (Visual.iIndexOfCurrentValue !== -1) {
                            d3.select(sDivIdName).append('div').classed('neutral', true).classed('indicator', true)
                                .attr('title', 'Neutral indicator');
                        }
                        Visual.appendData(oDataView, `tliChangePriceNeutral tliChangePrice`, 1, iIndex, sDivIdName);
                        Visual.appendData(oDataView, `tliChangeNeutral tliChange`, 0, iIndex, sDivIdName);
                        break;
                    // when nTliStatus is 1 that is positive change therefore positive value
                    case 1:
                        if (Visual.iIndexOfCurrentValue !== -1) {
                            d3.select(sDivIdName).append('div').classed('arrowUp', true).classed('arrow', true)
                                .attr('title', 'Positive indicator');
                        }
                        Visual.appendData(oDataView, `tliChangePricePositive tliChangePrice`, 1, iIndex, sDivIdName);
                        Visual.appendData(oDataView, `tliChangePositive tliChange`, 0, iIndex, sDivIdName);
                        break;
                    // when nTliStatus is -1 that is negative change therefore negative value
                    case -1:
                        if (Visual.iIndexOfCurrentValue !== -1) {
                            d3.select(sDivIdName).append('div').classed('arrowDown', true).classed('arrow', true)
                                .attr('title', 'Negative indicator');
                        }
                        Visual.appendData(oDataView, `tliChangePriceNegative tliChangePrice`, 1, iIndex, sDivIdName);
                        Visual.appendData(oDataView, `tliChangeNegative tliChange`, 0, iIndex, sDivIdName);
                        break;
                    default:
                        break;
                }
            } else { // if Visual.iIndexOfStatus is -1
                // to append indicators before kpi change value
                Visual.appendData(oDataView, `tliChange`, 0, iIndex, sDivIdName);
                Visual.appendData(oDataView, `tliChangePrice`, 1, iIndex, sDivIdName);
            }
        }
        // when status column is not selected depending on the user given threshold percentage this function will be called
        private static hresholdtliChangeImage(oDataView: DataView, tStatus: number, iIndex: number, sDivIdName: string): void {
            switch (tStatus) {
                // when nTliStatus is 0 that is no change therefore neutral value
                case 0:
                    if (Visual.iIndexOfCurrentValue !== -1) {
                        d3.select(sDivIdName).append('div').classed('neutral', true).classed('indicator', true)
                            .attr('title', 'Neutral indicator');
                    }
                    break;
                // when nTliStatus is 1 that is positive change therefore positive value
                case 1:
                    if (Visual.iIndexOfCurrentValue !== -1) {
                        d3.select(sDivIdName).append('div').classed('arrowUp', true).classed('arrow', true)
                            .attr('title', 'Positive indicator');
                    }
                    break;
                // when nTliStatus is -1 that is negative change therefore negative value
                case -1:
                    if (Visual.iIndexOfCurrentValue !== -1) {
                        d3.select(sDivIdName).append('div').classed('arrowDown', true).classed('arrow', true)
                            .attr('title', 'Negative indicator');
                    }
                    break;
                default:
                    break;
            }
        }

        /*
        * method to load data in the div
        * @param {DataView} oDataView - DataView of the visual
        * @param {number} nDivID - ID of div which is to be loaded
        * @param {number} iIndex - Index of data to be loaded
        */
        private static populateDiv(oDataView: DataView, nDivID: number, iIndex: number): void {
            // storing the div name to be used
            let sDivIdName: string;
            sDivIdName = `#container${nDivID}`;
            const className: string = '.tliName';
            const tliName: string = 'tliName';
            // populate name if KPI Name column is selected
            //AK47
            if (Visual.iResponsive) {

                if (Visual.iVerticalStack) {
                    if (Visual.iIndexOfName !== -1) {
                        d3.select(sDivIdName).append('div').classed('tliName', true)
                            .style({
                                'text-align': Visual.iNameAlignment, width: `${Visual.iMaxDynamicWidthVertical -
                                    Visual.iMarginForKPIName}px`
                            })
                            .classed(tliName + iIndex, true)
                            .style({
                                'text-align': Visual.iNameAlignment, width: `${Visual.iMaxDynamicWidthVertical -
                                    Visual.iMarginForKPIName}px`
                            })
                            .style('height', '35px')
                            .style('padding-top', '6px')
                            .text(<string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex])
                            .attr('title', <string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex]);
                        d3.select(className + iIndex).text(<string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex]);
                    }
                    Visual.tliChangeImage(Visual.oDataView, iIndex, sDivIdName);
                } else {
                    if (Visual.iIndexOfName !== -1) {
                        d3.select(sDivIdName).append('div').classed('tliName', true)
                            .style({
                                width: `${((Visual.dynamicWidth - Visual.iMarginForScroll * 2) / Visual.iNumberOfKPI)
                                    - Visual.iMarginForKPIName}px`
                                , 'text-align': Visual.iNameAlignment
                            })
                            .classed(tliName + iIndex, true)
                            .style('height', '35px')
                            .style('padding-top', '6px')
                            .style({
                                width: `${((Visual.dynamicWidth - Visual.iMarginForScroll * 2) / Visual.iNumberOfKPI)
                                    - Visual.iMarginForKPIName}px`
                                , 'text-align': Visual.iNameAlignment
                            })

                            .text(<string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex])
                            .attr('title', <string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex]);
                        d3.select(className + iIndex).text(<string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex]);
                    }
                    Visual.tliChangeImage(Visual.oDataView, iIndex, sDivIdName);
                }
            } else {
                if (Visual.iVerticalStack) {
                    if (Visual.iIndexOfName !== -1) {
                        d3.select(sDivIdName).append('div').classed('tliName', true)
                            .style({
                                width: `${(Visual.iWidthOfTiles - Visual.iMarginForKPIName)}px`,
                                'text-align': Visual.iNameAlignment
                            })
                            .classed(tliName + iIndex, true)
                            .style('height', '35px')
                            .style('padding-top', '6px')
                            .style({
                                width: `${(Visual.iWidthOfTiles - Visual.iMarginForKPIName)}px`,
                                'text-align': Visual.iNameAlignment
                            })
                            .text(<string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex])
                            .attr('title', <string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex]);
                        d3.select(className + iIndex).text(<string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex]);
                    }
                    Visual.tliChangeImage(Visual.oDataView, iIndex, sDivIdName);
                } else {
                    if (Visual.iIndexOfName !== -1) {
                        d3.select(sDivIdName).append('div').classed('tliName', true)
                            .style({ width: `${(Visual.iWidthOfTiles - Visual.iMarginForKPIName)}px` })
                            .style({ 'text-align': Visual.iNameAlignment })
                            .style('height', '35px')
                            .style('padding-top', '6px')
                            .classed(tliName + iIndex, true)
                            .style({ width: `${(Visual.iWidthOfTiles - Visual.iMarginForKPIName)}px` })
                            .text(<string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex])
                            .attr('title', <string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex]);
                        d3.select(className + iIndex).text(<string>oDataView.categorical.categories[Visual.iIndexOfName].values[iIndex]);
                    }
                    Visual.tliChangeImage(Visual.oDataView, iIndex, sDivIdName);
                }
            }
        }

        /*
        * method to create wrapper according to parameter passed.
        * @param {number} iWrapperID  - ID of the wrapper to be created
        */
        private static createWrapper(iWrapperID: number): void {
            let sWrapperName: string;
            sWrapperName = `wrapper${iWrapperID}`;
            let sWrapperDivName: string;
            let sClassOfContainer: string;
            let iStartPoint: number;
            let iEndPoint: number;
            let iIndex: number = 0;
            sClassOfContainer = `kpi${Visual.iNumberOfKPI}`;
            // append the wrapper with appropriate id to "wrapper" div and then change
            // its top so that it is below the existing wrapper
            // When responsive is turned on, assign dynamic height and width
            Visual.dynamicWrapper(sWrapperName);

            if (iWrapperID === 1) {
                iStartPoint = 1;
                iEndPoint = Visual.iNumberOfKPI;
            } else if (iWrapperID === 2) {
                iStartPoint = Visual.iNumberOfKPI + 1;
                iEndPoint = 2 * Visual.iNumberOfKPI;
            }
            // append div to the wrapper just created on the basis of which wrapper id was created and the number of containers
            for (iIndex = iStartPoint; iIndex <= iEndPoint; iIndex++) {

                sWrapperDivName = `container${iIndex}`;
                if (Visual.iVerticalStack) {
                    if (Visual.iResponsive) {
                        $('<div>').attr('id', sWrapperDivName).appendTo(`#${sWrapperName}`)
                            .css({
                                height: ((Visual.dynamicHeight - Visual.iMarginForScroll * 2) / Visual.iNumberOfKPI),
                                width: Visual.iMaxDynamicWidthVertical
                            })
                            .addClass('containers');
                    } else {
                        $('<div>').attr('id', sWrapperDivName).appendTo(`#${sWrapperName}`)
                            .css({
                                height: (Visual.iHeightOfTiles),
                                width: Visual.iWidthOfTiles
                            })
                            .addClass('containers');
                    }
                } else {
                    if (Visual.iResponsive) {
                        $('<div>').attr('id', sWrapperDivName).appendTo(`#${sWrapperName}`)
                            .css('width', ((Visual.dynamicWidth - Visual.iMarginForScroll * 2) / Visual.iNumberOfKPI) - 1)
                            .addClass('containers');
                    } else {
                        $('<div>').attr('id', sWrapperDivName).appendTo(`#${sWrapperName}`)
                            .css({ width: (Visual.iWidthOfTiles), height: Visual.iHeightOfTiles })
                            .addClass('containers');
                    }
                }
            }
        }

        private static dynamicWrapper(sWrapperName: string) {
            if (Visual.iResponsive) {
                if (Visual.iAnimationStyle !== 'fade') {
                    if (Visual.iVerticalStack) {
                        $('<div>').attr('id', sWrapperName).appendTo('#wrapper')
                            .css({
                                height: Visual.dynamicHeight - Visual.iMarginForScroll * 2,
                                width: Visual.iMaxDynamicWidthVertical
                            });
                        if (Visual.iHorizontalScroll) {
                            $(`#${sWrapperName}`).css('left', `${Visual.iMaxDynamicWidthVertical}px`);
                        } else {
                            $(`#${sWrapperName}`).css('top', `${Visual.dynamicHeight}px`);
                        }
                    } else {
                        $('<div>').attr('id', sWrapperName).appendTo('#wrapper')
                            .css({
                                height: Visual.iMinHeightOfTilesHorizontal,
                                width: Visual.dynamicWidth - Visual.iMarginForScroll * 2
                            });
                        if (Visual.iHorizontalScroll) {
                            $(`#${sWrapperName}`).css('left', `${Visual.dynamicWidth}px`);
                        } else {
                            $(`#${sWrapperName}`).css('top', `${Visual.iMinHeightOfTilesHorizontal}px`);
                        }
                    }
                } else {
                    if (Visual.iVerticalStack) {
                        $('<div>').attr('id', sWrapperName).appendTo('#wrapper')
                            .css({ height: Visual.dynamicHeight - Visual.iMarginForScroll * 2 });
                    } else {
                        $('<div>').attr('id', sWrapperName).appendTo('#wrapper')
                            .css({
                                height: Visual.iMinHeightOfTilesHorizontal,
                                width: Visual.dynamicWidth - Visual.iMarginForScroll * 2
                            });
                    }
                    $(`#${sWrapperName}`).addClass('initialWrapper');
                    $(`#${sWrapperName}`).hide().fadeIn(Visual.iFadeInDuration);
                }
            } else { // assign height and width according to the format pane

                if (Visual.iAnimationStyle !== 'fade') {
                    if (Visual.iVerticalStack) {
                        $('<div>').attr('id', sWrapperName).appendTo('#wrapper')
                            .css({
                                height: (Visual.iHeightOfTiles * Visual.iNumberOfKPI)
                                , width: Visual.iWidthOfTiles
                            });
                        if (Visual.iHorizontalScroll) {
                            $(`#${sWrapperName}`)
                                .css('padding-top', '6px')
                                .css('left', `${Visual.iWidthOfTiles}px`);
                        } else {
                            $(`#${sWrapperName}`)
                                .css('padding-top', '6px')
                                .css('top', `${Visual.iHeightOfTiles * Visual.iNumberOfKPI}px`);
                        }
                    } else {

                        $('<div>').attr('id', sWrapperName).appendTo('#wrapper')
                            .css({
                                height: Visual.iHeightOfTiles,
                                width: (Visual.iWidthOfTiles * Visual.iNumberOfKPI) + Visual.iBorderOfContainer
                            });
                        if (Visual.iHorizontalScroll) {
                            $(`#${sWrapperName}`).css('left', `${Visual.iWidthOfTiles * Visual.iNumberOfKPI
                                + Visual.iBorderOfContainer}px`);
                        } else {
                            $(`#${sWrapperName}`).css('top', `${Visual.iHeightOfTiles}px`);
                        }
                    }
                } else {
                    if (Visual.iVerticalStack) {
                        $('<div>').attr('id', sWrapperName).css('padding-top', '6px').appendTo('#wrapper')
                            .css({
                                height: (Visual.iHeightOfTiles * Visual.iNumberOfKPI)
                                , width: Visual.iWidthOfTiles
                            });
                    } else {
                        $('<div>').attr('id', sWrapperName).css('padding-top', '6px').appendTo('#wrapper')
                            .css({
                                height: Visual.iHeightOfTiles,
                                width: (Visual.iWidthOfTiles * Visual.iNumberOfKPI) + Visual.iBorderOfContainer
                            });
                    }
                    $(`#${sWrapperName}`).addClass('initialWrapper');
                    $(`#${sWrapperName}`).hide().fadeIn(Visual.iFadeInDuration);
                }
            }
        }

        /*
        * method to change the css of containers whenever update is called
        * The css is changed according to the formatting options
        * @param {number} cssDivStart - The id of div from which the wrapper starts
        */
        private static width(w:number):number {
            if(w>=190)
                return w/2.5;
            else if(w>=180 && w<190)
                return w/3;
            else if(w>=170 && w<180)
                return w/3.5;
            else if(w>=160 && w<170)   
                return w/4;
            else
                return w/4.3;
        }
        private static changeCSS(iCssDivStart: number): void {
            // change the css according to the number of KPI that are to be displayed at a time
            let iEndPoint: number = 0;
            let iIndex: number = 0;
            let sPriceMarginLeft: string;
            sPriceMarginLeft = '10px';
            // to decide how many div are there to change the css
            if (iCssDivStart === 1) {
                iEndPoint = Visual.iNumberOfKPI;
            } else {
                iEndPoint = 2 * Visual.iNumberOfKPI;
            }
            // change the values as per the number of containers selected in the format pane
            switch (Visual.iNumberOfKPI) {

                case 1:
                    $('.tliName').addClass('tliNameKPIOne');
                    $('.tliPrice').addClass('tliPriceKPIOne');
                    $('.tliChangePrice').addClass('tliChangePriceKPIOne');
                    $('.tliChange').addClass('tliChangeKPIOne');
                    $('.arrow').addClass('indicatorKPIOne');
                    break;
                case 2:
                    $('.tliName').addClass('tliNameKPITwo');
                    $('.arrow').addClass('indicatorKPITwo');
                    break;
                case 3:
                    $('.tliName').addClass('tliNameKPIThree');
                    $('.arrow').addClass('indicatorKPIThree');
                    break;
                default:
                    break;
            }
            if (Visual.iEnableDelta == 1) {
                Visual.iMaxCurrentValueWidth = $('.containers').width() / 3.8;
                Visual.iMaxPriceChangeValueWidth = $('.containers').width() / 3.8;
                Visual.iMaxDeltaWidth = this.width($('.containers').width());
            } else {
                Visual.iMaxCurrentValueWidth = $('.containers').width() / 2.6;
                Visual.iMaxPriceChangeValueWidth = $('.containers').width() / 2.6;
            }
            // change the background color of the containers on the basis of
            for (iIndex = iCssDivStart; iIndex <= iEndPoint; iIndex++) {
                const sContainerId: string = `#container${iIndex}`;
                $(sContainerId).css('background', <string>Visual.iBackgroundColor);
            }
            // change the css on the basis of font size selected in format pane
            $('.tliName').css('font-size', `${Visual.iNameFontSize}px`);
            $('.tliPrice').css('font-size', `${Visual.iValueFontSize}px`);
            $('.tliPrice').css('max-width', `${Visual.iMaxCurrentValueWidth}px`);
            $('.tliChangePrice').css('font-size', `${Visual.iValueFontSize}px`);
            $('.tliChangePrice').css('max-width', `${Visual.iMaxPriceChangeValueWidth}px`);
            $('.tliChange').css('font-size', `${Visual.iValueFontSize}px`);
            // change the css on the basis of font color selected in format pane
            $('.tliName').css('color', <string>Visual.iNameFontColor);
            $('.tliPrice').css('color', <string>Visual.iValueFontColor);
            // fontfamily
            $('.tliName').css('font-family', <string>Visual.iNameFontFamily);
            $('.tliPrice').css('font-family', <string>Visual.iValueFontFamily);
            $('.tliChange').css('max-width', `${Visual.iMaxDeltaWidth}px`);
            $('.tliChange').css('font-family', <string>Visual.iValueFontFamily);
            $('.tliChangePrice').css('font-family', <string>Visual.iValueFontFamily);
            // change the color of indicators and the font color as per the selection in format pane if the Status column is selected
            if (Visual.iIndexOfStatus !== -1) {
                $('.arrowDown').css('margin-bottom', `${Visual.iValueFontSize - 10}px`);
                $('.arrowUp').css('margin-bottom', `${Visual.iValueFontSize - 10}px`);
                $('.neutral').css('margin-bottom', `${Visual.iValueFontSize - 10}px`);
                $('.arrowDown').css('border-top-color', <string>Visual.iNegativeIndicatorColor);
                $('.tliChangeNegative').css('color', <string>Visual.iNegativeIndicatorColor);
                $('.tliChangePriceNegative').css('color', <string>Visual.iNegativeIndicatorColor);
                $('.neutral').css('background', <string>Visual.iNeutralIndicatorColor);
                $('.tliChangeNeutral').css('color', <string>Visual.iNeutralIndicatorColor);
                $('.tliChangePriceNeutral').css('color', <string>Visual.iNeutralIndicatorColor);
                $('.arrowUp').css('border-bottom-color', <string>Visual.iPositiveIndicatorColor);
                $('.tliChangePositive').css('color', <string>Visual.iPositiveIndicatorColor);
                $('.tliChangePricePositive').css('color', <string>Visual.iPositiveIndicatorColor);
            } else {  // if Status column is not selected then the font color is same as KPI Name and KPI Value
                $('.tliChange').css('color', <string>Visual.iValueFontColor);
                $('.tliChangePrice').css('color', <string>Visual.iValueFontColor);
            }
            // change the color of threshold indicator as per the selection in format pane if the Status column is not selected
            if (Visual.iIndexOfStatus === -1 && Visual.iIndexOfCurrentValue !== -1 && Visual.iIndexOfLastValue !== -1) {
                $('.arrowDown').css('border-top-color', <string>Visual.iNegativeThresholdIndicatorColor);
                $('.tliChangeNegative').css('color', <string>Visual.iNegativeThresholdIndicatorColor);
                $('.tliChangePriceNegative').css('color', <string>Visual.iNegativeThresholdIndicatorColor);
                $('.neutral').css('background', <string>Visual.iNeutralThresholdIndicatorColor);
                $('.tliChangeNeutral').css('color', <string>Visual.iNeutralThresholdIndicatorColor);
                $('.tliChangePriceNeutral').css('color', <string>Visual.iNeutralThresholdIndicatorColor);
                $('.arrowUp').css('border-bottom-color', <string>Visual.iPositiveThresholdIndicatorColor);
                $('.tliChangePositive').css('color', <string>Visual.iPositiveThresholdIndicatorColor);
                $('.tliChangePricePositive').css('color', <string>Visual.iPositiveThresholdIndicatorColor);
            }

            // if KPI Value is not selected only show other data with appropriate margin
            if (Visual.iIndexOfCurrentValue === -1) {
                $('.tliChangePrice').css('margin-left', sPriceMarginLeft);
            }
        }

        /*
        *method to add next data after duration is over
        */
        private static addNextData(): void {
            // Reset currentPosition to 0, if it becomes negative
            if (Visual.iCurrentPosition < 0) {
                Visual.iCurrentPosition = 0;
            }
            // add next data only if mouse is not on the wrapper
            if (!($('#wrapper').is(':hover'))) {
                // flag to check if the index has exceeded the data length
                Visual.bFlag = true;
                Visual.bIsUpdated = false;
                let iDivStart: number = 0;
                // to change the iCurrentPosition value
                Visual.iCheckIndex = 0;
                // to start with first value when div is empty but data is not available
                Visual.iFlagIndex = 0;
                if (Visual.iCurrentPosition !== Visual.oData.length - 1) {
                    Visual.iCurrentPosition = Visual.iCurrentPosition % (Visual.oData.length - 1);
                }
                // if wrapper1 is present, create wrapper2 and remove wrapper1 after animating it.
                if ($('#wrapper1').length) {
                    Visual.createWrapper(2);
                    iDivStart = Visual.iNumberOfKPI + 1;
                    Visual.populateWrapper(2, iDivStart);
                } else { // if wrapper2 is present, create wrapper1 and remove wrapper2 after animating it.
                    Visual.createWrapper(1);
                    iDivStart = 1;
                    Visual.populateWrapper(1, iDivStart);
                }
                // check if index has exceeded the length of data and populate accordingly
                if (Visual.bFlag) {
                    if (Visual.iCheckIndex === (Visual.oData.length - 1)) {
                        Visual.iCurrentPosition = 0;
                    } else {
                        Visual.iCurrentPosition += Visual.iNumberOfKPI;
                        if (Visual.iCurrentPosition > Visual.oData.length - 1) {
                            Visual.iCurrentPosition = 0;
                        }
                    }
                } else {
                    Visual.iCurrentPosition = Visual.iFlagIndex;
                }
            }
            // convert duration into milliseconds
            Visual.iDuration = Visual.iDurationS * 1000;
            // set the value of delay according to the duration of animation in a particukar ratio

            if ((Visual.iAnimationStyle !== 'noAnimation') && (Visual.iShowCarousel)) {
                Visual.iDelay = 3 * (Visual.iDuration / 10);
            }
            if (Visual.iShowAnimation === true) {
                Visual.iInterval = window.setTimeout(Visual.addNextData, Visual.iDuration);
            }
        }
        /*
        * method to populate wrapper which was created by addNextData and animate it
        * @param {number} iWrapperID - id of the wrapper that was created
        * @param {number} iDivStart - id of the first div in the wrapper created
        */
        private static populateWrapper(iWrapperID: number, iDivStart: number): void {
            let iIndex: number;
            iIndex = 0;
            Visual.iCheckIndex = 0;
            Visual.iFlagIndex = 0;
            Visual.bFlag = true;
            for (iIndex = Visual.iCurrentPosition; iIndex < Visual.iCurrentPosition + Visual.iNumberOfKPI; iIndex++) {
                Visual.iCheckIndex = iIndex;
                if (iIndex <= Visual.oData.length - 1) {
                    Visual.populateDiv(Visual.oDataView, iDivStart, iIndex);
                } else {
                    Visual.populateDiv(Visual.oDataView, iDivStart, Visual.iFlagIndex);
                    Visual.iFlagIndex++;
                    Visual.bFlag = false;
                }
                iDivStart++;
            }
            // change the css according to the default value or the custom value selected by the user
            Visual.changeCSS(iWrapperID);

            // animate the wrappers up only if it is not the first time
            if (!Visual.bIsUpdated) {
                if (Visual.iShowAnimation === true || Visual.iShowCarousel === true) {
                    Visual.animateWrapper(iWrapperID);
                }
            }
        }

        /*
        * method to animate wrapper which was created by addNextData
        * @param {number} iWrapperID - id of the wrapper that was created
        */
        private static animateWrapper(iWrapperID: number): void {
            let sWrapperTop: string;
            let sWrapperBottom: string;
            if (iWrapperID === 1) {
                sWrapperTop = '#wrapper2';
                sWrapperBottom = '#wrapper1';
            } else {
                sWrapperTop = '#wrapper1';
                sWrapperBottom = '#wrapper2';
            }
            if (Visual.iResponsive) { // if responsive is turned on
                Visual.animateWrapperHelper(sWrapperTop,sWrapperBottom);
            } else { // When responsive is turned OFF
                if (Visual.iAnimationStyle !== 'fade') {
                    if (Visual.iVerticalStack) {
                        if (Visual.iHorizontalScroll) {
                            $(sWrapperTop).animate({
                                left: `-=${Visual.iWidthOfTiles}px`
                            },
                                                   Visual.iDelay).dequeue();
                            // tslint:disable-next-line:typedef
                            $(sWrapperBottom).animate({
                                left: `-=${Visual.iWidthOfTiles}px`
                            },
                                // tslint:disable-next-line:typedef
                                                      Visual.iDelay, ()=> {
                                    Visual.iTimeout = window.setTimeout(()=> {
                                        $(sWrapperTop).remove();
                                        clearTimeout(Visual.iTimeout);
                                    },                                     Visual.iDelay);
                                });
                        } else {
                            $(sWrapperTop).animate({ top: `-=${Visual.iHeightOfTiles * Visual.iNumberOfKPI}px` }
                                ,                  Visual.iDelay).dequeue();

                            // tslint:disable-next-line:typedef
                            $(sWrapperBottom).animate({ top: `-=${Visual.iHeightOfTiles * Visual.iNumberOfKPI}px` },
                                // tslint:disable-next-line:typedef
                                                      Visual.iDelay, ()=> {
                                    Visual.iTimeout = window.setTimeout(()=> {
                                        $(sWrapperTop).remove();
                                        clearTimeout(Visual.iTimeout);
                                    },                                     Visual.iDelay);
                                });
                        }
                    } else {
                        if (Visual.iHorizontalScroll) {
                            $(sWrapperTop).animate({
                                left: `-=${(Visual.iWidthOfTiles * Visual.iNumberOfKPI) + 10}px`
                            },
                                                   Visual.iDelay).dequeue();

                            // tslint:disable-next-line:typedef
                            $(sWrapperBottom).animate({
                                left: `-=${(Visual.iWidthOfTiles * Visual.iNumberOfKPI) + 10}px`
                            },
                                // tslint:disable-next-line:typedef
                                                      Visual.iDelay, ()=> {
                                    Visual.iTimeout = window.setTimeout(()=> {
                                        $(sWrapperTop).remove();
                                        clearTimeout(Visual.iTimeout);
                                    },                                     Visual.iDelay);
                                });
                        } else {
                            $(sWrapperTop).animate({ top: `-=${Visual.iHeightOfTiles}px` }, Visual.iDelay).dequeue();

                            // tslint:disable-next-line:typedef
                            $(sWrapperBottom).animate({ top: `-=${Visual.iHeightOfTiles}px` }, Visual.iDelay, ()=> {
                                Visual.iTimeout = window.setTimeout(()=> {
                                    $(sWrapperTop).remove();
                                    clearTimeout(Visual.iTimeout);
                                },                                     Visual.iDelay);
                            });
                        }
                    }
                } else {
                    Visual.iTimeout = setTimeout(()=> {
                        $(sWrapperTop).remove();
                        clearTimeout(Visual.iTimeout);
                    });
                }
            }

        }

        private static animateWrapperHelper(sWrapperTop: string, sWrapperBottom: string) {
            if (Visual.iAnimationStyle !== 'fade') { // if animationstyle is not fade
                if (Visual.iVerticalStack) {
                    if (Visual.iHorizontalScroll) {
                        $(sWrapperTop).animate({
                            left: `-=${Visual.iMaxDynamicWidthVertical}px`
                        },                                          Visual.iDelay).dequeue();

                        // tslint:disable-next-line:typedef
                        $(sWrapperBottom).animate({
                            left: `-=${Visual.iMaxDynamicWidthVertical}px`
                        },
                            // tslint:disable-next-line:typedef
                                                  Visual.iDelay, ()=> {
                                Visual.iTimeout = window.setTimeout(()=> {
                                    $(sWrapperTop).remove();
                                    clearTimeout(Visual.iTimeout);
                                },                                  Visual.iDelay);
                            });
                    } else {
                        $(sWrapperTop).animate({ top: `-=${Visual.dynamicHeight}px` }, Visual.iDelay).dequeue();

                        // tslint:disable-next-line:typedef
                        $(sWrapperBottom).animate({ top: `-=${Visual.dynamicHeight}px` }, Visual.iDelay, ()=> {
                            Visual.iTimeout = window.setTimeout(()=> {
                                $(sWrapperTop).remove();
                                clearTimeout(Visual.iTimeout);
                            },                                     Visual.iDelay);
                        });
                    }
                } else {
                    if (Visual.iHorizontalScroll) {
                        $(sWrapperTop).animate({
                            left: `-=${Visual.dynamicWidth}px`
                        },                       Visual.iDelay).dequeue();

                        // tslint:disable-next-line:typedef
                        $(sWrapperBottom).animate({
                            left: `-=${Visual.dynamicWidth}px`
                        },
                            // tslint:disable-next-line:typedef
                                                  Visual.iDelay, ()=> {
                                Visual.iTimeout = window.setTimeout(()=> {
                                    $(sWrapperTop).remove();
                                    clearTimeout(Visual.iTimeout);
                                },                                     Visual.iDelay);
                            });
                    } else {
                        $(sWrapperTop).animate({ top: `-=${Visual.iMinHeightOfTilesHorizontal}px` }, Visual.iDelay).dequeue();

                        // tslint:disable-next-line:typedef
                        $(sWrapperBottom).animate({ top: `-=${Visual.iMinHeightOfTilesHorizontal}px` }
                            // tslint:disable-next-line:typedef
                            ,                     Visual.iDelay, ()=> {
                                Visual.iTimeout = window.setTimeout(()=> {
                                    $(sWrapperTop).remove();
                                    clearTimeout(Visual.iTimeout);
                                },                                     Visual.iDelay);
                            });
                    }
                }
            } else { // if animation style is fade
                Visual.iTimeout = setTimeout(()=> {
                    $(sWrapperTop).remove();
                    clearTimeout(Visual.iTimeout);
                });
            }
        }

    }
}
