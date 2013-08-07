//Credit List Generator, reading XMP Metadata from all Images. Indesign CS6 or later
//(c) 2013 Kai Kasugai
app.scriptPreferences.version = 8; //8 = Indesign CS6 and later. Use 7.5 to use CS5.5 features

var myDocument, timeStart, timeEnd;
myDocument = app.documents.item(0);
timeStart = new Date().getTime();

var xTolerance = 10; //maximum offset of the images to the text frame in X
var yTolerance = 1; //maximum offset of the images to the text frame in Y
var allInfo = new Array(); //Array with all credits
var creditsParagraphStyle; //paragraph style for credit list
var myCreditsTextFrame; //text frame for credits
var numberOfImages = 0;
var numberOfFoundCredits = 0;
var xmlSettingsTag = 'creditListSettings';

//Settings
var writeParagraphNumber, writePageNumber, includeAuthor, includeCredits, includeInstructions;
var captionParagraphStyleString = "Bildunterschrift"; //default. will be overwritten if set in xml
var captionedImageObjectStyleString= "Bild Umfluss Bounding Box"; //default. will be overwritten if set in xml
var pageHeaderParagraphStyleString = "Level \\ Level 1";
var pageHeaderParagraphStyle; //header paragraph style for credit list
var captionParagraphStyle;
var captionedImageObjectStyle;
var authorPrefix = "© ";
var langCreditsName = "Bildnachweis"; //header of credit text frame
var langCaption = "Abbildung "; //header of credit text frame
var langPage = "Seite "; //header of credit text frame

creditsParagraphStyle = returnParagraphStyleOrCreatenew("Bildnachweis");

if (ask() == true) {
    main();
}
else {
    alert("Credit list generation canceled. The document was not changed.");
}

function ask(){
    
    var myDialog = app.dialogs.add({name:"Credit List", canCancel:true});
    with(myDialog){
        //Add a dialog column.
        with(dialogColumns.add()){
            with(borderPanels.add()){
                with(dialogColumns.add()){
                    //paragraph style
                    staticTexts.add({staticLabel:"Choose Image Caption Paragraph Style"});
                    //object style
                    staticTexts.add({staticLabel:"Choose Object Style of Image"});
                    //paragraph style of header
                    staticTexts.add({staticLabel:"Set Page Header Paragraph Style"});
                    //text
                    staticTexts.add({staticLabel:"Author Prefix"});
                    staticTexts.add({staticLabel:"Page Header"});
                    staticTexts.add({staticLabel:"Caption Number Prefix"});
                    staticTexts.add({staticLabel:"Page Number Prefix"});
                }
                with(dialogColumns.add()){
                    
                    var dialogCaptionParagraphStyle = dropdowns.add(createDialogDropDownParagraphs(captionParagraphStyleString, "captionParagraphStyleString"));
                    var dialogCaptionedImageObjectStyle = dropdowns.add(createDialogDropDownObjectStyles(captionedImageObjectStyleString, "captionedImageObjectStyleString"));
                    var dialogPageHeaderParagraphStyle = dropdowns.add(createDialogDropDownParagraphs(pageHeaderParagraphStyleString, "pageHeaderParagraphStyleString"));
                    
                    //---------------------------------------------
                    //---------------------texts------------------
                    //---------------------------------------------
                    var dialogAuthorPrefix = textEditboxes.add({editContents: checkOrWriteSetting ("authorPrefix") ? checkOrWriteSetting ("authorPrefix") : authorPrefix, minWidth: 180});
                    var dialogLangCreditsName = textEditboxes.add({editContents: checkOrWriteSetting ("langCreditsName") ? checkOrWriteSetting ("langCreditsName") : langCreditsName, minWidth: 180});
                    var dialogLangCaption = textEditboxes.add({editContents: checkOrWriteSetting ("langCaption") ? checkOrWriteSetting ("langCaption") : langCaption, minWidth: 180});
                    var dialogLangPage = textEditboxes.add({editContents: checkOrWriteSetting ("langPage") ? checkOrWriteSetting ("langPage") : langPage, minWidth: 180});
                    
                }
            }
            with(borderPanels.add()){
                with(dialogColumns.add()){
                        var dialogWriteParagraphNumber = checkboxControls.add({staticLabel:"Write paragraph number", checkedState: checkOrWriteSetting ("writeParagraphNumber") == "no" ? false : true});
                        var dialogWritePageNumber = checkboxControls.add({staticLabel:"Write page number", checkedState: checkOrWriteSetting ("writePageNumber") == "no" ? false : true});
                        var dialogIncludeAuthor= checkboxControls.add({staticLabel:"Include author metadata", checkedState: checkOrWriteSetting ("includeAuthor") == "no" ? false : true});
                        var dialogIncludeCredits = checkboxControls.add({staticLabel:"Include credits metadata", checkedState: checkOrWriteSetting ("includeCredits") == "no" ? false : true});
                        var dialogIncludeInstructions = checkboxControls.add({staticLabel:"Include instructions metadata", checkedState: checkOrWriteSetting ("includeInstructions") == "no" ? false : true});
                }
            }
        }
    }
    //Display the dialog box.
    if(myDialog.show() == true){
        //Get the value from the dialog box
        var myCaptionParagraphStyleString = dialogCaptionParagraphStyle.stringList[dialogCaptionParagraphStyle.selectedIndex];
        var myCaptionedImageObjectStyleString = dialogCaptionedImageObjectStyle.stringList[dialogCaptionedImageObjectStyle.selectedIndex];
        var myPageHeaderParagraphStyleString = dialogPageHeaderParagraphStyle.stringList[dialogPageHeaderParagraphStyle.selectedIndex];
        
        //write to document xml
        checkOrWriteSetting ("captionParagraphStyleString", myCaptionParagraphStyleString);
        checkOrWriteSetting ("captionedImageObjectStyleString", myCaptionedImageObjectStyleString);
        checkOrWriteSetting ("pageHeaderParagraphStyleString", myPageHeaderParagraphStyleString);
        checkOrWriteSetting ("authorPrefix", dialogAuthorPrefix.editContents);
        checkOrWriteSetting ("langCreditsName", dialogLangCreditsName.editContents);
        checkOrWriteSetting ("langCaption", dialogLangCaption.editContents);
        checkOrWriteSetting ("langPage", dialogLangPage.editContents);
        checkOrWriteSetting ("writeParagraphNumber", dialogWriteParagraphNumber.checkedState == false ? "no" : "yes");
        checkOrWriteSetting ("writePageNumber", dialogWritePageNumber.checkedState == false ? "no" : "yes");
        checkOrWriteSetting ("includeAuthor", dialogIncludeAuthor.checkedState == false ? "no" : "yes");
        checkOrWriteSetting ("includeCredits", dialogIncludeCredits.checkedState == false ? "no" : "yes");
        checkOrWriteSetting ("includeInstructions", dialogIncludeInstructions.checkedState == false ? "no" : "yes");

        //set for runtime
        captionParagraphStyle = parseParagraphStyleString(myCaptionParagraphStyleString);
        captionedImageObjectStyle = parseObjectStyleString(myCaptionedImageObjectStyleString);
        pageHeaderParagraphStyle = parseParagraphStyleString(myPageHeaderParagraphStyleString);
        authorPrefix = dialogAuthorPrefix.editContents;
        langCreditsName = dialogLangCreditsName.editContents;
        langCaption = dialogLangCaption.editContents;
        langPage = dialogLangPage.editContents;
        writeParagraphNumber = dialogWriteParagraphNumber.checkedState;
        writePageNumber = dialogWritePageNumber.checkedState;
        includeAuthor = dialogIncludeAuthor.checkedState;
        includeCredits = dialogIncludeCredits.checkedState;
        includeInstructions = dialogIncludeInstructions.checkedState;

        //clear dialog from memory
        myDialog.destroy();

    }
    else{
        myDialog.destroy();
        return false;
    }

    return true;
}


function main(){
    //parse ALL TEXTFRAMES
    for (var i = 0; i < myDocument.textFrames.count(); i++){
        var currentTextFrame = myDocument.textFrames[i];
        
        //parse ALL PARAGRAPHS in CURRENT TEXTFRAME
        for (var p = 0; p < currentTextFrame.paragraphs.count(); p++){
            var currentParagraph = currentTextFrame.paragraphs[p];
            //Found a Paragraph with a caption?
            if (currentParagraph.appliedParagraphStyle == captionParagraphStyle){
                numberOfImages++;
                //Textframe, current Page and top Y Position to compare to images Y Positions on page
                var theParentTextFrame = currentParagraph.parentTextFrames[0];
                var theParentPage = theParentTextFrame.parentPage;
                var topY = theParentTextFrame.geometricBounds[0];
                var leftX = theParentTextFrame.geometricBounds[1];
                var found = false;
                var currentRectangle;
                var currentGraphic;
                var theInfo;
                
                //parse ALL RECTANGLES on CURRENT PAGE
                for (var r = 0; r < theParentPage.rectangles.count(); r++){
                    //see if the object style is correct and the X and Y Positions match within set tolerance
                    currentRectangle = theParentPage.rectangles[r];
                    if (currentRectangle.appliedObjectStyle == captionedImageObjectStyle) {
                        if (currentRectangle.geometricBounds[2] > topY - yTolerance && currentRectangle.geometricBounds[2] < topY + yTolerance) {
                            if (currentRectangle.geometricBounds[1] > leftX - xTolerance && currentRectangle.geometricBounds[1] < leftX + xTolerance) {
                                found = true;
                                
                                currentGraphic = currentRectangle.graphics[0].itemLink;
                                theInfo = "";
                                
                                //Copyright
                                try{
                                    if (includeAuthor)  theInfo += currentGraphic.linkXmp.copyrightNotice ? authorPrefix + toTitleCase(currentGraphic.linkXmp.copyrightNotice) : "";
                                    if (includeCredits) theInfo += currentGraphic.linkXmp.getProperty("http://ns.adobe.com/photoshop/1.0/","photoshop:Credit") ? (theInfo ? ", " : "") + currentGraphic.linkXmp.getProperty("http://ns.adobe.com/photoshop/1.0/","photoshop:Credit") : "";
                                    if (includeInstructions) theInfo += currentGraphic.linkXmp.getProperty("http://ns.adobe.com/photoshop/1.0/","photoshop:Instructions") ? (theInfo ? ", " : "") + currentGraphic.linkXmp.getProperty("http://ns.adobe.com/photoshop/1.0/","photoshop:Instructions") : "";
                                    numberOfFoundCredits++;
                                }
                                catch(e) {
                                    theInfo += "ERROR RETREIVING METADATA FOR " + currentGraphic.name;
                                }
                                
                                break;
                            }
                        }
                    }
                }
            
                //search GROUPS if nothing has been found
                if (!found){
                    //parse ALL GROUPS on CURRENT PAGE
                    for (var r = 0; r < theParentPage.groups.count(); r++){
                        //see if the object style is correct and the X and Y Positions match within set tolerance
                        currentGroup = theParentPage.groups[r];
                        if (currentGroup.appliedObjectStyle == captionedImageObjectStyle) {
                            if (currentGroup.geometricBounds[2] > topY - yTolerance && currentGroup.geometricBounds[2] < topY + yTolerance) {
                                if (currentGroup.geometricBounds[1] > leftX - xTolerance && currentGroup.geometricBounds[1] < leftX + xTolerance) {
                                    found = true;
                                    theInfo = currentGroup.label;
                                    
                                    //count, if not empty/null
                                    if (theInfo != null & theInfo != ''){
                                        numberOfFoundCredits++;
                                    }
                                    
                                    break;
                                }
                            }
                        }
                    }
                }
                allInfo.push((writeParagraphNumber ? langCaption + currentParagraph.numberingResultNumber + "\t" : "") + (writePageNumber ? langPage + theParentPage.name + "\t" : "") + (found ? theInfo : "NO IMAGE FOUND") + "");
            }
        }
    }

    allInfo.sort(sortCredits);

    //CREATE CREDIT LIST TEXT FRAME
    //find reference text frame
    for (var i = 0; i < myDocument.pages.count(); i++){
        for (var j = 0; j < myDocument.pages[i].textFrames.count(); j++){
            if (myDocument.pages[i].textFrames[j].label == 'creditListTextframe'){
                myCreditsTextFrame = myDocument.pages[i].textFrames[j];
                myCreditsTextFrame.parentStory.contents = "";
                addFormattedTextToStory(myCreditsTextFrame,false,(langCreditsName == '' ? " " : langCreditsName),pageHeaderParagraphStyle);
            }
        }
    }

    //create new page with textframe for references if none was found
    if (!myCreditsTextFrame){
        var newPage = myDocument.pages.add();
        myCreditsTextFrame = newPage.textFrames.add();
        myCreditsTextFrame.label = 'creditListTextframe';
        if (newPage.side == PageSideOptions.RIGHT_HAND){
            myCreditsTextFrame.geometricBounds = [newPage.marginPreferences.top,newPage.bounds[1] + newPage.marginPreferences.left,newPage.bounds[2] - newPage.marginPreferences.bottom,newPage.bounds[3] - newPage.marginPreferences.right];
        } else {
            myCreditsTextFrame.geometricBounds = [newPage.marginPreferences.top,newPage.bounds[1] + newPage.marginPreferences.right,newPage.bounds[2] - newPage.marginPreferences.bottom,newPage.bounds[3] - newPage.marginPreferences.left];
        }
        addFormattedTextToStory(myCreditsTextFrame,false,(langCreditsName == '' ? " " : langCreditsName),pageHeaderParagraphStyle);
        alert("Page " + (newPage.documentOffset+1) + " was created with the textframe for the credits.\n\nIf you want to define your own credits textframe, please create a textframe with the script-label \"creditListTextframe\" (case sensitive!).\n\nThis textframe (and the parent story) will be emptied and filled with references.");
    }

    //enter content into frame
    for(var i = 0; i < allInfo.length; i++){
        addFormattedTextToStory(myCreditsTextFrame,false, "\r",false);
        addFormattedTextToStory(myCreditsTextFrame,false, allInfo[i],creditsParagraphStyle);
        //myCreditsTextFrame.parentStory.insertionPoints.item(-1).contents = SpecialCharacters.COPYRIGHT_SYMBOL;
    }


    //TIMEs
    timeEnd = new Date().getTime();
    alert(
        "Done\nImages: " + numberOfImages +
        "\nCredits found: " + numberOfFoundCredits +
        "\n\nin " + Math.ceil((timeEnd - timeStart) / 1000.0) + " seconds"
    );
}

//----------------------------
//-------FUNCTIONS-------
//----------------------------

function sortCredits(x,y){
    var zahlX = x.match(/^[^0-9]*([0-9]+)/);
    zahlX = parseInt(zahlX[1]);
    var zahlY = y.match(/^[^0-9]*([0-9]+)/);
    zahlY = parseInt(zahlY[1]);
    
    return zahlX > zahlY;
}

function addFormattedTextToStory(myTextframe,myFormat,myContent,myParagraphFormat){
    if (!myContent) return false;
    //safe insertion point index
    var firstInsertionPoint = myTextframe.parentStory.insertionPoints[-1].index;
    //add text
    myTextframe.parentStory.insertionPoints[-1].contents += myContent;
    var myAdditions = myTextframe.parentStory.characters.itemByRange(myTextframe.parentStory.insertionPoints[firstInsertionPoint],myTextframe.parentStory.insertionPoints[-1]);

    //add formatting
    if (myFormat) {
        myAdditions.appliedCharacterStyle = myFormat;
    }
    if (myParagraphFormat) {
        myAdditions.appliedParagraphStyle = myParagraphFormat;
    }
    return true;
}

//return the reference to a paragraph style. if the style did not exist, create the style
function returnParagraphStyleOrCreatenew(stylename, groupname, stylePreferences){
    var style;
    var group;

    //prepare style preferences. if none are given, only include the name
    if (stylePreferences == null){
        stylePreferences = {name: stylename};
    } else {
        stylePreferences.name = stylename;
    }

    //is the style in a group?
    //NO GROUP
    if (!groupname){
        style = myDocument.paragraphStyles.item(stylename);
    } 
    //GROUP
    else {
        try { //add group first, if it does not exist
            group = myDocument.paragraphStyleGroups.itemByName(groupname);
            gname = group.name; //will trigger error if group does not exist
        }
        catch(e){
            group = myDocument.paragraphStyleGroups.add({name: groupname});
            style = group.paragraphStyles.add(stylePreferences); //then add style in the group
        }
        style = group.paragraphStyles.itemByName(stylename); //select style in group (for the second time, if the style had already been created, but that should be ok)
    }

    try{
        name=style.name; //will trigger error if style does not exist
    } catch(e) {
        if (group != null){
            style = group.paragraphStyles.add(stylePreferences);
        } else {
            style = myDocument.paragraphStyles.add(stylePreferences);
        }
    }
    return style;    
}

function toTitleCase(str)
{
    return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}

//to store user settings in document
function checkOrWriteSetting(settingName, newSetting){
    //check for settings root first
    var mySettingsRoot = myDocument.xmlElements[0].xmlElements.itemByName(xmlSettingsTag)
    if (!mySettingsRoot || !mySettingsRoot.isValid){
        //if there is no settings entry in the xml tree yet and no settings are to be safed, return false because no setting can be read
        if (newSetting == null) {
            return false;
        }
        mySettingsRoot = myDocument.xmlElements[0].xmlElements.add(xmlSettingsTag);
        mySettingsRoot = mySettingsRoot.move(LocationOptions.atBeginning)
    }    
    
    //check for the attribute defaultDirectory
    var defDir  = mySettingsRoot.xmlAttributes.itemByName(settingName);
    //if there is no default directory set
    if (!defDir.isValid){
        //if no default directory was set
        if (newSetting == null){
            return false;
        }
        //create new attribute or write to attribute
        mySettingsRoot.xmlAttributes.add(settingName,newSetting);
    }
    //if the value exists, but is empty
    else if (newSetting != null || defDir.value.toString() == ''){
        if (newSetting == null){
            return false;
        }
        mySettingsRoot.xmlAttributes.itemByName(settingName).value = newSetting;        
    }
    //attribute exists, is not empty, return the attributes value
    else {
        return defDir.value;
    }
}

//creates the drop down object with a list of paragraphs styles including paragraph styles in groups for a dialog drop down
function createDialogDropDownParagraphs(thisParagraphStyleStringVar, thisParagraphStyleString){
    //if there is a safed value for caption paragraph style and image object style, overwrite default value
    if (checkOrWriteSetting(thisParagraphStyleString)) {
        thisParagraphStyleStringVar = checkOrWriteSetting(thisParagraphStyleString);
    }

    var pSelectedIndex = 0; //paragraph index
    var pCurrentIndex = -1; //paragraph index

    //prepare array with all styles
    var thisStyleArray = new Array();
    
    //parse all non-grouped paragraph styles
    for (i = 0; i < myDocument.paragraphStyles.count(); i++) {
        pCurrentIndex++;
        //compare with default string
        if (myDocument.paragraphStyles.item(i).name == thisParagraphStyleStringVar) {
            pSelectedIndex = pCurrentIndex;
        }
        thisStyleArray.push(myDocument.paragraphStyles.item(i).name);
    }
    //parse all grouped paragraph styles
    for (i = 0; i < myDocument.paragraphStyleGroups.count(); i++){
        var currentGroup = myDocument.paragraphStyleGroups[i];
        for (j = 0; j < currentGroup.paragraphStyles.count(); j++) {
            pCurrentIndex++;
            var currentParagraphStyleString = currentGroup.name + " \\ " + currentGroup.paragraphStyles.item(j).name;
            if (currentParagraphStyleString == thisParagraphStyleStringVar) {
                pSelectedIndex = pCurrentIndex;
            }
            thisStyleArray.push(currentParagraphStyleString);
        }
     }
 
    //return drop down object
    return {stringList:thisStyleArray, selectedIndex:pSelectedIndex};
}

//creates the drop down object with a list of object styles including object styles in groups for a dialog drop down
function createDialogDropDownObjectStyles(thisObjectStyleStringVar, thisObjectStyleString){
    //if there is a safed value for caption paragraph style and image object style, overwrite default value
    if (checkOrWriteSetting(thisObjectStyleString)) {
        thisObjectStyleStringVar = checkOrWriteSetting(thisObjectStyleString);
    }

    var oSelectedIndex = 0; //object style index
    var oCurrentIndex = -1; //object style index

    //prepare array with all styles
    var thisStyleArray = new Array();
    
    //parse all non-grouped object styles
    for (i = 0; i < myDocument.objectStyles.count(); i++) {
        oCurrentIndex++;
        //compare with default string
        if (myDocument.objectStyles.item(i).name == thisObjectStyleStringVar) {
            oSelectedIndex = oCurrentIndex;
        }
        thisStyleArray.push(myDocument.objectStyles.item(i).name);
    }
    //parse all grouped object styles
    for (i = 0; i < myDocument.objectStyleGroups.count(); i++){
        var currentGroup = myDocument.objectStyleGroups[i];
        for (j = 0; j < currentGroup.objectStyles.count(); j++) {
            oCurrentIndex++;
            var currentObjectStyleString = currentGroup.name + " \\ " + currentGroup.objectStyles.item(j).name;
            if (currentObjectStyleString == thisObjectStyleStringVar) {
                oSelectedIndex = oCurrentIndex;
            }
            thisStyleArray.push(currentObjectStyleString);
        }
     }
 
    //return drop down object
    return {stringList:thisStyleArray, selectedIndex:oSelectedIndex};
}

//parses a string like "GROUPNAME \ STYLENAME" and returns the style object. Also works if no Groupname (and no division) is given.
function parseParagraphStyleString(pStyleString){
    var groupRegex = pStyleString.match(/([^\\]*) \\ ([^\\]*)/);
    var group;
    var style;
    //keine Gruppe erkannt
    if (groupRegex == null){
        style = myDocument.paragraphStyles.item(pStyleString);
        //return FALSE if style is invalid. otherwise return style
    }
    //Gruppe erkannt
    else {
        group = myDocument.paragraphStyleGroups.item(groupRegex[1]);
        style = group.paragraphStyles.item(groupRegex[2]);
        //return FALSE if style is invalid. otherwise return style
    }

    try{
        var name = style.name;
    }
    catch(e){
        return false;
    }
    return style;

}

//parses a string like "GROUPNAME \ STYLENAME" and returns the style object. Also works if no Groupname (and no division) is given.
function parseObjectStyleString(oStyleString){
    var groupRegex = oStyleString.match(/([^\\]*) \\ ([^\\]*)/);
    var group;
    var style;
    //keine Gruppe erkannt
    if (groupRegex == null){
        style = myDocument.objectStyles.item(oStyleString);
        //return FALSE if style is invalid. otherwise return style
    }
    //Gruppe erkannt
    else {
        group = myDocument.objectStyleGroups.item(groupRegex[1]);
        style = group.objectStyles.item(groupRegex[2]);
    }

    try{
        var name = style.name;
    }
    catch(e){
        return false;
    }
    return style;
}
