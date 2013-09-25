//List of figure generator, reading XMP Metadata from all Images. Indesign CS6 or later
//(c) 2013 Kai Kasugai
app.scriptPreferences.version = 8; //8 = Indesign CS6 and later. Use 7.5 to use CS5.5 features

var myDocument, timeStart, timeEnd;
myDocument = app.documents.item(0);
timeStart = new Date().getTime();

var xTolerance = 10; //maximum offset of the images to the text frame in X
var yTolerance = 1; //maximum offset of the images to the text frame in Y
var allInfo = new Array(); //Array with all credits
var creditsParagraphStyle; //paragraph style for list of figures
var myCreditsTextFrame; //text frame for credits
var numberOfImages = 0;
var numberOfFoundCredits = 0;
var xmlSettingsTag = 'creditListSettings'; //all user parameters (see below) will be saved in the document structure with this tag

//Settings
var writeParagraphNumber, writePageNumber, includeAuthor, includeCredits, includeInstructions, writeParagraphContents; //booleans
var captionParagraphStyleString = "Bildunterschrift"; //default. will be overwritten if set in xml
var captionedImageObjectStyleString= "Bild Umfluss Bounding Box"; //default. will be overwritten if set in xml
var pageHeaderParagraphStyleString = "Level \\ Level 1"; //paragraph style of the page header
var pageHeaderParagraphStyle; //header paragraph style for list of figures
var captionParagraphStyle; //paragraph style of the caption underneath the image
var captionedImageObjectStyle; //object style of the image that has a caption underneath
var divisionAfterParagraphNumber = "\t"; //separator after the paragraph number
var divisionAfterPageNumber = "\t"; //sepearator after page number
var divisionAfterParagraphContents = "\t"; //separator after the paragraph contents
var authorPrefix = "© "; //prefix before the author
var creditsPrefix = ", "; //prefix before the contents of the credits meta data
var instructionsPrefix = ", "; //prefix before the contents of the instructions meta data
var paragraphContentCharacterLimit = 0; //crop text to x characters

//progress bar
var myProgressPanel;
var myMaximumValue = 100;
var myTextFrameTotalValue = 80;
var myProgressBarWidth = 400;

var langCreditsName = "Bildnachweis"; //header of credit text frame
var langCaption = "Abbildung "; //header of credit text frame
var langPage = "Seite "; //header of credit text frame

//debugging
var cleanUpAllLinks = false; //if true, all links that are created by this script are deleted in the beginning. otherwise, sources and hyperlinks are implicitly deleted by the emptying of the text frames with the list of figures) 

//Select or create paragraph style "Bildnachweis" for all lines in the list of figures
creditsParagraphStyle = returnParagraphStyleOrCreatenew("Bildnachweis");

//show user dialog and then start main program
if (ask() == true) {
	main();
}
else {
	alert("List of figure generation canceled. The document was not changed.");
}

//user dialog
function ask(){
	var myDialog = app.dialogs.add({name:"List of figures", canCancel:true});
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
					var dialogLangCreditsName = textEditboxes.add({editContents: checkOrWriteSetting ("langCreditsName") ? checkOrWriteSetting ("langCreditsName") : langCreditsName, minWidth: 180});
					var dialogLangCaption = textEditboxes.add({editContents: checkOrWriteSetting ("langCaption") ? checkOrWriteSetting ("langCaption") : langCaption, minWidth: 180});
					var dialogLangPage = textEditboxes.add({editContents: checkOrWriteSetting ("langPage") ? checkOrWriteSetting ("langPage") : langPage, minWidth: 180});
				}
			}
			with(borderPanels.add()){
				with(dialogColumns.add()){
					var dialogWriteParagraphNumber = checkboxControls.add({staticLabel:"Write paragraph number", checkedState: checkOrWriteSetting ("writeParagraphNumber") == "no" ? false : true});
					var dialogWritePageNumber = checkboxControls.add({staticLabel:"Write page number", checkedState: checkOrWriteSetting ("writePageNumber") == "no" ? false : true});
					var dialogWriteParagraphContents = checkboxControls.add({staticLabel:"Write paragraph contents", checkedState: checkOrWriteSetting ("writeParagraphContents") == "yes" ? true : false});
					var dialogIncludeAuthor= checkboxControls.add({staticLabel:"Include author metadata", checkedState: checkOrWriteSetting ("includeAuthor") == "no" ? false : true});
					var dialogIncludeCredits = checkboxControls.add({staticLabel:"Include credits metadata", checkedState: checkOrWriteSetting ("includeCredits") == "no" ? false : true});
					var dialogIncludeInstructions = checkboxControls.add({staticLabel:"Include instructions metadata", checkedState: checkOrWriteSetting ("includeInstructions") == "no" ? false : true});
				}
				with(dialogColumns.add()){
					staticTexts.add({staticLabel:"Division after paragraph number"});
					staticTexts.add({staticLabel:"Division after page number"});
					staticTexts.add({staticLabel:"Division after paragraph contents"});
					staticTexts.add({staticLabel:"Author prefix"});
					staticTexts.add({staticLabel:"credits prefix"});
					staticTexts.add({staticLabel:"instructions prefix"});
				}
				with(dialogColumns.add()){
					var dialogDivisionAfterParagraphNumber = textEditboxes.add({editContents: checkOrWriteSetting ("divisionAfterParagraphNumber") ? checkOrWriteSetting ("divisionAfterParagraphNumber") : divisionAfterParagraphNumber, minWidth: 100});
					var dialogDivisionAfterPageNumber = textEditboxes.add({editContents: checkOrWriteSetting ("divisionAfterPageNumber") ? checkOrWriteSetting ("divisionAfterPageNumber") : divisionAfterPageNumber, minWidth: 100});
					var dialogDivisionAfterParagraphContents = textEditboxes.add({editContents: checkOrWriteSetting ("divisionAfterParagraphContents") ? checkOrWriteSetting ("divisionAfterParagraphContents") : divisionAfterParagraphContents, minWidth: 100});
					var dialogAuthorPrefix = textEditboxes.add({editContents: checkOrWriteSetting ("authorPrefix") ? checkOrWriteSetting ("authorPrefix") : authorPrefix, minWidth: 100});
					var dialogCreditsPrefix = textEditboxes.add({editContents: checkOrWriteSetting ("creditsPrefix") ? checkOrWriteSetting ("creditsPrefix") : creditsPrefix, minWidth: 100});
					var dialogInstructionsPrefix = textEditboxes.add({editContents: checkOrWriteSetting ("instructionsPrefix") ? checkOrWriteSetting ("instructionsPrefix") : instructionsPrefix, minWidth: 100});
				}
				with(dialogColumns.add()){
					staticTexts.add({staticLabel:""});
					staticTexts.add({staticLabel:""});
					staticTexts.add({staticLabel:"Character Limit"});
				}
				with(dialogColumns.add()){
					staticTexts.add({staticLabel:""});
					staticTexts.add({staticLabel:""});
					var dialogParagraphContentCharacterLimit = textEditboxes.add({editContents: checkOrWriteSetting ("paragraphContentCharacterLimit") ? checkOrWriteSetting ("paragraphContentCharacterLimit") : paragraphContentCharacterLimit.toString(), minWidth: 50});
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
	checkOrWriteSetting ("creditsPrefix", dialogCreditsPrefix.editContents);
	checkOrWriteSetting ("instructionsPrefix", dialogInstructionsPrefix.editContents);
	checkOrWriteSetting ("divisionAfterParagraphNumber", dialogDivisionAfterParagraphNumber.editContents);
	checkOrWriteSetting ("divisionAfterPageNumber", dialogDivisionAfterPageNumber.editContents);
	checkOrWriteSetting ("divisionAfterParagraphContents", dialogDivisionAfterParagraphContents.editContents);
	checkOrWriteSetting ("langCreditsName", dialogLangCreditsName.editContents);
	checkOrWriteSetting ("langCaption", dialogLangCaption.editContents);
	checkOrWriteSetting ("langPage", dialogLangPage.editContents);
	checkOrWriteSetting ("writeParagraphNumber", dialogWriteParagraphNumber.checkedState == false ? "no" : "yes");
	checkOrWriteSetting ("writePageNumber", dialogWritePageNumber.checkedState == false ? "no" : "yes");
	checkOrWriteSetting ("writeParagraphContents", dialogWriteParagraphContents.checkedState == false ? "no" : "yes");
	checkOrWriteSetting ("includeAuthor", dialogIncludeAuthor.checkedState == false ? "no" : "yes");
	checkOrWriteSetting ("includeCredits", dialogIncludeCredits.checkedState == false ? "no" : "yes");
	checkOrWriteSetting ("includeInstructions", dialogIncludeInstructions.checkedState == false ? "no" : "yes");
	checkOrWriteSetting ("paragraphContentCharacterLimit", dialogParagraphContentCharacterLimit.editContents);

	//set for runtime
	captionParagraphStyle = parseParagraphStyleString(myCaptionParagraphStyleString);
	captionedImageObjectStyle = parseObjectStyleString(myCaptionedImageObjectStyleString);
	pageHeaderParagraphStyle = parseParagraphStyleString(myPageHeaderParagraphStyleString);
	authorPrefix = dialogAuthorPrefix.editContents;
	creditsPrefix = dialogCreditsPrefix.editContents;
	instructionsPrefix = dialogInstructionsPrefix.editContents;
	divisionAfterParagraphNumber = dialogDivisionAfterParagraphNumber.editContents;
	divisionAfterPageNumber = dialogDivisionAfterPageNumber.editContents;
	divisionAfterParagraphContents = dialogDivisionAfterParagraphContents.editContents;
	langCreditsName = dialogLangCreditsName.editContents;
	langCaption = dialogLangCaption.editContents;
	langPage = dialogLangPage.editContents;
	writeParagraphNumber = dialogWriteParagraphNumber.checkedState;
	writePageNumber = dialogWritePageNumber.checkedState;
	writeParagraphContents = dialogWriteParagraphContents.checkedState;
	includeAuthor = dialogIncludeAuthor.checkedState;
	includeCredits = dialogIncludeCredits.checkedState;
	includeInstructions = dialogIncludeInstructions.checkedState;
	paragraphContentCharacterLimit = parseInt(dialogParagraphContentCharacterLimit.editContents);

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
	
	//Init progress bar
	myCreateProgressPanel(myMaximumValue, myProgressBarWidth);
	myProgressPanel.show();
	myProgressPanel.myProgressBar.value = 0;
	
	//CLEAN UP
	//delete all hyperlink destinations
	if (myDocument.hyperlinkTextDestinations.length > 0){
		myProgressPanel.myText.text = "Cleaning up hyperlink destinations";
		for (var i = myDocument.hyperlinkTextDestinations.length - 1; i >= 0; i--){
			if (myDocument.hyperlinkTextDestinations[i].label == 'lofLinkDest' || myDocument.hyperlinkTextDestinations[i].name.match(/figureRef-[0-9]+/i)) {
				myDocument.hyperlinkTextDestinations[i].remove();
			}
		}
	}

	if (cleanUpAllLinks){
		//delete existing hyperlink sources (this should be unnecessary, as sources should be deleted when emptying the text frames with the list of figures which contain the hyperlink sources)
		if (myDocument.hyperlinkTextSources.length > 0){
			myProgressPanel.myText.text = "Cleaning up hyperlink text sources";
			for(var i = myDocument.hyperlinkTextSources.length -1; i >= 0; i--){
				if (myDocument.hyperlinkTextSources[i].label == 'lofLinkSrc'){
					myDocument.hyperlinkTextSources[i].remove();
				}
			}
		}

		//delete all hyperlinks
		if (myDocument.hyperlinks.length > 0){
			myProgressPanel.myText.text = "Cleaning up hyperlinks";
			for (var i = myDocument.hyperlinks.length - 1; i >= 0; i--){
				if (myDocument.hyperlinks[i].label == 'lofLinkHyperlink' || myDocument.hyperlinks[i].name.match(/listOfFigures[0-9]+/i)){
					myDocument.hyperlinks[i].remove();
				}
			}
		}
	}
	

	myProgressPanel.myText.text = "Parsing textframes";
	//parse ALL TEXTFRAMES
	var totalNumberOfTextFrames = myDocument.textFrames.count();
	for (var i = 0; i < totalNumberOfTextFrames; i++){

		//progress bar update
		myProgressPanel.myProgressBar.value += myTextFrameTotalValue / totalNumberOfTextFrames;
		myProgressPanel.myText.text = "Parsing textframe " + i + " of " + totalNumberOfTextFrames;

		//parse ALL PARAGRAPHS in CURRENT TEXTFRAME
		var currentTextFrame = myDocument.textFrames[i];
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
								if (includeCredits) theInfo += currentGraphic.linkXmp.getProperty("http://ns.adobe.com/photoshop/1.0/","photoshop:Credit") ? (theInfo ? creditsPrefix : "") + currentGraphic.linkXmp.getProperty("http://ns.adobe.com/photoshop/1.0/","photoshop:Credit") : "";
								if (includeInstructions) theInfo += currentGraphic.linkXmp.getProperty("http://ns.adobe.com/photoshop/1.0/","photoshop:Instructions") ? (theInfo ? instructionsPrefix : "") + currentGraphic.linkXmp.getProperty("http://ns.adobe.com/photoshop/1.0/","photoshop:Instructions") : "";
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
			
				//create a hyperlink text destination and safe when pushing the info to the array
				var theHyperlinkDestination = myDocument.hyperlinkTextDestinations.add(currentParagraph.insertionPoints[-1],{name:"figureRef-" + currentParagraph.numberingResultNumber, label: 'lofLinkDest'});

				//safe content string and trim if necessary
				var contentString;
				if (writeParagraphContents) {
					//if the length should be limited and the length exceeds the limit
					if (paragraphContentCharacterLimit > 0 && currentParagraph.contents.length > paragraphContentCharacterLimit){
						contentString = currentParagraph.contents.substr(0,paragraphContentCharacterLimit) + "...";
					} else {
						contentString = currentParagraph.contents;
					}
				}
				
				//construct entire line
				var theText = (writeParagraphNumber ? langCaption + currentParagraph.numberingResultNumber + divisionAfterParagraphNumber : "") +
					(writePageNumber ? langPage + theParentPage.name + divisionAfterPageNumber : "") +
					(writeParagraphContents ? contentString + divisionAfterParagraphContents : "") +
					(found ? theInfo : "NO IMAGE FOUND");
				allInfo.push({textContents: theText, thisHyperlinkDestination: theHyperlinkDestination});
			}
		}
	}

	//progress bar update
	myProgressPanel.myProgressBar.value = myTextFrameTotalValue + 1;
	myProgressPanel.myText.text = "Sorting numbers";

	//sort entries
	allInfo.sort(sortCredits);

	//progress bar update
	myProgressPanel.myProgressBar.value = myTextFrameTotalValue + 3;
	myProgressPanel.myText.text = "Finding or creating text frame for list of figures";

	//CREATE LIST OF FIGURES TEXT FRAME
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

	//progress bar update
	myProgressPanel.myProgressBar.value = myTextFrameTotalValue + 5;
	myProgressPanel.myText.text = "Write content to text frame";

	//enter content into frame
	for(var i = 0; i < allInfo.length; i++){
		
		//progress bar update
		myProgressPanel.myText.text = "Writing info " + (i+1) + " / " + allInfo.length;
		myProgressPanel.myProgressBar.value += 15 / allInfo.length;

		addFormattedTextToStory(myCreditsTextFrame,false, "\r",false);
		var newLine = addFormattedTextToStory(myCreditsTextFrame,false, allInfo[i].textContents,creditsParagraphStyle);

		var myReferenceSource = myDocument.hyperlinkTextSources.add(newLine,{name:"lofLinkSrc_" + i, label: "lofLinkSrc"});
		//if the following line causes an error, FOR SOME REASON it helped, to change the name (like from "listOfFigures-" to "listOfFigures_")
		myDocument.hyperlinks.add(myReferenceSource,allInfo[i].thisHyperlinkDestination,{name: "listOfFigures" + i, label:"lofLinkHyperlink"});
	}

	//add a final line break to avoid hyperlink errors in zotero reference importer (if all image credits end with a citation and there was no linebreak, the script would report an error upon the first run)
	myCreditsTextFrame.parentStory.contents += "\r";

	//final progress bar update and hide
	myProgressPanel.myProgressBar.value = myMaximumValue;
	myProgressPanel.myText.text = "DONE";
	myProgressPanel.hide();
	
	//TIMEs
	timeEnd = new Date().getTime();
	alert(
		"Done\nImages: " + numberOfImages +
		"\nCredits found: " + numberOfFoundCredits +
		"\n\nin " + Math.ceil((timeEnd - timeStart) / 1000.0) + " seconds"
	);
}

//-----------------------------
//-------FUNCTIONS-------
//-----------------------------

function sortCredits(x,y){
	var zahlX = x.textContents.match(/^[^0-9]*([0-9]+)/);
	zahlX = parseInt(zahlX[1]);
	var zahlY = y.textContents.match(/^[^0-9]*([0-9]+)/);
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
	return myAdditions;
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

//helper to create progress panel
function myCreateProgressPanel(myMaximumValue, myProgressBarWidth){
	var test;
	myProgressPanel = new Window('window', 'creating list of figures');
	with(myProgressPanel){
		test = myProgressPanel.myProgressBar = add('progressbar', [12, 12, myProgressBarWidth, 24], 0, myMaximumValue);
		myProgressPanel.myText = add('statictext', {x:60, y:0, width:myProgressBarWidth, height:20});
		myText.text = "Starting";
	}
}