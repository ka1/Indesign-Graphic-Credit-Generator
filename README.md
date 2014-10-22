Adobe InDesign List of Figures
=================================

Creates a list of figures in the document and appends the formatted list in the document.

Includes the caption number (i.e. "image 15"), the page number (i.e. "page 5) of the image. Reads selected Adobe XMP properties (IPTC) of the graphics:
- copyright notice
- credits
- instructions

Creates paragraph- and character styles for the formatted list, if they do not exist yes.
Creates a new page with a text frame with the script label "creditListTextframe" if that does not exist yes. If it does, all contents in this textframe is overwritten.

Prerequisites:
--------------------
- all image captions need to be assigned to one paragraph style
- all images need to be assigned to one object style
- the image caption needs to be right underneath the image (touching +- 1mm), otherwise the image will not be found
- the image caption must not be offset from the image more than 10mm horizontically

If that all is given, you will get a nice list with all your image captions.

If you want to group images and have assigned a caption to the group, you can write the copyright information into the script label of the group.

Sample file
--------------------
Please refer to the sample file (/sample/sample-CC2014.indd) for further details.
The three JPG files in this folder contain all the IPTC Info necessary for the figure list creation (take a look
at the IPTC contents using Adobe Bridge or your favorite IPTC editor).

In the inDesign file, note the following things:
- all images that should be listed have the object format "Bild Umfluss Bounding Box" attached. This format is selected in the script under "choose object style of image".
- all text frames containing image captions are directly (within a settable tolerance) underneath the image frame. The paragraph style in the sample is called "Bildunterschrift", this can also be selected in the dialog box of the script.
- (this is optional) the text frame containing the list of figures has a script label named "creditListTextframe". If no such frame is found upon script execution, an appropriate text frame is created at the end of the document.

supported image file formats
--------------------
All file formats that are supported by Adobe Indesign which support IPTC headers should be compatible with this script.
Thus far I used it with JPG, PNG, TIFF, PDF and AI