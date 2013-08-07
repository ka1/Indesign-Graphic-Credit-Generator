Adobe InDesign List of Figures
=================================

Creates a list of figures in the document and appends the formatted list in the document.

Includes the caption number (i.e. "image 15"), the page number (i.e. "page 5) of the image. Reads selected Adobe XMP properties of the graphics:
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
