Image2TEX by Borde
==================

This program can convert BMP, JPG, GIF, ICO, WMF and EMF images
into TEX textures compatibles with Final Fantasy 7 (PC version).
It can also save the loaded images as BMPs.
The textures keep the same color depth they originally had.
This tool is heavily untested, use it at your own risk.

Usage:
------
1-Open an image by using the "Open image" button
2-Save the image as a tex (or bmp) file using the "Save texture" button
- OR -
1-Use the "Mass convert" button to convert a batch of image to or from a 
TEX file. The name of the resulting files will be the same but with the 
extension changed accordingly.

-Color 0 as transparent: When this is checked the completly black pixels of
the image will be interpreted as transparent.

Known issues:
-------------
-The "color 0 as transparent" flag seems to have no effect on the original
Final Fantasy VII engine for 24 bits images. It works fine with Aali's
graphics engine, though.


Thanks:
-------
-To Mirex and Aali for their work on decoding the TEX file specification.
