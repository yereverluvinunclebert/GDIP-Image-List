Purpose   :  A replacement of the RichClient imageList using scripting.dictionary and GDI+ to store and extract VB6 native type images.
(JPG, BMP, PNG tested)

PNGs will display correctly with transparency when rendered to a VB6 control using GDI+.

Why is this ImageList useful?

* It is a Dictionary-backed ImageList, using the scripting.dictionary object, a Windows component.
* When using Cristian Buse's dictionary replacement, there is no scripting runtime dependency (recommended).
* It can store JPG, BMP, PNG files and other types too.
* It can store images with alpha transparency as GDI+ provides this capability.
* It can store images of varying size, not just small 16x16 or 32x32 icons as was the limit of the old VB6 imageList.
* Avoids RichClient dependency.
* Avoids runtime obsolescence.
* Uses dependable GDI+ to load and unload the images.
* You have full control as it is FOSS.
* The usage syntax is very similar to that of the Rich Client ImageList for easy drop-in replacement.
* It is quicker to pull images from the dictionary at runtime than directly from file.
* If used with Elroy's standard picture Ex project it can parse and render alpha images (PNGs &c) directly to VB6 picture/imageboxes.

Limitations?

* You can't currently use this imageList with RichClient as Olaf's code is designed specifically to work with his own imageList, eg. CC.RenderSurfaceContent.
* VB6 still can't handle PNGs with alpha unless you use something like Elroy's standard picture Ex project or GDI+ render to a VB6 control.
  Alpha PNGs will display with a black background on any standard VB6 control as VB6 does not support transparency.
  It should display perfectly using TwinBasic's native controls that have automatic PNG support built-in.

Usage:

Add a public or private variable to a module (BAS) in order to instantiate/create a new GDI+ image list instance.

**Public gdipImageList As New cGdipImageList**

Add a public variable to a module (BAS) to provide an instance counter for each usage of the class.

**Public gGdipImageListInstanceCount As Long**

Properties/routines available:

* To add an image to the image list - gdipImageList.AddImage "about-icon-dark", App.Path & "\Resources\images\about-icon-dark-1010.jpg"
* To add an image to a standard VB6 image control - Set imgAbout.Picture = gdipImageList.Picture("about-icon-dark")
* To remove an image from the imageList - gdipImageList.Remove "about-icon-dark"
* To obtain a count of the images in the imageList - dictionaryCount = gdipImageList.count
* To check if an image is already loaded into the imageList - If gdipImageList.Exists(thiskey) Then ...
* To set the imageWidth - gdipImageList.ImageWidth = 150  ' note: by default a value of 0, the image's real width will be used
* To set the imageHeight - gdipImageList.ImageHeight  ditto
* To set the opacity of the image priror to loading - gdipImageList.ImageOpacity = 100  ' 0-100% - TwinBasic will handle the opacity, VB6 won't.
* To clear the imageList - gdipImageList.Clear   
