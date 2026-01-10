**Purpose**   :  A replacement of the RichClient imageList using scripting.dictionary and GDI+ to store and extract VB6 native type images.
(JPG, BMP, PNG tested). Created to provide an imageList class that has potential for 64bit operation using TwinBasic. RichClient 5/6 are 32bit only
and cannot compile to 64bit binaries using TwinBasic. This component removes a dependency on the 32bit Richclient imagelist for loading non-RC6 image controls.

PNGs will display correctly with transparency when rendered to a VB6 control using GDI+.

Why is this ImageList useful?

* It is quicker to pull pre-loaded images from the dictionary at runtime than pulling them when needed, directly from file.
* It can store JPG, BMP, PNG files and other types too.
* It can store images with alpha transparency as GDI+ provides this capability.
* It can store images of varying size, not just small 16x16 or 32x32 icons as was the limit of the old VB6 imageList.
* You have full control as it is FOSS.
* The usage syntax is very similar to that of the Rich Client ImageList for easy drop-in replacement.
* It is a Dictionary-backed ImageList, using the scripting.dictionary object, a standard dependable Windows component.
* When using [Cristian Buse's dictionary replacement](https://github.com/cristianbuse/VBA-FastDictionary), there is no scripting runtime dependency (recommended).
* Avoids RichClient dependency.
* Avoids runtime obsolescence.
* Uses dependable GDI+ to load and unload the images.
* If used with Elroy's standard picture Ex project it can parse and render alpha images (PNGs &c) directly to VB6 picture/imageboxes.

**Limitations?**

* You can't currently use this imageList as a full replacement for RichClient's image list as the rest of Olaf's code is designed specifically to work only with his own imageList, eg. CC.RenderSurfaceContent.
  However, even in an RC project it is still useful for loading other non-RC image controls.
  
* VB6 still can't handle PNGs with alpha unless you use something like Elroy's standard picture Ex project or GDI+ render to a VB6 control.
  Alpha PNGs will display with a black background on any standard VB6 control as VB6 does not support transparency. JPgs and BMPs are fine.

  Note: It should load and display perfectly using TwinBasic's native controls that have automatic PNG support built-in.

**Dependencies:**

If using the scripting.dictionary, you will need to add a project reference to the MS scripting runtime scrrun.dll

<img width="445" height="359" alt="image" src="https://github.com/user-attachments/assets/936f6161-8361-447a-8f32-cc0681ad3656" />



**Usage:**

Add a public or private variable to a module (BAS) in order to instantiate/create a new GDI+ image list instance.

    Public gdipImageList As New cGdipImageList

Add a public variable to a module (BAS) to provide an instance counter for each usage of the class.

    Public gGdipImageListInstanceCount As Long

if you are using Cristian Buse's Dictionary replacement for the Scripting.Dictionary     

In Class declarations, comment out this line:

    'Private mDict As scripting.dictionary
    
and replace it with:    

    Private mDict As Dictionary

In Class_Initialize, then comment out this line:

    'Set mDict = CreateObject("Scripting.Dictionary")
    
and replace it with:

    Set mDict = New Dictionary 

**Properties/routines available:**

To add an image to the image list:
  
    gdipImageList.AddImage "about-icon-dark", App.Path & "\Resources\images\about-icon-dark-1010.jpg"
  
To add an image to a standard VB6 image control
  
    Set imgAbout.Picture = gdipImageList.Picture("about-icon-dark")
  
To remove an image from the imageList
  
    gdipImageList.Remove "about-icon-dark"
  
To obtain a count of the images in the imageList
  
    dictionaryCount = gdipImageList.count
  
To check if an image is already loaded into the imageList
  
    If gdipImageList.Exists(thiskey) Then ...
  
To set the imageWidth
  
    gdipImageList.ImageWidth = 150  ' note: by default a value of 0, the image's real width will be used
  
To set the imageHeight
  
    gdipImageList.ImageHeight  ditto
  
To set the opacity of the image priror to loading
  
    gdipImageList.ImageOpacity = 100  ' 0-100% - TwinBasic will handle the opacity, VB6 won't.
  
To clear the imageList
  
    gdipImageList.Clear   
