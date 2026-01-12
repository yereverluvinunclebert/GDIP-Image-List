# GDIPImageList Class for VB6 and TwinBasic

**Purpose:**   

A replacement of the RichClient imageList using scripting.dictionary, GDI+ and/or a TwinBasic Collection to store and extract VB6 native type images, JPG, BMP 
and non-native types such as PNGs. RichClient versions 5/6 are currently 32bit only and cannot compile to 64bit binaries using TwinBasic, so this component removes 
a 32bit dependency for loading images to non-RC6 image controls.

Two classes, one for VB6, the other for TwinBasic. 

The first, **cGdipImageList.cls** uses GDI+ to provide an imageList class that can be used in VB6. PNGs will display 
correctly with transparency when rendered to a VB6 control using GDI+.

The second, **cTBImageList.cls**, uses a TB collection and thus has potential for 64bit compilation using TwinBasic. 

Why is cGdipImageList useful?

* It is quicker to pull images from a collection in memory than directly from file using LoadPicture
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

' Why is cTBImageList useful?

* It is quicker to pull images from a collection in memory than directly from file using LoadPicture
* It will load and extract modern image types for display using TwinBasic's native controls, TB having automatic support built-in.
* It can store JPG, BMP, PNG files and other types too.
* It can store images with alpha transparency as TB provides this capability
* It can store images of varying size, not just small 16x16 or 32x32 icons as per the old VB6 imageList.
* It is a TB collection-backed ImageList with no scripting runtime dependency.
* Avoids RichClient dependency.
* Avoids GDI+ dependency so potentially platform-independent in TB's multiplatform future.
* Avoids runtime obsolescence.
* Uses RichClient-familiar syntax to load and unload the images for easy drop-in replacement.
* It is quicker to pull images from TB's collection than a dictionary


**Limitations?**

* You can't currently use either of these imageLists as a **full** replacement for RichClient's image list as the rest of Olaf's code is designed specifically to work only with his own imageList, eg. CC.RenderSurfaceContent.
  However, even in an RC project it is still useful for loading other non-RC image controls.
  
* VB6 still can't handle PNGs with alpha unless you use something like Elroy's standard picture Ex project or GDI+ render to a VB6 control.
  Alpha PNGs will display with a black background on any standard VB6 control as VB6 does not support transparency. JPgs and BMPs are fine.


**Dependencies:**

If using the cGdipImageList class you have the choice of utilising the scripting.dictionary, or Cristian Buse's Dictionary alternative. If you use the former,
you will need to add a project reference to the MS scripting runtime scrrun.dll

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


**Examples of Use:**

I use the imageList in a few of my own programs:

The concept came from my dock replacement, SteamyDock:

![hollerith0002](https://github.com/user-attachments/assets/f661e9c0-40ad-4c83-ad49-6a3de44077fd)

In SteamyDock all the images are loaded at startup into two scripting.dictionaries, one for the large images and the other for the small images. They are then pulled from the dictionary as needed.

The GDIPImageList is also used to place the Jpeg images on the top of the configuration form for all of my recent steampunk 'widgets'.

![gdpImageList001](https://github.com/user-attachments/assets/48b95067-c772-40b7-81d0-ffc55f0bf32e)

