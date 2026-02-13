# GDIPImageList Class for VB6 and TwinBasic

**Purpose:**   

A replacement of the RichClient imageList using scripting.dictionary, GDI+ and/or a TwinBasic Collection to store and extract VB6 native type images, JPG, BMP 
and non-native types such as PNGs. Basically, a wrapper around two dictionaries to provide the functionality of - and compatibility with - a RichClient collection.

Why? RichClient versions 5/6 are currently 32bit only and cannot compile to 64bit binaries using TwinBasic, so this component removes 
a 32bit dependency for loading images to non-RC6 image controls. If you want to convert a Richclient program that uses RC collections then this imagelist class is 
a partial drop-in replacement. I use it in my programs to load JPGs quickly from memory.

Two classes, one for VB6, the other for TwinBasic. If you include both in your program, the appropriate class will be called by the chosen environment. 

The first, **cGdipImageList.cls** uses GDI+ to provide an imageList class that can be used in VB6. PNGs will display 
correctly with transparency when rendered to a VB6 control using GDI+. It will also work with TwinBasic but it is not required 
as TB can use its own collection. Has potential for 64bit compilation using TwinBasic. 

The second, **cTBImageList.cls**, uses a TB collection and thus has potential for 64bit compilation using TwinBasic. 

[![Ask DeepWiki](https://deepwiki.com/badge.svg)](https://deepwiki.com/yereverluvinunclebert/GDIP-Image-List) Click here for a full documentation describing the program structure. 

**Why is cGdipImageList useful?**

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
* If used with TwinBasic will output alpha images direct for display in TB image controls - however, best to use cTBImageList below!

**Why is cTBImageList useful?**

* It is quicker to pull images from a collection in memory than directly from file using LoadPicture
* It is quicker to pull images from TB's collection than a dictionary
* It will load and extract modern image types for display using TwinBasic's native controls, TB having automatic support built-in.
* It can store JPG, BMP, PNG files and other types too.
* It can store images with alpha transparency as TB provides this capability by default, just worth mentioning again!
* It can store images of varying size, not just small 16x16 or 32x32 icons as per the old VB6 imageList.
* It is a TB collection-backed ImageList with no scripting runtime dependency.
* Uses RichClient-familiar syntax to load and unload the images for easy drop-in replacement.
* Avoids RichClient dependency.
* Avoids GDI+ dependency so potentially platform-independent in TB's multiplatform future.
* Avoids runtime obsolescence.
* Has a confirmed 64bit future with TwinBasic.

**Limitations?**

* You can't currently use either of these imageLists as a **full** replacement for RichClient's image list as the rest of Olaf's code is designed specifically to work only with his OWN imageList, eg. CC.RenderSurfaceContent.
  However, even in an RC project it is still useful for loading other non-RichClient image controls. I am working on creating my own non-RC image widgets using GDI+ and/or Cairo and if I manage this, the imageLists will support these types.
  
* VB6 still can't handle PNGs with alpha unless you use something like Elroy's standard picture Ex project or GDI+ render to a VB6 control.
  Alpha PNGs will display with a black background on any standard VB6 control as VB6 does not support transparency. JPGs and BMPs are fine.
  
  TwinBasic version has no such limitations using either class.


**Dependencies:**

If using the cGdipImageList class you have the choice of utilising the scripting.dictionary, or Cristian Buse's Dictionary alternative. If you use the former,
you will need to add a project reference to the MS scripting runtime scrrun.dll. The latter has no external dependency.

<img width="445" height="359" alt="image" src="https://github.com/user-attachments/assets/936f6161-8361-447a-8f32-cc0681ad3656" />

If you use TwinBasic there are no external dependencies.

**Usage:**

Pull both classes into your existing project.

Add a public or private variable as required to a module (BAS) in order to instantiate/create a new GDI+ image list instance.

    #If twinbasic Then
        ' Wrapper around TwinBasic's collection
        Public thisImageList As New cTBImageList
    #Else
        ' new GDI+ image list instance
        Public thisImageList As New cGdipImageList
    #End If

Add a public variable to a module (BAS) to provide an instance counter for each usage of the class.

    Public gGdipImageListInstanceCount As Long

if you are using Cristian Buse's Dictionary replacement for the Scripting.Dictionary     

In Class declarations, comment out this line:

    'Private mDict As scripting.dictionary
    
and replace it with:    

    Private mDict As Dictionary

In Class_Initialize, then comment out this line:

    Set mDict = CreateObject("Scripting.Dictionary")
    
and replace it with:

    Set mDict = New Dictionary 

**Properties/routines available:**

To add an image to the image list:
  
    thisImageList.AddImage "about-icon-dark", App.Path & "\Resources\images\about-icon-dark-1010.jpg"
    thisImageList.AddImage key, filename
  
To add an image to a standard VB6 image control
  
    Set imgAbout.Picture = thisImageList.Picture("about-icon-dark")
    set pic.Picture = thisImageList.Picture(key)
  
To remove an image from the imageList
  
    thisImageList.Remove "about-icon-dark"
    thisImageList.Remove key
  
To obtain a count of the images in the imageList
  
    dictionaryCount = thisImageList.count
  
To check if an image is already loaded into the imageList
  
    If thisImageList.Exists(key) Then ...
  
To set the imageWidth, (currently only functional in the cGdipImageList class, in the TBImageList, I am experimenting with non-GDI+ Cairo to achieve this)
  
    thisImageList.ImageWidth = 150  ' note: by default a value of 0, the image's real width will be used
  
To set the imageHeight, (in the TBImageList, non functional as per the imageWidth)
  
    thisImageList.ImageHeight   ' note: by default a value of 0, the image's real width will be used
  
To set the opacity of the image priror to loading, (in the TBImageList, non functional as per the imageWidth):
  
    thisImageList.ImageOpacity = 100  ' 0-100% - TwinBasic will handle the opacity, VB6 won't.
  
To clear the imageList:
  
    thisImageList.Clear   
  
To enumerate through the imageList externally:
  
    For Each img In thisImageList
      ...
    Next

**Examples of Use:**

I use the imageList in a few of my own programs:

The concept came from my dock replacement, SteamyDock:

![hollerith0002](https://github.com/user-attachments/assets/f661e9c0-40ad-4c83-ad49-6a3de44077fd)

In SteamyDock all the images are loaded at startup into two scripting.dictionaries, one for the large images and the other for the small images. The images being PNGs are then pulled from the dictionary as needed and rendered to screen using GDI+. Recently, the scripting.dictionary was replaced seamlessly with Cristian Buse's VBA dictionary replacement.

The GDIPImageList is also used to place the Jpeg images on the top of the configuration form for all of my recent steampunk 'widgets/trinkets'.

![gdpImageList001](https://github.com/user-attachments/assets/48b95067-c772-40b7-81d0-ffc55f0bf32e)

