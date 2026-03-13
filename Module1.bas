Attribute VB_Name = "Module1"
Option Explicit
          
#If twinbasic Then
    ' Wrapper around TwinBasic's collection
    Public thisImageList As New cTBImageList
#Else
    ' new GDI+ image list instance
    Public thisImageList As New cGdipImageList
#End If

' counter for each usage of the class
Public gGdipImageListInstanceCount As Long
