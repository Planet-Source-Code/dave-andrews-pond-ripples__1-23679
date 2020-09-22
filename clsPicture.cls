VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Public hDC As Long
Public bmType As Long
Public bmWidth As Long
Public bmHeight As Long
Public bmWidthBytes As Long
Public bmPlanes As Integer
Public bmBitsPixel As Integer
Public bmBits As Long
Sub ClearMemory()
DeleteDC hDC
End Sub

Sub CreateDC(ParentDc As Long)
hDC = CreateCompatibleDC(ParentDc)
End Sub

Public Sub GetImageDC(ParentDc As Long, IMG As Long)
'Left TO Right,Top To Bottom
Dim BMP As BITMAP
Dim TempBMP As Long
'Create a handle out of thin air!
CreateDC ParentDc
'Get the info about our image
GetObject IMG, Len(BMP), BMP
'Here's where we 'paste' our image into the new DC
TempBMP = SelectObject(hDC, IMG)
'Now we set the class objects public variables so they can be treated as properties
bmType = BMP.bmType
bmWidth = BMP.bmWidth
bmHeight = BMP.bmHeight
bmWidthBytes = BMP.bmWidthBytes
bmPlanes = BMP.bmPlanes
bmBitsPixel = BMP.bmBitsPixel
bmBits = BMP.bmBits
'Now we free up that memory for the bitmap object
DeleteObject TempBMP
End Sub
