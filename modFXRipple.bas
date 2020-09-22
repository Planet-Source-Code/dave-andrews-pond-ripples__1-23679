Attribute VB_Name = "modFXRipple"
Option Explicit
Public Const Pi = 3.14159265358979
Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Type Ripple
    cX As Long
    cY As Long
    Rad As Integer
    Size As Integer
End Type
Dim Ripples() As Ripple
Dim OrigPic As New clsPicture
Dim TempPic As New clsPicture
Dim DestDC As Long
Dim StartRipple As Long
Dim MaxRad As Integer
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Sub AddRipple(x As Long, y As Long)
On Error GoTo eTrap
Dim i As Integer
i = UBound(Ripples) + 1
ReDim Preserve Ripples(i) As Ripple
Ripples(i).cX = x
Ripples(i).cY = y
Ripples(i).Size = 8
Exit Sub
eTrap:
    i = 0
    Resume Next
End Sub


Sub InitRipples(DestinationHDC As Long, FileName As String)
DestDC = DestinationHDC
OrigPic.GetImageDC DestDC, LoadPicture(FileName)
TempPic.GetImageDC DestDC, LoadPicture(FileName)
MaxRad = IIf(OrigPic.bmWidth > OrigPic.bmHeight, OrigPic.bmWidth, OrigPic.bmHeight)
End Sub




Sub RenderRipples()
Dim i As Long
Dim j As Integer
Dim k As Integer
Dim x As Long
Dim y As Long
'First we copy our orig back to the temp
BitBlt TempPic.hDC, 0, 0, OrigPic.bmWidth, OrigPic.bmHeight, OrigPic.hDC, 0, 0, vbSrcCopy
For i = StartRipple To UBound(Ripples)
    k = Ripples(i).Size + CInt(Rnd * 2)
    For j = 0 To 360 'all ripples start circular
        x = (Ripples(i).Rad * Cos(j * Pi / 180)) + Ripples(i).cX
        y = (Ripples(i).Rad * Sin(j * Pi / 180)) + Ripples(i).cY
        BitBlt TempPic.hDC, x, y, k, k, OrigPic.hDC, x + 2, y + 2, vbSrcCopy
    Next j
    DoEvents
    Ripples(i).Rad = Ripples(i).Rad + 3 ' advance our ripple radius
    If Ripples(i).Size > 1 Then Ripples(i).Size = Ripples(i).Size - 1
Next i
BitBlt DestDC, 0, 0, OrigPic.bmWidth, OrigPic.bmHeight, TempPic.hDC, 0, 0, vbSrcCopy
For i = 0 To UBound(Ripples)
    'since all ripples grow at the same rate, we can exclude the earliest when they get too big.
    If Ripples(i).Rad >= MaxRad Then StartRipple = i + 1
Next i
End Sub


