VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Run Length Blit by dafhi (MaskBlit in mSurfaceDesc.bas)

'thanks to Robert Rayment for load image (SurfaceDescFromFile)
'using GetDiBits which i still didn't know how to use

Dim Sprite     As SurfaceDescriptor 'two graphics 'objects' (there's nothing to clean up)
Dim BackBuffer As SurfaceDescriptor

Dim mSA        As SAFEARRAY1D 'required to call ColorFill()
Dim my1D()     As Long        'ColorFill intends to show 1D hook subs

Private Type SpritePos
    center_x   As Single
    center_y   As Single
    length     As Single
    angle      As Single
    rot_spd    As Single
End Type

Private Const Num_Vectors As Long = 40

Dim Vector(1 To Num_Vectors) As SpritePos

Dim I As Long

Private Sub Form_Load()
    
    ForeColor = vbWhite
    ScaleMode = vbPixels
    
    SurfaceDescFromFile Sprite, "star.bmp", hDC, FlipRB(vbBlue)
    
    Move 100, 100, 6400, 4800 'backbuffer and vector pos in Form_Resize
    
    Show
    
    FPS_Init
    
    Do While DoEvents
    
        For I = 1 To Num_Vectors
            DrawSprite Vector(I)
        Next
        
        Blit BackBuffer
        
        If CheckFPS Then
            Caption = "FPS: " & Round(sFPS, 1)
        End If
    
        ColorFill BackBuffer, my1D, mSA, FlipRB(vbBlue)
        
    Loop
    
End Sub

Private Sub DrawSprite(pVec As SpritePos)
    
    MaskBlit BackBuffer, Sprite, _
      pVec.center_x + pVec.length * Cos(pVec.angle), _
      pVec.center_y + pVec.length * Sin(pVec.angle)
    
    pVec.angle = pVec.angle + pVec.rot_spd * speed 'time-based variable speed from mGeneral.bas

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    End Select
End Sub

Private Sub Form_Resize()

    For I = 1 To Num_Vectors
        DefineVec Vector(I), _
         -Sprite.Wide / 2, -Sprite.High / 2, _
          I * 7, _
          0, 0.003 * pi * I / Num_Vectors
    Next
    
    CreateSurfaceDesc BackBuffer, hDC, ScaleWidth, ScaleHeight, -ScaleWidth / 2, -ScaleHeight / 2

End Sub

Private Sub DefineVec(pVec As SpritePos, Optional ByVal center_x As Single, Optional ByVal center_y As Single, Optional ByVal pLength As Single = 10, Optional ByVal start_angle As Single = -halfPi, Optional ByVal rotation_speed As Single = 1)
    pVec.center_x = center_x
    pVec.center_y = center_y
    pVec.angle = start_angle
    pVec.length = pLength
    pVec.rot_spd = rotation_speed
End Sub
