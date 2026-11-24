VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VB OpenGL"
   ClientHeight    =   3810
   ClientLeft      =   11055
   ClientTop       =   1470
   ClientWidth     =   4140
   Icon            =   "Proj10GLfrm2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   240
   End
   Begin VB.Menu Control 
      Caption         =   "&Control"
      Begin VB.Menu mnuLightswitch 
         Caption         =   "&Light Switch"
         Begin VB.Menu mnuLight0 
            Caption         =   "Light &0"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuLight1 
            Caption         =   "Light &1"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuShadeMode 
         Caption         =   "&Shade Mode"
         Begin VB.Menu mnuFlat 
            Caption         =   "&Flat"
         End
         Begin VB.Menu mnuSmooth 
            Caption         =   "&Smooth"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuLightModel 
         Caption         =   "Light &Model"
         Begin VB.Menu mnuLMLocalViewer 
            Caption         =   "Local Viewer"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuLMAmbient 
            Caption         =   "Ambient"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuLM2Sided 
            Caption         =   "&Two - Sided"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuFog 
      Caption         =   "Fo&g"
      Begin VB.Menu mnuFogEnabled 
         Caption         =   "&Enabled"
      End
      Begin VB.Menu mnuFogIncreaseDensity 
         Caption         =   "&Increase Density"
      End
      Begin VB.Menu mnuFogDecreaseDensity 
         Caption         =   "&Decrease Density"
      End
   End
   Begin VB.Menu mnuObject 
      Caption         =   "&Object"
      Begin VB.Menu mnuOReset 
         Caption         =   "&Reset Colours"
      End
      Begin VB.Menu mnuFrontFace 
         Caption         =   "Change &Front Face"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuBackFace 
         Caption         =   "Change &Back Face"
      End
      Begin VB.Menu mnuOAmbient 
         Caption         =   "&Ambient"
         Begin VB.Menu mnuOAR 
            Caption         =   "&Red"
            Begin VB.Menu mnuOARI 
               Caption         =   "&Increase"
            End
            Begin VB.Menu mnuOARD 
               Caption         =   "&Decrease"
            End
         End
         Begin VB.Menu mnuOAG 
            Caption         =   "&Green"
            Begin VB.Menu mnuOAGI 
               Caption         =   "&Increase"
            End
            Begin VB.Menu mnuOAGD 
               Caption         =   "&Decrease"
            End
         End
         Begin VB.Menu mnuOAB 
            Caption         =   "&Blue"
            Begin VB.Menu mnuOABI 
               Caption         =   "&Increase"
            End
            Begin VB.Menu mnuOABD 
               Caption         =   "&Decrease"
            End
         End
      End
      Begin VB.Menu mnuODiffuse 
         Caption         =   "&Diffuse"
         Begin VB.Menu mnuODR 
            Caption         =   "&Red"
            Begin VB.Menu mnuODRI 
               Caption         =   "&Increase"
            End
            Begin VB.Menu mnuODRD 
               Caption         =   "&Decrease"
            End
         End
         Begin VB.Menu mnuODG 
            Caption         =   "&Green"
            Begin VB.Menu mnuODGI 
               Caption         =   "&Increase"
            End
            Begin VB.Menu mnuODGD 
               Caption         =   "&Decrease"
            End
         End
         Begin VB.Menu mnuODB 
            Caption         =   "&Blue"
            Begin VB.Menu mnuODBI 
               Caption         =   "&Increase"
            End
            Begin VB.Menu mnuODBD 
               Caption         =   "&Decrease"
            End
         End
      End
   End
   Begin VB.Menu mnuFloor 
      Caption         =   "&Floor"
      Begin VB.Menu mnuFShow 
         Caption         =   "&Show"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuSky 
      Caption         =   "&Sky"
      Begin VB.Menu mnuSShow 
         Caption         =   "&Show"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuMouse 
      Caption         =   "&Info"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*********** My Declaration Bits ******************
'error count - if it all goes wrong too much!
Dim errorcount As Integer


'fog
Dim fltFogcolour(3) As GLfloat
Dim FogDensity(0) As GLfloat

'homogenous points as in part one
Dim tx(1 To 4), ty(1 To 4), tz(1 To 4) As Single

' a shininess value - all my objects use the same one
Dim MatShininess(0) As GLfloat

'object properties
Dim MatSpecular(3) As GLfloat
Dim MatDiffuseFront(3) As GLfloat
Dim MatAmbientFront(3) As GLfloat
Dim MatDiffuseBack(3) As GLfloat
Dim MatAmbientBack(3) As GLfloat

'Floor
Dim FloorShow As Boolean
Dim FloorSpecular(3) As GLfloat
Dim FloorDiffuse(3) As GLfloat
Dim FloorAmbient(3) As GLfloat

'Sky
Dim BackdropShow As Boolean
Dim BackdropSpecular(3) As GLfloat
Dim BackdropDiffuse(3) As GLfloat
Dim BackdropAmbient(3) As GLfloat

'lighting
Dim LightPosition(3) As GLfloat
Dim LightAmbient(3) As GLfloat
Dim LightSpecular(3) As GLfloat
Dim LightDiffuse(3) As GLfloat

Dim lightOneAmbient(3) As GLfloat, lightOnePosition(3) As GLfloat
Dim lightOneSpecular(3) As GLfloat, lightOneDiffuse(3) As GLfloat
Dim lightonecolour(3) As GLfloat

'moving the view with the mouse...
Dim StartX As GLfloat, StartY As GLfloat
Dim angle As GLfloat, angle2 As GLfloat
Dim zoom1 As GLfloat, zoom2 As GLfloat
Dim moving As Integer

'**************End of My Declaration bits**************
'-------------------------------------------------------------------
'*********** Standard GL stuff when using VB *********
' +-------------------------------------------------------------------
' | First, define all our Win32 BS for getting access to OpenGL and
' | other display whatnots.
' +-------------------------------------------------------------------

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(0 To 255) As PALETTEENTRY
End Type
Private Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlpgaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

Const PFD_TYPE_RGBA = 0
Const PFD_TYPE_COLORINDEX = 1
Const PFD_MAIN_PLANE = 0
Const PFD_DOUBLEBUFFER = 1
Const PFD_DRAW_TO_WINDOW = &H4
Const PFD_SUPPORT_OPENGL = &H20
Const PFD_NEED_PALETTE = &H80

Private Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hDC As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare Function CreatePalette Lib "gdi32" (pPal As LOGPALETTE) As Long
Private Declare Sub DeleteObject Lib "gdi32" (hObject As Long)
Private Declare Sub DescribePixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal PixelFormat As Long, ByVal nBytes As Long, pfd As PIXELFORMATDESCRIPTOR)
Private Declare Function GetDC Lib "gdi32" (ByVal hWnd As Long) As Long
Private Declare Function GetPixelFormat Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Sub GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal start As Long, ByVal entries As Long, ByVal ptrEntries As Long)
Private Declare Sub RealizePalette Lib "gdi32" (ByVal hPalette As Long)
Private Declare Sub SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bln As Long)
Private Declare Function SetPixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal i As Long, pfd As PIXELFORMATDESCRIPTOR) As Boolean
Private Declare Sub SwapBuffers Lib "gdi32" (ByVal hDC As Long)
Private Declare Function wglCreateContext Lib "OpenGL32" (ByVal hDC As Long) As Long
Private Declare Sub wglDeleteContext Lib "OpenGL32" (ByVal hContext As Long)
Private Declare Sub wglMakeCurrent Lib "OpenGL32" (ByVal l1 As Long, ByVal l2 As Long)

Dim hPalette As Long
Dim hGLRC As Long

' +-------------------------------------------------------------------
' | Set some constants.
' +-------------------------------------------------------------------

' global variables for spinning things
Dim spin As GLint
Dim planetspin As GLfloat

Sub FatalError(ByVal strMessage As String)
'*********** Standard GL sub *********
' +-------------------------------------------------------------------
' | A simple exit handler should something NASTY happen.
' +-------------------------------------------------------------------


    MsgBox "Fatal Error: " & strMessage, vbCritical + vbApplicationModal + vbOKOnly + vbDefaultButton1, "Fatal Error In " & App.Title
    Unload frmMain
    Set frmMain = Nothing
    End
End Sub

' *******************************************************
' * Start of Standard OpenGL initialisation routines    *
' *******************************************************

Sub SetupPixelFormat(ByVal hDC As Long)
'*********** Standard GL sub *********
' I don't understand this much, but it works
' something to do with using OpenGL under win32 - see below
' +-------------------------------------------------------------------
' | Retrieve/set a Win32 pixel format for OpenGL modes with double-
' | buffering, and direct draw to window with RGBA color mode.
' | 16bit (65536 colors) depth is preferable.
' +-------------------------------------------------------------------

    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim PixelFormat As Integer
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 16
    pfd.cDepthBits = 16
    pfd.iLayerType = PFD_MAIN_PLANE
    PixelFormat = ChoosePixelFormat(hDC, pfd)
    If PixelFormat = 0 Then FatalError "Could not retrieve pixel format!"
    SetPixelFormat hDC, PixelFormat, pfd
End Sub
Sub SetupPalette(ByVal lhDC As Long)
'*********** Standard GL sub *********
' I don't understand this much, but it works
' something to do with using OpenGL under win32
' +-------------------------------------------------------------------
' | Initialize the Win32 form palette.
' +-------------------------------------------------------------------
    
    Dim PixelFormat As Long
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim pPal As LOGPALETTE
    Dim PaletteSize As Long
    PixelFormat = GetPixelFormat(lhDC)
    DescribePixelFormat lhDC, PixelFormat, Len(pfd), pfd
    If (pfd.dwFlags And PFD_NEED_PALETTE) <> 0 Then
        PaletteSize = 2 ^ pfd.cColorBits
    Else
        Exit Sub
    End If
    
    pPal.palVersion = &H300
    pPal.palNumEntries = PaletteSize
    Dim redMask As Long
    Dim GreenMask As Long
    Dim BlueMask As Long
    Dim i As Long
    redMask = 2 ^ pfd.cRedBits - 1
    GreenMask = 2 ^ pfd.cGreenBits - 1
    BlueMask = 2 ^ pfd.cBlueBits - 1
    For i = 0 To PaletteSize - 1
        With pPal.palPalEntry(i)
            .peRed = i
            .peGreen = i
            .peBlue = i
            .peFlags = 0
        End With
    Next
    GetSystemPaletteEntries hDC, 0, 256, VarPtr(pPal.palPalEntry(0))
    hPalette = CreatePalette(pPal)
    If hPalette <> 0 Then
        SelectPalette lhDC, hPalette, False
        RealizePalette lhDC
    End If
End Sub

Private Sub Form_Resize()
'*********** Standard GL sub *********
' +-------------------------------------------------------------------
' | Resize the OpenGL view if the form resizes, and redraw.
' +-------------------------------------------------------------------

    On Error GoTo BadResize
    
    glViewport 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    ' I've changed the angle and depth in the following line
    gluPerspective 50, (frmMain.ScaleWidth / frmMain.ScaleHeight), 1, 160
    glMatrixMode GL_MODELVIEW
    
    Form_Paint
    Exit Sub
BadResize:
    MsgBox "Error #" & CStr(Err.Number) & " has occured in " & CStr(Err.Source) & " during resizing." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Description: " & Err.Description
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
'*********** Standard GL sub *********
' +-------------------------------------------------------------------
' | Release OpenGL if we decide to quit.
' +-------------------------------------------------------------------
    
    If hGLRC <> 0 Then
        wglMakeCurrent 0, 0
        wglDeleteContext hGLRC
    End If
    
    If hPalette <> 0 Then
        DeleteObject hPalette
    End If

End Sub
' *******************************************************
' * End of Standard OpenGL initialisation routines    *
' *******************************************************


Private Sub Form_Load()
' My initialisation routine
'*********** Initial set up of loads of bits...
' +-------------------------------------------------------------------
' | Set angles and initialize.
' +-------------------------------------------------------------------
    On Error GoTo BadInit
    Dim hGLRC As Long
    
    spin = 0

    SetupPixelFormat hDC
    
    hGLRC = wglCreateContext(hDC)
    wglMakeCurrent hDC, hGLRC
'********* The rest of this sub is mine*************
    ' set moving value to 0
    moving = 0

    'reset error count
    errorcount = 0
    ' initialise spin value
    planetspin = 0
    
    'fog
    fltFogcolour(0) = 0.6
    fltFogcolour(1) = 0.6
    fltFogcolour(2) = 0.7
    fltFogcolour(3) = 1
    
    FogDensity(0) = 0.04
    
    ' Lighting Properties
    'light 0
    LightAmbient(0) = 0.15
    LightAmbient(1) = 0.15
    LightAmbient(2) = 0.15
    LightAmbient(3) = 1
    
    LightDiffuse(0) = 0.4
    LightDiffuse(1) = 0.4
    LightDiffuse(2) = 0.4
    LightDiffuse(3) = 1
    
    LightSpecular(0) = 0.4
    LightSpecular(1) = 0.4
    LightSpecular(2) = 0.4
    LightSpecular(3) = 1
    
    LightPosition(0) = 0
    LightPosition(1) = 0
    LightPosition(2) = 0
    LightPosition(3) = 1
    
    'light 1
    lightonecolour(0) = 1
    lightonecolour(1) = 1
    lightonecolour(2) = 1
    lightonecolour(3) = 1
    
    lightOneAmbient(0) = 0.3
    lightOneAmbient(1) = 0.3
    lightOneAmbient(2) = 0.3
    lightOneAmbient(3) = 1
    
    lightOneDiffuse(0) = 0.9
    lightOneDiffuse(1) = 0.9
    lightOneDiffuse(2) = 0.9
    lightOneDiffuse(3) = 1
    
    lightOneSpecular(0) = 0.9
    lightOneSpecular(1) = 0.9
    lightOneSpecular(2) = 0.9
    lightOneSpecular(3) = 1
    
    lightOnePosition(0) = -0.5
    lightOnePosition(1) = 0.5
    lightOnePosition(2) = -0.5
    lightOnePosition(3) = 1
    
    
    ' object properties
    MatSpecular(0) = 0.9
    MatSpecular(1) = 0.6
    MatSpecular(2) = 0.6
    MatSpecular(3) = 1
    
    MatAmbientFront(0) = 0.1
    MatAmbientFront(1) = 0
    MatAmbientFront(2) = 0
    MatAmbientFront(3) = 1
    
    MatDiffuseFront(0) = 0.7
    MatDiffuseFront(1) = 0
    MatDiffuseFront(2) = 0
    MatDiffuseFront(3) = 1
    
    MatAmbientBack(0) = 0.097
    MatAmbientBack(1) = 0.03
    MatAmbientBack(2) = 0
    MatAmbientBack(3) = 1
    
    MatDiffuseBack(0) = 0.67
    MatDiffuseBack(1) = 0.3
    MatDiffuseBack(2) = 0
    MatDiffuseBack(3) = 1
    
    MatShininess(0) = 125
    
    'floor properties
    FloorShow = True
    
    FloorSpecular(0) = 0.1
    FloorSpecular(1) = 0.7
    FloorSpecular(2) = 0.1
    FloorSpecular(3) = 1
    
    FloorAmbient(0) = 0
    FloorAmbient(1) = 0.7
    FloorAmbient(2) = 0
    FloorAmbient(3) = 1
    
    FloorDiffuse(0) = 0
    FloorDiffuse(1) = 0.6
    FloorDiffuse(2) = 0
    FloorDiffuse(3) = 1
    
    'Backdrop properties
    BackdropShow = True
    
    BackdropSpecular(0) = 0.99
    BackdropSpecular(1) = 0.99
    BackdropSpecular(2) = 0.99
    BackdropSpecular(3) = 1
    
    BackdropAmbient(0) = 0.7
    BackdropAmbient(1) = 0.7
    BackdropAmbient(2) = 0.9
    BackdropAmbient(3) = 1
    
    BackdropDiffuse(0) = 0.7
    BackdropDiffuse(1) = 0.7    ' basically the object colour
    BackdropDiffuse(2) = 0.9
    BackdropDiffuse(3) = 1
    
    ' black background (if there's no sky)
    glClearColor 0, 0, 0, 1
    ' smooth rather than flat shading the user can change this whilst running
    glShadeModel GL_SMOOTH
    
   'lighting again
   'light 0 settings
   glLightfv GL_LIGHT0, GL_AMBIENT, LightAmbient(0)
   glLightfv GL_LIGHT0, GL_DIFFUSE, LightDiffuse(0)
   glLightfv GL_LIGHT0, GL_SPECULAR, LightSpecular(0)
   glLightf GL_LIGHT0, GL_CONSTANT_ATTENUATION, 0.7
   glLightf GL_LIGHT0, GL_LINEAR_ATTENUATION, 0.5

   'light 1 settings
   glLightfv GL_LIGHT1, GL_COLOR, lightonecolour(0)
   glLightfv GL_LIGHT1, GL_AMBIENT, lightOneAmbient(0)
   glLightfv GL_LIGHT1, GL_DIFFUSE, lightOneDiffuse(0)
   glLightfv GL_LIGHT1, GL_SPECULAR, lightOneSpecular(0)
   glLightfv GL_LIGHT1, GL_POSITION, lightOnePosition(0)
   glLightf GL_LIGHT1, GL_CONSTANT_ATTENUATION, 0.1
   glLightf GL_LIGHT1, GL_LINEAR_ATTENUATION, 0.05
   
   ' render both the inside and outside of objects
   glLightModeli GL_LIGHT_MODEL_TWO_SIDE, 1
   
    'turn on the lights
    glEnable GL_LIGHT0
    glEnable GL_LIGHT1

    'enable all the flashy lighting effects
    glEnable GL_LIGHTING

    ' enable the depth sort type things
    glEnable GL_DEPTH_TEST
    
    ' dithering - I don't think I need this
    'glEnable GL_DITHER
    
    ' calculation of normal vectors
    glEnable GL_AUTO_NORMAL
    glEnable GL_NORMALIZE
    
    ' depth sort things
    glDepthFunc GL_LEQUAL
    
    ' anti-clockwise front surface rendering
    glFrontFace GL_CCW
    
    ' viewing area
    glViewport 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
    
    'draw the stuff
    Form_Paint
    
    
    Exit Sub
BadInit:
    'error handling
        MsgBox "Error #" & CStr(Err.Number) & " has occured in " & CStr(Err.Source) & " during initialization." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Description: " & Err.Description
        errorcount = errorcount + 1
        Exit Sub
End Sub
' *******************************************************
' ******* Start of Object and Scene routines    *********
' *******************************************************
Private Sub auxSolidObject()
' this originally had assorted object properties in it,
' but now it just calls another sub
        On Error GoTo BadObject
        MyObject
        Exit Sub
BadObject:
    MsgBox "Error #" & CStr(Err.Number) & " has occured in " & CStr(Err.Source) & " during auxSolidObject creation." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Description: " & Err.Description
    Exit Sub

End Sub

Private Sub auxSolidSphere(radius As GLdouble)
'*********** Taken from example source *********
' I've taken out the object properties, as they're set elsewhere
    On Error GoTo BadSphere
    Dim QuadOBJ As GLUquadric
      
    QuadOBJ = gluNewQuadric
    gluQuadricDrawStyle QuadOBJ, GLU_FILL
    gluQuadricNormals QuadOBJ, GLU_SMOOTH
    gluSphere QuadOBJ, radius, 16, 16
    Exit Sub
BadSphere:
    MsgBox "Error #" & CStr(Err.Number) & " has occured in " & CStr(Err.Source) & " during auxSolidSphere creation." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Description: " & Err.Description
    Exit Sub
End Sub

Private Sub myScene()
'***********another one of mine!!*********
On Error GoTo BadScene
  
    
    'floor bits - actually a large sphere
    If FloorShow = True Then
        ' make sure only the floor is affected
        glPushMatrix
        
        ' move it down to the bottom of the scene
        glTranslated 0, -123, 0
        
        ' spin it on its axis
        glRotated (spin / 5), 1, 0, 0
        
        'set its properties
        glMaterialfv GL_FRONT_AND_BACK, GL_AMBIENT, FloorAmbient(0)
        glMaterialfv GL_FRONT_AND_BACK, GL_SPECULAR, FloorSpecular(0)
        
        glMaterialfv GL_FRONT_AND_BACK, GL_DIFFUSE, FloorDiffuse(0)
        glMaterialfv GL_FRONT_AND_BACK, GL_SHININESS, MatShininess(0)
    
        'draw the floor sphere
        auxSolidSphere 120
        
        'back to the last matrix
        glPopMatrix
    End If
    
    'sun bits
    glPushMatrix
    
    glRotated spin, 1, 1, 1
    
            'draw its light
            glDisable GL_LIGHTING
            glLightfv GL_LIGHT1, GL_POSITION, lightOnePosition(0)
            glEnable GL_LIGHTING
   
    glPopMatrix
    
    'backdrop bits - another sphere ( the whole scene sits inside this one)...
    'should get a really good cloud effect with a good light source
    ' this is all similar to the floor
    If BackdropShow = True Then
        glPushMatrix
        
        'spin it a bit
        glRotated spin, 0, 1, 1
        
        'object properties
        glMaterialfv GL_FRONT_AND_BACK, GL_AMBIENT, BackdropAmbient(0)
        glMaterialfv GL_FRONT_AND_BACK, GL_SPECULAR, BackdropSpecular(0)
        
        glMaterialfv GL_FRONT_AND_BACK, GL_DIFFUSE, BackdropDiffuse(0)
        glMaterialfv GL_FRONT_AND_BACK, GL_SHININESS, MatShininess(0)
    
        'draw it
        auxSolidSphere 60
    
        glPopMatrix
    End If
    
    Exit Sub
BadScene:
    'handle errors
    MsgBox "Error #" & CStr(Err.Number) & " has occured in " & CStr(Err.Source) & " during myScene creation." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Description: " & Err.Description
    Exit Sub

End Sub

Private Sub MyObject()
'***************** This one's all mine **************
On Error GoTo BadObject

    ' make sure only the object itself is affected
    glPushMatrix
    ' Clockwise rendering of faces
    glFrontFace GL_CW
    
    ' material settings for object - not front and back faces are different
    glMaterialfv GL_FRONT, GL_AMBIENT, MatAmbientFront(0)
    glMaterialfv GL_BACK, GL_AMBIENT, MatAmbientBack(0)

    glMaterialfv GL_FRONT_AND_BACK, GL_SPECULAR, MatSpecular(0)
    
    
    glMaterialfv GL_FRONT, GL_DIFFUSE, MatDiffuseFront(0)
    glMaterialfv GL_BACK, GL_DIFFUSE, MatDiffuseBack(0)

    glMaterialfv GL_FRONT_AND_BACK, GL_SHININESS, MatShininess(0)

'my object bits
'read in all the quads / points from the object in part 1
'of the assignment and stick them into glquads
If Linecount > 0 And sweptflag = True Then
  Dim tmpquad() As Tquad
  ReDim tmpquad(1 To (30 * Linecount))
  Dim c, s, a As Integer
  
  'depthsort stuff
  glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
  
  glBegin GL_QUADS

    a = 0
    For c = 1 To 30
    For s = 1 To Linecount
        a = a + 1
        
        tmpquad(a) = tmpobject(a)
        
        tx(1) = tmpquad(a).quad(1).homopoint(1) / 1000
        ty(1) = tmpquad(a).quad(1).homopoint(2) / 1000
        tz(1) = tmpquad(a).quad(1).homopoint(3) / 1000
        
        tx(2) = tmpquad(a).quad(2).homopoint(1) / 1000
        ty(2) = tmpquad(a).quad(2).homopoint(2) / 1000
        tz(2) = tmpquad(a).quad(2).homopoint(3) / 1000
        
        tx(3) = tmpquad(a).quad(3).homopoint(1) / 1000
        ty(3) = tmpquad(a).quad(3).homopoint(2) / 1000
        tz(3) = tmpquad(a).quad(3).homopoint(3) / 1000
        
        tx(4) = tmpquad(a).quad(4).homopoint(1) / 1000
        ty(4) = tmpquad(a).quad(4).homopoint(2) / 1000
        tz(4) = tmpquad(a).quad(4).homopoint(3) / 1000
    
        glVertex3f tx(1), ty(1), tz(1): glVertex3f tx(2), ty(2), tz(2)
        glVertex3f tx(3), ty(3), tz(3): glVertex3f tx(4), ty(4), tz(4)
   
    Next s
    Next c
    
  glEnd
End If
' back to counter-clockwise face rendering
glFrontFace GL_CCW
' back to the last matrix
glPopMatrix
Exit Sub

BadObject:

    'handle errors
    If errorcount <= 5 Then
        MsgBox "Error #" & CStr(Err.Number) & " has occured in " & CStr(Err.Source) & " during Object mathematical operations." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Description: " & Err.Description
        errorcount = errorcount + 1
    Else
        MsgBox "Error #" & CStr(Err.Number) & " has occured in " & CStr(Err.Source) & " during Object mathematical operations." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Description: " & Err.Description & Chr$(13) & Chr$(10) & "This window will close, click on the openGL button again to have another go."
        Unload Me
        Exit Sub
    End If
    Exit Sub
End Sub

Private Sub Form_Paint()
'********* This one's mine (now) ************
On Error GoTo BadRender
    
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    glPushMatrix
        ' move object into view
        glTranslatef 0, 0, -5
        ' if the mouse is moved, rotate the lot
        
        glRotatef angle2, 1, 0, 0
        glRotatef angle, 0, 1, 0
        
        auxSolidObject ' draw my object
        
        myScene     ' note, draw the scene around the object!
    
    glPopMatrix
    
    SwapBuffers hDC
    Exit Sub
BadRender:
    MsgBox "Error #" & CStr(Err.Number) & " has occured in " & CStr(Err.Source) & " during rendering." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Description: " & Err.Description
    Exit Sub
End Sub

Private Sub Timer1_Timer()
'*******   This is mine ************
' simple timer routine to animate scene
' set spin values to be used by objects later
On Error GoTo badspin
spin = (spin + 1) Mod 360
planetspin = (planetspin + 0.15) Mod 360
Form_Paint
Exit Sub
badspin:
    MsgBox "Error #" & CStr(Err.Number) & " has occured in " & CStr(Err.Source) & " during timer spin ." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Description: " & Err.Description
    Exit Sub
End Sub

' *******************************************************
' ******* End of Object and Scene routines    *********
' *******************************************************

' *******************************************************
' ******* Start of mouse/keyboard User interface        *
' ******                    routines            *********
' * most of these started out in the examples and have  *
' *       been adapted by me do do extra things         *
' * I'm particularly proud of the zoom in/out function  *
' *         by using the right mouse button...          *
' *******************************************************

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        moving = 1
        StartX = X
        StartY = Y
    End If
        If Button = vbRightButton Then
        moving = 2
        StartX = X
        StartY = Y
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' if the left button is down, spin the view
    If moving = 1 Then
        angle = angle + (X - StartX)
        angle2 = angle2 + (Y - StartY)
        StartX = X
        StartY = Y
        Form_Paint
    End If
    If moving = 2 Then
    ' if the right button is down, zoom in and out
        zoom1 = zoom1 + (X - StartX)
        zoom2 = zoom2 + (Y - StartY)
        StartX = X
        StartY = Y
    ' bugger about with the perspective projection stuff
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    ' change the viewing angle in the following line based on the zoom value
    gluPerspective zoom2 Mod 90, (frmMain.ScaleWidth / frmMain.ScaleHeight), 1, 160
    glMatrixMode GL_MODELVIEW
    Form_Paint
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then moving = 0
    If Button = vbRightButton Then
    moving = 0
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' spin things faster than normal
    If KeyCode = vbKeySpace Then
        spin = (spin + 3) Mod 360
        Form_Paint
    End If
End Sub

' *******************************************************
' ******* End of mouse/keyb User interface routines     *
' *******************************************************

'****************************************************
'* Loads of menu routines - written by me to enable *
'* user to change various values whilst the program *
'* is running - this is good to illustrate all the  *
'* various bits and pieces that I've put in.        *
'* It also means you can sit and fiddle for hours   *
'****************************************************

Private Sub mnuBackFace_Click()
' affects which faces colours are changed - in this case the back
If mnuBackFace.Checked = False Then
    mnuFrontFace.Checked = False
    mnuBackFace.Checked = True
End If
End Sub

Private Sub mnuFlat_Click()
' change to flat shading mode
' slightly blockier but interesting to look at
If mnuFlat.Checked = False Then
    mnuSmooth.Checked = False
    mnuFlat.Checked = True
    glShadeModel GL_FLAT
    
Else
    mnuSmooth.Checked = True
    mnuFlat.Checked = False
    glShadeModel GL_SMOOTH
End If
End Sub

Private Sub mnuFogDecreaseDensity_Click()
' as it says, decrease the density of the fog
FogDensity(0) = FogDensity(0) - 0.01
glFogf GL_FOG_DENSITY, FogDensity(0)
End Sub

Private Sub mnuFogEnabled_Click()
' turn fogging on or off
If mnuFogEnabled.Checked = False Then
glEnable GL_FOG
    glFogi GL_FOG_MODE, GL_EXP2
    glFogf GL_FOG_START, 200
    glFogf GL_FOG_END, 40000
    glFogf GL_FOG_DENSITY, FogDensity(0)
    glFogfv GL_FOG_COLOR, fltFogcolour(0)
    mnuFogEnabled.Checked = True
Else
    glDisable GL_FOG
    mnuFogEnabled.Checked = False
End If
End Sub

Private Sub mnuFogIncreaseDensity_Click()
'increase the density of the fog
FogDensity(0) = FogDensity(0) + 0.01
glFogf GL_FOG_DENSITY, FogDensity(0)
End Sub

Private Sub mnuFrontFace_Click()
' affects which faces colours are changed - in this case the front
If mnuFrontFace.Checked = False Then
    mnuFrontFace.Checked = True
    mnuBackFace.Checked = False
End If
End Sub

Private Sub mnuFShow_Click()
' Show or hide the floor object
If mnuFShow.Checked = True Then
    mnuFShow.Checked = False
    FloorShow = False
Else
    mnuFShow.Checked = True
    FloorShow = True
End If
    
End Sub

Private Sub mnuLight0_Click()
' light switch for light 0
If mnuLight0.Checked = False Then
    mnuLight0.Checked = True
    glEnable GL_LIGHT0
Else
    mnuLight0.Checked = False
    glDisable GL_LIGHT0
End If

End Sub

Private Sub mnuLight1_Click()
' light switch for light 1
If mnuLight1.Checked = False Then
    mnuLight1.Checked = True
    glEnable GL_LIGHT1
Else
    mnuLight1.Checked = False
    glDisable GL_LIGHT1
End If
End Sub

Private Sub mnuLM2Sided_Click()
' change the light model.
' note this menu option is disabled - cos it doesn't work properly
' to me this is the nicest looking lighting model anyway
If mnuLM2Sided.Checked = False Then
    mnuLM2Sided.Checked = True
    mnuLMAmbient.Checked = False
    mnuLMLocalViewer.Checked = False
    glLightModeli GL_LIGHT_MODEL_TWO_SIDE, 1
End If
End Sub

Private Sub mnuLMAmbient_Click()
' change the light model.
' note this menu option is disabled - cos it doesn't work properly
If mnuLMAmbient.Checked = False Then
    mnuLMAmbient.Checked = True
    mnuLM2Sided.Checked = False
    mnuLMLocalViewer.Checked = False
    glLightModeli GL_LIGHT_MODEL_AMBIENT, 1
End If
End Sub

Private Sub mnuLMLocalViewer_Click()
' change the light model.
' note this menu option is disabled - cos it doesn't work properly
If mnuLMLocalViewer.Checked = False Then
    mnuLMLocalViewer.Checked = True
    mnuLMAmbient.Checked = False
    mnuLM2Sided.Checked = False
    glLightModeli GL_LIGHT_MODEL_LOCAL_VIEWER, 1
End If
End Sub

Private Sub mnuMouse_Click()
'actually the info menu - display extra information
frmAbout.Show
End Sub

Private Sub mnuOABD_Click()
'decrement blue ambient
If mnuFrontFace.Checked = True Then
    MatAmbientFront(2) = MatAmbientFront(2) - 0.1
Else
    MatAmbientBack(2) = MatAmbientBack(2) - 0.1
End If
End Sub

Private Sub mnuOABI_Click()
'increment blue ambient
If mnuFrontFace.Checked = True Then
    MatAmbientFront(2) = MatAmbientFront(2) + 0.1
Else
    MatAmbientBack(2) = MatAmbientBack(2) + 0.1
End If
End Sub

Private Sub mnuOAGD_Click()
'decrement green ambient
If mnuFrontFace.Checked = True Then
    MatAmbientFront(1) = MatAmbientFront(1) - 0.1
Else
    MatAmbientBack(1) = MatAmbientBack(1) - 0.1
End If
End Sub

Private Sub mnuOAGI_Click()
'increment green ambient
If mnuFrontFace.Checked = True Then
    MatAmbientFront(1) = MatAmbientFront(1) + 0.1
Else
    MatAmbientBack(1) = MatAmbientBack(1) + 0.1
End If
End Sub

Private Sub mnuOARD_Click()
'decrement red ambient
If mnuFrontFace.Checked = True Then
    MatAmbientFront(0) = MatAmbientFront(0) - 0.1
Else
    MatAmbientBack(0) = MatAmbientBack(0) - 0.1
End If
End Sub

Private Sub mnuOARI_Click()
'increment red ambient
If mnuFrontFace.Checked = True Then
    MatAmbientFront(0) = MatAmbientFront(0) + 0.1
Else
    MatAmbientBack(0) = MatAmbientBack(0) + 0.1
End If

End Sub
Private Sub mnuODBD_Click()
'decrement blue diffuse
If mnuFrontFace.Checked = True Then
    MatDiffuseFront(2) = MatDiffuseFront(2) - 0.1
Else
    MatDiffuseBack(2) = MatDiffuseBack(2) - 0.1
End If
End Sub

Private Sub mnuODBI_Click()
'increment blue diffuse
If mnuFrontFace.Checked = True Then
    MatDiffuseFront(2) = MatDiffuseFront(2) + 0.1
Else
    MatDiffuseBack(2) = MatDiffuseBack(2) + 0.1
End If
End Sub

Private Sub mnuODGD_Click()
'decrement green diffuse
If mnuFrontFace.Checked = True Then
    MatDiffuseFront(1) = MatDiffuseFront(1) - 0.1
Else
    MatDiffuseBack(1) = MatDiffuseBack(1) - 0.1
End If
End Sub

Private Sub mnuODGI_Click()
'increment green diffuse
If mnuFrontFace.Checked = True Then
    MatDiffuseFront(1) = MatDiffuseFront(1) + 0.1
Else
    MatDiffuseBack(1) = MatDiffuseBack(1) + 0.1
End If
End Sub

Private Sub mnuODRD_Click()
'decrement red diffuse
If mnuFrontFace.Checked = True Then
    MatDiffuseFront(0) = MatDiffuseFront(0) - 0.1
Else
    MatDiffuseBack(0) = MatDiffuseBack(0) - 0.1
End If
End Sub

Private Sub mnuODRI_Click()
'increment red diffuse
If mnuFrontFace.Checked = True Then
    MatDiffuseFront(0) = MatDiffuseFront(0) + 0.1
Else
    MatDiffuseBack(0) = MatDiffuseBack(0) + 0.1
End If

End Sub


Private Sub mnuOReset_Click()
    ' reset the colour values back to the original
    MatAmbientFront(0) = 0.1
    MatAmbientFront(1) = 0
    MatAmbientFront(2) = 0
    MatAmbientFront(3) = 1
    
    MatDiffuseFront(0) = 0.7
    MatDiffuseFront(1) = 0
    MatDiffuseFront(2) = 0
    MatDiffuseFront(3) = 1
    
    MatAmbientBack(0) = 0.097
    MatAmbientBack(1) = 0.03
    MatAmbientBack(2) = 0
    MatAmbientBack(3) = 1
    
    MatDiffuseBack(0) = 0.67
    MatDiffuseBack(1) = 0.3 '.2     ' basically the object colour
    MatDiffuseBack(2) = 0 '.2
    MatDiffuseBack(3) = 1
End Sub

Private Sub mnuSmooth_Click()
' change to smooth shading mode
If mnuSmooth.Checked = False Then
    mnuSmooth.Checked = True
    mnuFlat.Checked = False
    glShadeModel GL_SMOOTH
Else
    mnuSmooth.Checked = False
    mnuFlat.Checked = True
    glShadeModel GL_FLAT
End If
End Sub

Private Sub mnuSShow_Click()
' show or hide the sky
If mnuSShow.Checked = True Then
    mnuSShow.Checked = False
    BackdropShow = False
Else
    mnuSShow.Checked = True
    BackdropShow = True
End If
End Sub
'****************************************************
'********* End of all the menu stuff... *************
'****************************************************


