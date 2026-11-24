VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VB OpenGL Simple example"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   Icon            =   "Proj10GLfrm3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tx(1 To 4), ty(1 To 4), tz(1 To 4), tw(1 To 4) As Single
' +-------------------------------------------------------------------

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

Dim xAngle As GLfloat
Dim yAngle As GLfloat
Dim zAngle As GLfloat
Dim doubleBuffer As GLboolean
Dim displayListInited As GLboolean
Sub FatalError(ByVal strMessage As String)

' +-------------------------------------------------------------------
' | A simple exit handler should something NASTY happen.
' +-------------------------------------------------------------------


    MsgBox "Fatal Error: " & strMessage, vbCritical + vbApplicationModal + vbOKOnly + vbDefaultButton1, "Fatal Error In " & App.Title
    Unload frmMain
    Set frmMain = Nothing
    End
End Sub
Sub SetupPixelFormat(ByVal hDC As Long)

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
Private Sub Form_KeyPress(KeyAscii As Integer)

' +-------------------------------------------------------------------
' | Handle keypresses.  1 rotates the X axis.  2 rotates the Y axis.
' | 3 handles the Z axis.
' +-------------------------------------------------------------------

    Select Case KeyAscii
        Case Asc("1")
            xAngle = xAngle + 10
            glMatrixMode GL_MODELVIEW
            glLoadIdentity
            glTranslatef 0, 0, -3
            glRotatef xAngle, 0.1, 0, 0
            'glRotatef yAngle, 0, 0.1, 0
            'glRotatef zAngle, 0, 0, 1
            Form_Paint
        Case Asc("2")
            yAngle = yAngle + 10
            glMatrixMode GL_MODELVIEW
            glLoadIdentity
            glTranslatef 0, 0, -3
            'glRotatef xAngle, 0.1, 0, 0
            glRotatef yAngle, 0, 0.1, 0
            'glRotatef zAngle, 0, 0, 1
            Form_Paint
        Case Asc("3")
            zAngle = zAngle + 10
            glMatrixMode GL_MODELVIEW
            glLoadIdentity
            glTranslatef 0, 0, -3
            'glRotatef xAngle, 0.1, 0, 0
            'glRotatef yAngle, 0, 0.1, 0
            glRotatef zAngle, 0, 0, 1
            Form_Paint
    End Select
End Sub
Private Sub Form_Load()

' +-------------------------------------------------------------------
' | Set angles and initialize.
' +-------------------------------------------------------------------

    Dim hGLRC As Long
    
    xAngle = 42
    yAngle = 82
    zAngle = 112
    doubleBuffer = GL_TRUE
    displayListInited = GL_FALSE

    SetupPixelFormat hDC
    SetupPalette hDC
    
    hGLRC = wglCreateContext(hDC)
    wglMakeCurrent hDC, hGLRC
    glEnable GL_DEPTH_TEST
    glEnable GL_DITHER
    glDepthFunc GL_LESS
    glClearDepth 1
    glClearColor 0, 0, 0, 0
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    glFrustum -4, 4, -4, 4, 2, 10
    glViewport 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
    
    displayListInited = GL_FALSE

    glMatrixMode GL_MODELVIEW
    glLoadIdentity
    glTranslatef 0, 0, -3
    glRotatef xAngle, 0.1, 0, 0
    glRotatef yAngle, 0, 0.1, 0
    glRotatef zAngle, 0, 0, 1
    Form_Paint

End Sub
Private Sub Form_Paint()
If Linecount > 0 And sweptflag = True Then
Dim tmpquad() As Tquad
ReDim tmpquad(1 To (30 * Linecount))
Dim c, s, a As Integer
Dim tmpr As Variant
Dim tmpg As Variant
Dim tmpb As Variant

' +-------------------------------------------------------------------
' | If a display list has been created, use it.  Otherwise, create it.
' +-------------------------------------------------------------------
    
    If displayListInited = GL_TRUE Then
        glCallList 1
    Else
        glNewList 1, GL_COMPILE_AND_EXECUTE
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
'tw(1) = tmpquad(a).quad(1).homopoint(4) / 1000

tx(2) = tmpquad(a).quad(2).homopoint(1) / 1000
ty(2) = tmpquad(a).quad(2).homopoint(2) / 1000
tz(2) = tmpquad(a).quad(2).homopoint(3) / 1000
'tw(2) = tmpquad(a).quad(2).homopoint(4) / 1000

tx(3) = tmpquad(a).quad(3).homopoint(1) / 1000
ty(3) = tmpquad(a).quad(3).homopoint(2) / 1000
tz(3) = tmpquad(a).quad(3).homopoint(3) / 1000
'tw(3) = tmpquad(a).quad(3).homopoint(4) / 1000

tx(4) = tmpquad(a).quad(4).homopoint(1) / 1000
ty(4) = tmpquad(a).quad(4).homopoint(2) / 1000
tz(4) = tmpquad(a).quad(4).homopoint(3) / 1000
'tw(4) = tmpquad(a).quad(4).homopoint(4) / 1000
   

    'glNormal3f 0, 0, 1

    tmpr = 0.7
    tmpg = 0.9
    tmpb = 0.7
    glColor3f tmpr, tmpg, tmpb

    glVertex3f tx(1), ty(1), tz(1): glVertex3f tx(2), ty(2), tz(2)
    glVertex3f tx(3), ty(3), tz(3): glVertex3f tx(4), ty(4), tz(4)

    
Next s
Next c
    
    glEnd
        glEndList
        displayListInited = GL_TRUE
    End If
    End If
    
    SwapBuffers hDC
        
End Sub
Private Sub Form_Resize()

' +-------------------------------------------------------------------
' | Resize the OpenGL view if the form resizes, and redraw.
' +-------------------------------------------------------------------

    glViewport 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
    Form_Paint
End Sub
Private Sub Form_Unload(Cancel As Integer)

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
