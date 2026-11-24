VERSION 5.00
Begin VB.Form MainForm 
   AutoRedraw      =   -1  'True
   Caption         =   "MEDA 302 - Assignment 1 - 3D Drawing Application - By Simon Cromarty, CandI 4"
   ClientHeight    =   8685
   ClientLeft      =   345
   ClientTop       =   1185
   ClientWidth     =   10485
   ClipControls    =   0   'False
   Icon            =   "Proj10GLfrm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   10485
   Begin VB.CommandButton CmdGL 
      Caption         =   "OpenGL"
      Height          =   375
      Left            =   240
      TabIndex        =   31
      ToolTipText     =   "apply the changes made to original object"
      Top             =   8280
      Width           =   1575
   End
   Begin VB.TextBox TxtMouseX 
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "Current X co-ordinate for mouse"
      Top             =   8280
      Width           =   615
   End
   Begin VB.TextBox TxtMouseY 
      Height          =   285
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   15
      ToolTipText     =   "Current Y co-ordinate for mouse"
      Top             =   8280
      Width           =   615
   End
   Begin VB.HScrollBar Zscrollbar 
      DragIcon        =   "Proj10GLfrm1.frx":0442
      Height          =   255
      LargeChange     =   10
      Left            =   8880
      Max             =   360
      MousePointer    =   9  'Size W E
      SmallChange     =   10
      TabIndex        =   9
      Top             =   6600
      Value           =   1
      Width           =   1335
   End
   Begin VB.HScrollBar Yscrollbar 
      DragIcon        =   "Proj10GLfrm1.frx":074C
      Height          =   255
      LargeChange     =   10
      Left            =   8880
      Max             =   360
      MousePointer    =   9  'Size W E
      SmallChange     =   10
      TabIndex        =   8
      Top             =   6240
      Value           =   1
      Width           =   1335
   End
   Begin VB.HScrollBar Xscrollbar 
      DragIcon        =   "Proj10GLfrm1.frx":0A56
      Height          =   255
      LargeChange     =   10
      Left            =   8880
      Max             =   360
      MousePointer    =   9  'Size W E
      SmallChange     =   10
      TabIndex        =   7
      Top             =   5880
      Value           =   1
      Width           =   1335
   End
   Begin VB.Frame OutputFrame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Output"
      Height          =   8055
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Output Window - Object is displayed here"
      Top             =   120
      Width           =   7215
      Begin VB.PictureBox PictureOut 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000E&
         ClipControls    =   0   'False
         ForeColor       =   &H00000000&
         Height          =   7695
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   7635
         ScaleWidth      =   6915
         TabIndex        =   30
         ToolTipText     =   "Output window - where the object is displayed"
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Frame ViewFrame 
      Caption         =   "View"
      Height          =   3615
      Left            =   7440
      TabIndex        =   17
      ToolTipText     =   "Manipulate the object view"
      Top             =   4560
      Width           =   2895
      Begin VB.TextBox TxtZoomVal 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Current focal length value"
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "&Apply"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "apply the changes made to original object"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.HScrollBar Zoom 
         DragIcon        =   "Proj10GLfrm1.frx":0D60
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   500
         Min             =   1
         MousePointer    =   9  'Size W E
         TabIndex        =   3
         Top             =   600
         Value           =   10
         Width           =   1935
      End
      Begin VB.OptionButton TranslateOption 
         Caption         =   "&Translate"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "move the object"
         Top             =   1800
         Width           =   975
      End
      Begin VB.OptionButton RotateOption 
         Caption         =   "R&otate"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "rotate the object"
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton ScaleOption 
         Caption         =   "&Scale"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "scale the object"
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "&Reset"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "reset current changes"
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Axis:"
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Transform:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Focal Length:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Z:"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Y:"
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "X:"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   1320
         Width           =   255
      End
   End
   Begin VB.Frame DrawFrame 
      Caption         =   "Draw"
      Height          =   4335
      Left            =   7440
      TabIndex        =   0
      ToolTipText     =   "Drawing area - create your object here"
      Top             =   120
      Width           =   2895
      Begin VB.TextBox TxtCLineNum 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   29
         ToolTipText     =   "Current line number"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox TxtCLineCoords 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Current line co-ordinates"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.PictureBox PictureIn 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000E&
         ClipControls    =   0   'False
         ForeColor       =   &H00000000&
         Height          =   2415
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   2355
         ScaleWidth      =   2595
         TabIndex        =   16
         ToolTipText     =   "Input window - Draw lines here"
         Top             =   240
         Width           =   2655
         Begin VB.Line Line2 
            BorderColor     =   &H80000000&
            BorderStyle     =   3  'Dot
            X1              =   0
            X2              =   2640
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            BorderStyle     =   3  'Dot
            X1              =   1320
            X2              =   1320
            Y1              =   0
            Y2              =   2400
         End
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Clear all and restart"
         Top             =   3120
         Width           =   2655
      End
      Begin VB.CommandButton CmdSweep 
         Caption         =   "&Sweep"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Volume sweep the line in the draw window"
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Line:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   495
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Y:"
      Height          =   255
      Left            =   9000
      TabIndex        =   24
      Top             =   8280
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "X:"
      Height          =   255
      Left            =   7920
      TabIndex        =   23
      Top             =   8280
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Mouse Position:"
      Height          =   255
      Left            =   6480
      TabIndex        =   22
      Top             =   8280
      Width           =   1215
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' most variables used in this form are from globals declared within module1

' a variable for holding message boxes
Dim response As Variant


Private Sub CmdSweep_Click()
    ' mostly user interface stuff,
    ' calls the sweep routines when the sweep button is clicked
    If Linecount > 0 Then
        If sweptflag <> True Then
            sweptflag = True
            ReDim tmpobject(1 To 30)
            initialisePerspectiveMatrix (Zoom.Value / 10)
            Call dothesweep
            Call drawit
            drawoption = ""
            barchanged = ""
            ScaleOption.Value = False
            RotateOption.Value = False
            TranslateOption.Value = False
            TxtCLineCoords.Text = ""
            TxtCLineNum.Text = ""
            Rem enable manipulation buttons
            ScaleOption.Enabled = True
            RotateOption.Enabled = True
            TranslateOption.Enabled = True
            Zoom.Enabled = True
            Xscrollbar.Enabled = True
            Yscrollbar.Enabled = True
            Zscrollbar.Enabled = True
            CmdApply.Enabled = True
            CmdReset.Enabled = True
            Rem Disable Sweep button
            CmdSweep.Enabled = False
            Rem Disable drawing window
            PictureIn.Enabled = False
        End If
    End If
End Sub

Private Sub CmdClear_Click()
    ' resets everything
    PictureOut.Cls
    PictureIn.Cls
    TxtCLineCoords.Text = ""
    TxtCLineNum.Text = ""
    ScaleOption.Value = False
    RotateOption.Value = False
    TranslateOption.Value = False
    
    ' reinitialise graphics variables
    Call my3dinit
    
    ' disable manipulation buttons
    ScaleOption.Enabled = False
    RotateOption.Enabled = False
    TranslateOption.Enabled = False
    Zoom.Enabled = False
    Xscrollbar.Enabled = False
    Yscrollbar.Enabled = False
    Zscrollbar.Enabled = False
    CmdApply.Enabled = False
    CmdReset.Enabled = False
    ' re-enable Sweep button
    CmdSweep.Enabled = True
    ' re-enable drawing window
    PictureIn.Enabled = True
End Sub

Private Sub CmdApply_Click()
    ' store the last tranformation in the original object
    Apply = True
    
    Call dothelot
    Call drawit
    
    drawoption = ""
    ScaleOption.Value = False
    RotateOption.Value = False
    TranslateOption.Value = False
    Apply = False
End Sub

Private Sub CmdReset_Click()
    ' reset the object back to the last transformation
    drawoption = ""
    ScaleOption.Value = False
    RotateOption.Value = False
    TranslateOption.Value = False
    PictureOut.Cls
    Call dothelot
    Call drawit
End Sub

Private Sub CmdGL_Click()
    ' Open the OpenGL version of the object
    frmMain.Show
End Sub

Private Sub Form_Load()

    ' initialise graphics variables
    Call my3dinit
    
    ' initialise form variables, scroll bars and buttons
    
    ScaleOption.Value = False
    RotateOption.Value = False
    TranslateOption.Value = False
    TxtZoomVal.Text = Str(Zoom.Value / 10)
    Rem disable manipulation buttons
    ScaleOption.Enabled = False
    RotateOption.Enabled = False
    TranslateOption.Enabled = False
    Zoom.Enabled = False
    Xscrollbar.Enabled = False
    Yscrollbar.Enabled = False
    Zscrollbar.Enabled = False
    CmdApply.Enabled = False
    CmdReset.Enabled = False
End Sub

Private Sub Zoom_Change()
    ' change perspective view based on zoom slider
    Dim tmp As String
    
    tmp = drawoption
    drawoption = "View"
    Rem scroll bar value is in percent...
    initialisePerspectiveMatrix (Zoom.Value / 10)
    TxtZoomVal.Text = Str(Zoom.Value / 10)
    Rem clear the main window
    PictureOut.Cls
    Call dothelot
    Call drawit
    drawoption = tmp
End Sub

Private Sub Xscrollbar_Change()
    ' rotate, translate or scale on the X axis
    barchanged = "X"
    If drawoption = "Rotate" Then
      deg = ((Val(Xscrollbar.Value)) * (Pi / 180))
      initialiseRotationMatrixX (deg)
    End If
    
    If drawoption = "Translate" Then
        Call initialiseTranslationMatrix(Xscrollbar.Value, Yscrollbar.Value, Zscrollbar.Value)
    End If
    
    If drawoption = "Scale" Then
        Call initialiseScalingMatrix((Xscrollbar.Value / 100), (Yscrollbar.Value / 100), (Zscrollbar.Value / 100))
    End If
    Rem clear the main window
    PictureOut.Cls
    
    Call dothelot
    Call drawit
End Sub

Private Sub Yscrollbar_Change()
    ' rotate, translate or scale on the Y axis
    barchanged = "Y"
    If drawoption = "Rotate" Then
        deg = ((Val(Yscrollbar.Value)) * (Pi / 180))
        initialiseRotationMatrixY (deg)
    End If
    
    If drawoption = "Translate" Then
           Call initialiseTranslationMatrix(Xscrollbar.Value, Yscrollbar.Value, Zscrollbar.Value)
    End If
    
    If drawoption = "Scale" Then
         Call initialiseScalingMatrix((Xscrollbar.Value / 100), (Yscrollbar.Value / 100), (Zscrollbar.Value / 100))
    End If
    Rem clear the main window
    PictureOut.Cls
    
    Call dothelot
    Call drawit
End Sub

Private Sub Zscrollbar_Change()
    ' rotate, translate or scale on the Z axis
    barchanged = "Z"
    If drawoption = "Rotate" Then
        deg = ((Val(Zscrollbar.Value)) * (Pi / 180))
        initialiseRotationMatrixZ (deg)
    End If
    
    If drawoption = "Translate" Then
           Call initialiseTranslationMatrix(Xscrollbar.Value, Yscrollbar.Value, Zscrollbar.Value)
    End If
    
    If drawoption = "Scale" Then
         Call initialiseScalingMatrix((Xscrollbar.Value / 100), (Yscrollbar.Value / 100), (Zscrollbar.Value / 100))
    End If
    Rem clear the main window
    PictureOut.Cls
    
    Call dothelot
    Call drawit
End Sub

Private Sub ScaleOption_Click()
    ' if the user selects the scale option.
    ' initialise X, Y and Z scroll bars
    drawoption = ""
    Xscrollbar.Max = 1000
    Xscrollbar.Min = 10
    Xscrollbar.LargeChange = 100
    Xscrollbar.SmallChange = 10
    Xscrollbar.Value = 100
    
    Yscrollbar.Max = 1000
    Yscrollbar.Min = 10
    Yscrollbar.LargeChange = 100
    Yscrollbar.SmallChange = 10
    Yscrollbar.Value = 100
    
    Zscrollbar.Max = 1000
    Zscrollbar.Min = 10
    Zscrollbar.LargeChange = 100
    Zscrollbar.SmallChange = 10
    Zscrollbar.Value = 100
    
    drawoption = "Scale"
End Sub

Private Sub RotateOption_Click()
    ' if the user selects the rotate option.
    ' initialise X, Y and Z scroll bars
    drawoption = ""
    Xscrollbar.Max = 360
    Xscrollbar.Min = 1
    Xscrollbar.LargeChange = 10
    Xscrollbar.SmallChange = 1
    Xscrollbar.Value = 1
    
    Yscrollbar.Max = 360
    Yscrollbar.Min = 1
    Yscrollbar.LargeChange = 10
    Yscrollbar.SmallChange = 1
    Yscrollbar.Value = 1
    
    Zscrollbar.Max = 360
    Zscrollbar.Min = 1
    Zscrollbar.LargeChange = 10
    Zscrollbar.SmallChange = 1
    Zscrollbar.Value = 1
    
    drawoption = "Rotate"
End Sub

Private Sub TranslateOption_Click()
    ' if the user selects the translate option.
    ' initialise X, Y and Z scroll bars
    drawoption = ""
    Xscrollbar.Max = 4000
    Xscrollbar.Min = -4000
    Xscrollbar.LargeChange = 150
    Xscrollbar.SmallChange = 15
    Xscrollbar.Value = 0
    
    Yscrollbar.Max = 4000
    Yscrollbar.Min = -4000
    Yscrollbar.LargeChange = 150
    Yscrollbar.SmallChange = 15
    Yscrollbar.Value = 0
    
    Zscrollbar.Max = 4000
    Zscrollbar.Min = -4000
    Zscrollbar.LargeChange = 150
    Zscrollbar.SmallChange = 15
    Zscrollbar.Value = 0
    
    drawoption = "Translate"
End Sub


Private Sub PictureIn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' read in mouse clicks from the input frame
    If Linecount < 20 Then
      Call draw_initial_lines(Button, Shift, X, Y)
    Else
      response = MsgBox("Sorry, only a maximum of 20 lines can be drawn", 0, "Sorry")
    End If
End Sub

Private Sub PictureOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' give user translated X,Y coordinates for current mouse position
    Myx = (X - 3450)
    Myy = ((Y * -1) + 3810)
    TxtMouseY.Text = Myy
    TxtMouseX.Text = Myx
    PictureOut.ToolTipText = "X: " + Str(Myx) + ", Y: " + Str(Myy)
    
End Sub

Private Sub PictureIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' give user translated X,Y coordinates for current mouse position
    Myx = (X - 1320)
    Myy = ((Y * -1) + 1215)
    TxtMouseY.Text = Myy
    TxtMouseX.Text = Myx
    PictureIn.ToolTipText = "X: " + Str(Myx) + ", Y: " + Str(Myy)
End Sub

Private Sub drawit()
'output the stored object into the output frame
Dim tmpquad() As Tquad
ReDim tmpquad(1 To (30 * Linecount))
Dim c, s, q, l, a As Integer

a = 0
For c = 1 To 30
    For s = 1 To Linecount
        a = a + 1
        tmpquad(a) = (doperspectivebits(tmpobject(a)))
        tmpquad(a) = (dorescalebits(tmpquad(a)))
        tmpquad(a) = (dofittowindow(tmpquad(a)))
        PictureOut.Line (tmpquad(a).quad(1).homopoint(1), tmpquad(a).quad(1).homopoint(2))-(tmpquad(a).quad(2).homopoint(1), tmpquad(a).quad(2).homopoint(2))
        PictureOut.Line -(tmpquad(a).quad(3).homopoint(1), tmpquad(a).quad(3).homopoint(2))
        PictureOut.Line -(tmpquad(a).quad(4).homopoint(1), tmpquad(a).quad(4).homopoint(2))
        PictureOut.Line -(tmpquad(a).quad(1).homopoint(1), tmpquad(a).quad(1).homopoint(2))
    Next s
Next c
End Sub


Private Sub draw_initial_lines(Butt As Integer, Shiftin As Integer, Xin As Single, Yin As Single)
    ' Store line coordinates in a slice.
    ' note, the right button can be used to break lines
    If Butt = 1 Then
        If Counttmp = False Then
            Lx = Xin
            Ly = Yin
            Counttmp = True
        Else
            Linecount = Linecount + 1
            PictureIn.Line (Lx, Ly)-(Xin, Yin)
            Rem these are taken from the last point given!!!
            Myx = (Lx - 1320)
            Myy = ((-Ly) + 1215)
            
            myslice(1).slice(Linecount).aline(1).homopoint(1) = Myx
            myslice(1).slice(Linecount).aline(1).homopoint(2) = Myy
            myslice(1).slice(Linecount).aline(1).homopoint(3) = 1
            myslice(1).slice(Linecount).aline(1).homopoint(4) = 1
            Lx = Xin
            Ly = Yin
            Myx = (Xin - 1320)
            Myy = ((-Yin) + 1215)
            
            myslice(1).slice(Linecount).aline(2).homopoint(1) = Myx
            myslice(1).slice(Linecount).aline(2).homopoint(2) = Myy
            myslice(1).slice(Linecount).aline(2).homopoint(3) = 1
            myslice(1).slice(Linecount).aline(2).homopoint(4) = 1
            
            TxtCLineCoords.Text = (Str(myslice(1).slice(Linecount).aline(1).homopoint(1)) + "," + Str(myslice(1).slice(Linecount).aline(1).homopoint(2)) + " : " + Str(myslice(1).slice(Linecount).aline(2).homopoint(1)) + "," + Str(myslice(1).slice(Linecount).aline(2).homopoint(2)))
            TxtCLineNum.Text = (Str(Linecount))
        End If
    Else
        Counttmp = False
    End If
End Sub
