VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private AcroApp   As Object 'Acrobat.CAcroApp
'Private AcroAVDoc As Object 'Acrobat.CAcroAVDoc
'Private AcroPDDoc As Object 'Acrobat.CAcroPDDoc

Private AcroApp   As Acrobat.CAcroApp
Private AcroAVDoc As Acrobat.CAcroAVDoc
Private AcroPDDoc As Acrobat.CAcroPDDoc

Private bRetValue As Boolean

Private Const PDUseThumbs = 2
Private Const PDUseNone = 1

Private Sub Command1_Click()
  Dim filename        As String
  filename = "c:\download.pdf"
  
  Set AcroApp = CreateObject("AcroExch.App")
'  bRetValue = AcroApp.Show
'  If bRetValue = False Then MsgBox "ERROR"
  
  Set AcroPDDoc = CreateObject("AcroExch.PDDoc")
  bRetValue = AcroPDDoc.Open(filename)
  If bRetValue = False Then MsgBox "ERROR"
  
  Set AcroAVDoc = AcroPDDoc.OpenAVDoc("")
End Sub
Private Sub Command2_Click()
  bRetValue = AcroAVDoc.SetViewMode(PDUseThumbs)
End Sub
Private Sub Command3_Click()
  bRetValue = AcroAVDoc.SetViewMode(PDUseNone)
End Sub

Private Sub Command4_Click()
     AcroAVDoc.PrintPages 0, 0, 3, True, True
    If AcroAVDoc.PrintPagesSilent(0, 0, 3, True, True) Then
        MsgBox "Impresión OK!!!"
    Else
        MsgBox "ERROR!!!"
    End If

End Sub

Private Sub Form_Load()
  Command1.Caption = "Acrobat"
  Command2.Caption = "PDUseThumbs"
  Command3.Caption = "PDUseNone"
  Command4.Caption = "Imprimir"
End Sub


