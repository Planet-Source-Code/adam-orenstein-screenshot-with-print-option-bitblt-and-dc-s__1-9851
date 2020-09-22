VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   1560
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    MsgBox ("Press Scroll Lock to take a picture of the screen")
End Sub

Private Sub mnuexit_Click()
    AskForExit
End Sub

Private Sub mnuPrint_Click()
    Call Printer.PaintPicture(Form1.Image, 0&, 0&)
    Printer.EndDoc
End Sub

Private Sub mnuSave_Click()
    Dim SaveAs As String
    
    SaveAs$ = InputBox("Save file to:", "Save file", App.Path & "\Snapshot.BMP")
    
    Call SavePicture(Form1.Image, SaveAs$)
    
    MsgBox ("File has been saved to " & SaveAs$)
End Sub

Private Sub Timer1_Timer()
    If GetAsyncKeyState(vbKeyScrollLock) Then CaptureScreen Form1
End Sub
