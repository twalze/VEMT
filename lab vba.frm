VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10245
   OleObjectBlob   =   "lab vba.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub buttonClear_Click()
result.Caption = ""
End Sub

Private Sub buttonExit_Click()
End
End Sub

Private Sub buttonShow_Click()
If Val(inputWindow.Text) >= 8 And Val(inputWindow.Text) < 12 Then result.Caption = "Доброе утро!"
If Val(inputWindow.Text) >= 12 And Val(inputWindow.Text) < 17 Then result.Caption = "Добрый день!"
If Val(inputWindow.Text) >= 17 And Val(inputWindow.Text) < 23 Then result.Caption = "Добрый вечер!"
If Val(inputWindow.Text) >= 23 And Val(inputWindow.Text) <= 24 Or Val(inputWindow.Text) >= 24 And Val(inputWindow.Text) < 8 Then result.Caption = "Доброй ночи!"
If Val(inputWindow.Text) > 24 Then result.Caption = "Доп. интервал [0:24]"
End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub inputWindow_Change()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub result_Click()

End Sub

Private Sub UserForm_Click()

End Sub
