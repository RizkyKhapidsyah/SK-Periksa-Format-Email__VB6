VERSION 5.00
Begin VB.Form FormPeriksaAlamatEmail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Periksa Alamat Email"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPeriksa 
      Caption         =   "Periksa"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox textEmail 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "FormPeriksaAlamatEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function IsValidEmail(sEMail As String) As Boolean
    Dim sInvalidChars As String
    Dim bTemp As Boolean
    Dim i As Integer
    Dim sTemp As String

    sInvalidChars = "!#$%^&*()=+{}[]|\;:'/?>,< "

    bTemp = InStr(sEMail, "@") <= 0
    If bTemp Then GoTo exit_function

    bTemp = InStr(sEMail, ".") <= 0
    If bTemp Then GoTo exit_function

    bTemp = Len(sEMail) < 6
    If bTemp Then GoTo exit_function

    i = InStr(sEMail, "@")
    sTemp = Mid(sEMail, i + 1)
    bTemp = InStr(sTemp, "@") > 0
    
    If bTemp Then GoTo exit_function
    bTemp = InStr(sTemp, " ") > 0
    If bTemp Then GoTo exit_function

    bTemp = InStr(sTemp, ".") = 0
    If bTemp Then GoTo exit_function
    
    bTemp = InStr(sEMail, Chr(34)) > 0
    If bTemp Then GoTo exit_function
    
    If Len(sEMail) > Len(sInvalidChars) Then
        For i = 1 To Len(sInvalidChars)
            If InStr(sEMail, Mid(sInvalidChars, i, 1)) > 0 _
                  Then bTemp = True
            If bTemp Then Exit For
        Next
    Else
        For i = 1 To Len(sEMail)
            If InStr(sInvalidChars, Mid(sEMail, i, 1)) > 0 _
                   Then bTemp = True
            If bTemp Then Exit For
        Next
    End If
    If bTemp Then GoTo exit_function
    
    bTemp = InStr(sEMail, "..") > 0
    If bTemp Then GoTo exit_function
    
exit_function:
    IsValidEmail = Not bTemp
End Function



Private Sub cmdPeriksa_Click()
    If IsValidEmail(textEmail.Text) = True Then
        MsgBox ("Format Email sudah Benar."), vbOKOnly, "Hasil"
    Else
        MsgBox ("Format Email sudah Salah."), vbOKOnly, "Hasil"
    End If
End Sub

