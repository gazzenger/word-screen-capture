VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ScreenCapture v1"
   ClientHeight    =   1270
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   3390
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    'Minimise the userform
    Dim ret
    ret = ShowWindow(UserFormHdl, 0)
    'taken from https://stackoverflow.com/a/15952009
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    wsh.Run "snippingtool /clip", windowStyle, waitOnReturn

    If (CheckBox1.Value) Then
        Dim myValue As Variant
        myValue = InputBox("Please provide a description")
        Selection.TypeText Text:=myValue & vbNewLine
    End If
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.TypeText Text:=vbNewLine
    ret = ShowWindow(UserFormHdl, 1)
End Sub

Private Sub UserForm_Activate()
    Call Keep_Form_On_Top_64bit.KeepFormOnTop
End Sub

Private Sub UserForm_Initialize()
    Keep_Form_On_Top_64bit.xlHwnd = Keep_Form_On_Top_64bit.WordHdl
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim ret As Long
    ret = ShowWindow(xlHwnd, 1)
End Sub

