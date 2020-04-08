Attribute VB_Name = "Keep_Form_On_Top_64bit"
'Majority taken from http://www.vbaexpress.com/forum/showthread.php?58189-Make-userform-stay-on-top-of-all-windows-when-macro-is-fired

' Written: July 07, 2009
' Author:  Leith Ross
' Summary: Keeps a UserForm, or any window on top of all other windows.
'          Call this macro from the UserForm_Activate event code module.

' This variable is initalized in the UserForm_Initialize() event. It holds the hWnd to the Excel Application Window.
Global xlHwnd As LongPtr
'Future Variable for storing the ActiveDocument
Global objDoc As Document
'Word Handle
Global WordHdl As LongPtr
'Userform Handle
Global UserFormHdl As LongPtr

' This API call is used to hide or show the Excel Application.
Public Declare PtrSafe Function ShowWindow _
    Lib "user32.dll" _
        (ByVal hwnd As LongPtr, _
         ByVal nCmdShow As Long) _
    As Long

' Returns the Window Handle of the Window that is accepting User input.
Private Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As LongPtr

Private Declare PtrSafe Function SetWindowPos _
    Lib "user32.dll" _
        (ByVal hwnd As LongPtr, _
         ByVal hWndInsertAfter As LongPtr, _
         ByVal x As Long, _
         ByVal Y As Long, _
         ByVal cx As Long, _
         ByVal cy As Long, _
         ByVal wFlags As Long) _
    As Long

Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Sub KeepFormOnTop()
    Dim ret As Long
    Const HWND_TOPMOST  As Long = -1
    Const SWP_NOMOVE    As Long = &H2
    Const SWP_NOSIZE    As Long = &H1
    
    'Userform Handle
    UserFormHdl = GetForegroundWindow()
    ret = ShowWindow(xlHwnd, 0)
    ret = SetWindowPos(GetForegroundWindow(), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
End Sub



Sub Run()
    'Handle for Windows Form
    WordHdl = GetForegroundWindow()
    
    'Get Current Word Document
    Set objDoc = Word.ActiveDocument
    
    'check if the file has been saved anywhere, if not then save as dialog
    If ActiveDocument.Path = "" Then
        Dialogs(wdDialogFileSaveAs).Show
    End If
    'if document has unsaved changes, then save them
    If ActiveDocument.Saved = False Then
        ActiveDocument.Save
    End If
    
    'Show Userform
    UserForm1.Show
End Sub


