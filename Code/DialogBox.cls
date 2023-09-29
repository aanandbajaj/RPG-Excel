VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DialogBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class: DialogBox
'This object was created to allow for easy manipulation of a custom dialog box for the RPG game

'Declaring private variables
Private dialog As Object
Private dialogImg As Object
Private text As String
Private width As Integer
Private height As Integer
Private xPos As Integer
Private yPos As Integer
Private filePath As String

'Constructor
Function Init(path As String, sheet As Worksheet, X As Integer, Y As Integer, w As Integer, h As Integer, dialogText As String)
    filePath = path
    width = w
    height = h
    xPos = X
    yPos = Y
    text = dialogText

    Call AddDialogBox(sheet)
End Function

'https://social.msdn.microsoft.com/Forums/office/en-US/100e059f-8943-4299-ba2a-c5ccdca08acd/adding-picture-to-an-excel-textbox-using-vba-code?forum=exceldev
'Adding a dialog box to a specific worksheet
Private Sub AddDialogBox(ws As Worksheet)
    Set dialogImg = ws.Shapes.AddPicture(filePath, msoFalse, msoTrue, xPos, yPos, width, height)
    Set dialog = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, width, height)
    dialog.TextFrame.Characters.text = text

    'Ran a macro to understand how fonts work for textboxes in VBA
    'Setting a
    With dialog.TextFrame2.TextRange.Font
        .NameComplexScript = "Lucida Console"
        .NameFarEast = "Lucida Console"
        .name = "Lucida Console"
    End With

    'Formatting dialog box
    dialog.Fill.Transparency = 1

    dialog.TextFrame2.TextRange.Font.size = 20
    dialog.TextFrame2.TextRange.Font.Bold = msoTrue

End Sub

'Setting the text of the dialog box
Public Sub SetDialogBoxText(txt As String)
    text = txt
    dialog.TextFrame.Characters.text = text
End Sub

Public Function GetDialogBoxText() As String
    GetDialogBoxText = text
End Function

'Removing dialog box
Public Sub RemoveDialogBox()
    dialog.Delete
    dialogImg.Delete
End Sub
