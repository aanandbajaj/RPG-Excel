Attribute VB_Name = "Factory"
'Module: Factory
'The factory module will include instantiation functions
'The functions will call the initiation subroutine "Init" in each Class Module and return the initiated object
'Allows for easy "creation" of objects and emulates OOP

'Create an entity
Public Function CreateEntity(ID As String, X As Integer, Y As Integer, direction As Integer, sheet As Worksheet, _
                             m As Map, healthPercentage As Double, baseHealth As Integer, baseDmg As Integer, _
                             Optional beaten As Boolean, Optional name As String, _
                             Optional width As Variant, Optional height As Variant, Optional coinDrop As Integer) As Entity
    Dim entityObj As Entity
    Set entityObj = New Entity

    entityObj.Init ID, X, Y, direction, sheet, m, healthPercentage, baseHealth, baseDmg, beaten, name, width, height, coinDrop
    Set CreateEntity = entityObj
End Function

'Create map handler
Public Function CreateMapHandler(numMaps As Integer) As MapHandler

    Dim newMapArr As New MapHandler

    newMapArr.Init (numMaps)
    Set CreateMapHandler = newMapArr

End Function

'Create map
Public Function CreateMap(path As String, sheet As Worksheet, width As Integer, height As Integer, tiles As Variant, _
                          mapUID As String, Optional tilesWsName As String) As Map

    Dim newMap As New Map

    newMap.Init path, sheet, width, height, tiles, mapUID, tilesWsName
    Set CreateMap = newMap

End Function

'Create a dialog box
Public Function CreateDialogBox(path As String, sheet As Worksheet, X As Integer, Y As Integer, w As Integer, h As Integer, dialogText As String) As DialogBox

    Dim newDialog As New DialogBox

    newDialog.Init path, sheet, X, Y, w, h, dialogText
    Set CreateDialogBox = newDialog

End Function

