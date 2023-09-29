VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MapHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class: MapHandler
'MapHandler class allows us to create a collection of maps
'However, this handler specifically allows us to switch through maps (walking between maps)


Option Explicit
Option Base 1

' Properties -------------------------------------------------------------
Private mapArr() As Map
Private mapIndex As Integer

' Constructor ------------------------------------------------------------
Function Init(numMaps As Integer)
    ReDim Preserve mapArr(numMaps)

End Function

' Methods ----------------------------------------------------------------

' Add Map to the map array
Function addMap(index As Integer, path As String, sheet As Worksheet, w As Integer, h As Integer, mTiles As Variant, _
                mapUID As String, Optional tilesWsName As String)
    Set mapArr(index) = CreateMap(path, sheet, w, h, mTiles, mapUID, tilesWsName)
End Function

'Delete all maps part of the handler
Function removeAllMaps()
    Dim i As Integer

    For i = LBound(mapArr) To UBound(mapArr) Step 1
        mapArr(i).Remove
    Next
End Function

'Load a specific map
Function loadMap(index As Integer)
    mapArr(index).loadMap
End Function

'Load all maps
Function LoadAllMaps()
    Dim m As Variant
    For Each m In mapArr
        m.loadMap
    Next
End Function

' Switch between maps based on a specific index passed through
Function mapSwitch(index As Integer)
    Dim i As Integer

    For i = LBound(mapArr) To UBound(mapArr) Step 1
        If i = index Then
            mapArr(index).SetVisible (True)
        Else
            mapArr(i).SetVisible (False)
        End If
    Next

    'Loop through all entities. What entities are on map(index)
    'Based on which entities are on the map, set the ones that are visible
    'Set the entities that are not on the map being switched to invisible
    Dim clsEntity As Entity
    For Each clsEntity In entityCollection
        If clsEntity.getBeaten = False Then
            If (clsEntity.getMap.getUID = mapArr(index).getUID) Then
                'This means entity is on the map that we are switching to
                clsEntity.SetCharacterVisible (True)
                clsEntity.SetCharacterFrameVisible (clsEntity.getDirection)
            Else
                clsEntity.SetCharacterVisible (False)
            End If
        End If
    Next
End Function

' Getters and Setters ----------------------------------------------------------------
Property Get getMapIndex() As Integer
    getMapIndex = mapIndex
End Property
Property Let setMapIndex(index As Integer)
    mapIndex = index
End Property

Function getMap(index As Integer) As Variant
    Set getMap = mapArr(index)
End Function

'Get a map from a unique ID
Function getMapFromUID(UID As String) As Variant
    Dim m As Variant
    For Each m In mapArr
        If m.getUID = UID Then
            Set getMapFromUID = m
            Exit Function
        End If
    Next
End Function

Function getNumOfMaps() As Integer
    getNumOfMaps = UBound(mapArr)
End Function


