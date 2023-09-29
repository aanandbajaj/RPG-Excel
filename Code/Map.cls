VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: Map - defines map objects for the game and can display them

' Properties --------------------------------------------------------------
Option Explicit
Option Base 1

Private mapPath As String
Private width As Integer
Private height As Integer
Private tiles() As Variant
Private mapUID As String

'EntityNames is to store which entities are in this current map
Private entityNames() As String
Private mapShp As Object
Private collisionWsName

' Constructor -------------------------------------------------------------
'The constructor takes in:
'path - directory to load map image
'sheet - where to place the map
'w,h - width and height
'mTiles - associated 2D array of collision tiles for the map
'ID - Unique ID for the map for easy referencing

Function Init(path As String, sheet As Worksheet, w As Integer, h As Integer, mTiles As Variant, _
              ID As String, Optional tilesWsName As String)

    mapPath = path
    width = w
    height = h
    tiles = mTiles
    mapUID = ID
    collisionWsName = tilesWsName

End Function

' Methods -----------------------------------------------------------------

Public Function SetVisible(vis As Boolean)
    mapShp.Visible = vis
End Function

' LoadMap sub to display the map
Public Function loadMap()

    Set mapShp = Worksheets("Main").Shapes.AddPicture(mapPath, msoFalse, msoCTrue, 0, 0, width, height)

    'Clear the associated collision tiles worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(collisionWsName)
    Dim rng As Range
    Set rng = ws.Range("map")

    Dim cell As Object
    'If the tiles have enemy collisions on them, clear the contents of the cells
    For Each cell In rng
        If (InStr(cell, GOBLIN_CHAR) <> 0) Or (InStr(cell, BAT_CHAR) <> 0) Or (InStr(cell, SKELETON_CHAR) <> 0) Then
            cell.ClearContents
        End If
    Next

End Function

'Delete map
Public Function Remove()
    mapShp.Delete
End Function

' Getters and Setters -----------------------------------------------------------------
Property Get getMapPath()
    getMapPath = mapPath
End Property
Property Let setMapPath(path As String)
    mapPath = path
End Property

Property Get getTiles() As Variant
    getTiles = tiles
End Property

Public Sub setTiles(mTiles As Variant)
    tiles = mTiles
End Sub

Property Get getTilesWsName() As String
    getTilesWsName = collisionWsName
End Property

Property Let addEntity(name As String)
    Dim size As Integer
    size = UBound(entityNames)
    entityNames(size + 1) = name
End Property

Property Get getUID() As String
    getUID = mapUID
End Property

'Get a specific tile from the associated collision ws
Public Function GetTile(xTile As Integer, yTile As Integer) As String
    GetTile = tiles(xTile, yTile)
    Exit Function
End Function



