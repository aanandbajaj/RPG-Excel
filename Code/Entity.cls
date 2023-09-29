VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class: Entity
'This object represents an 'entity' which has various attributes
'These attributes include sprite (picture), direction (which way is the entity facing)
'What speed they are moving at, what their health is, what their base attack is, and so on
'Since we have many entities in this game, we decided to make an entity class to reduce the lines of code
'Methods:
'Init: Initializes the entity
'Move: Depending on which way the entity is facing, set the sprite
'Update: Method called in the game loop, constantly updates entity's position and acts as the "draw"
'function in other programming languages (sets the location of the sprite in the game)
'MoveLeft, MoveRight, MoveUp, MoveDown: sets entity position, checks for collision
'IsCollision: checks for collision with walls, enemies, doors, etc. - only for main entity (user), since
'There is no movement allowed for enemies
'InitializeCharacterImages: Initially sets the sprites, puts them on top of each other
'SetCharacterFrameVisible: sets a specific frame visible in the pictures array
'Can emulate animation
'SetCollisionOnMap: based on the entity's current location, sets the collision on the map, so that
'entities cannot move through each other
'clearCollisionOnMap: clears the collision when the entity dies

Option Explicit
Option Base 1

' Constants for character direction facing
Const LEFT_FACE As Integer = 1
Const RIGHT_FACE As Integer = 2
Const UP_FACE As Integer = 3
Const DOWN_FACE As Integer = 4

' Speed to move character one map tile at a time
Private Const SPEED As Integer = TILE_SIZE

Private xPos As Integer
Private yPos As Integer
Private xTile As Integer
Private yTile As Integer
Private nextXTile As Integer
Private nextYTile As Integer
Private health As Integer
Private baseHealth As Integer
Private entityWidth As Integer
Private entityHeight As Integer
Private numCoins As Long
Private baseDamage As Integer

'Can be 0, 1, 2, 3
Private direction As Integer

'Make array and group pictures
Dim charImg() As Object

'Group of objects
Dim characterGroup As Object

Dim frames As Integer

'The counter which will help us rotate through pictures
Private animCounter As Integer
Private entityID As String
Private entityName As String

Private isMoving As Boolean
Private currentMessage As String
Private currentMap As Map
Private collidingWith As String
Private entityUID As Integer
Private coinsDropped As Integer

'For enemies
'If False, don't start battle again
Private beaten As Boolean

' Constructor -------------------------------------------------------------

'This method accepts the arguments to send to the constructor
'Citation: https://stackoverflow.com/questions/15224113/pass-arguments-to-constructor-in-vba
'There will be an InitiateProperties method in each class module
Public Sub Init(ID As String, X As Integer, Y As Integer, dir As Integer, sheet As Worksheet, m As Map, _
                healthPercentage As Double, entityBaseHealth As Integer, baseDmg As Integer, _
                Optional beat As Boolean, Optional name As String, _
                Optional width As Variant, Optional height As Variant, Optional coinDrop As Integer)

'The character is intialized using tiles in x and y
    xTile = X
    yTile = Y

    'Therefore, the specific pixels on the screen must be set by multiplying the tile by the TILE_SIZE due to grid system
    xPos = (xTile - 1) * TILE_SIZE
    yPos = (yTile - 1) * TILE_SIZE

    'Managing optional variables
    If IsMissing(width) Then
        entityWidth = TILE_SIZE
    Else
        entityWidth = width
    End If

    If IsMissing(height) Then
        entityHeight = TILE_SIZE
    Else
        entityHeight = height
    End If

    nextXTile = 0
    nextYTile = 0

    direction = dir
    numCoins = 0
    coinsDropped = coinDrop

    Set currentMap = m
    beaten = beat
    baseHealth = entityBaseHealth
    health = healthPercentage * baseHealth
    entityName = name
    entityID = ID
    UIDCounter = UIDCounter + 1
    entityUID = UIDCounter
    baseDamage = baseDmg

    animCounter = 0

    'Add default image
    'From Stackoverflow - change later
    sheet.Activate

    ReDim Preserve charImg(4)
    Call InitializeCharacterImages(wbPath, sheet)
    Call Me.move
    Call Me.setCollisionOnMap

End Sub

' Methods -------------------------------------------------------------

'Moving function
Public Function move()
    Select Case getDirection
    Case LEFT_FACE
        SetCharacterFrameVisible (LEFT_FACE)
    Case RIGHT_FACE
        SetCharacterFrameVisible (RIGHT_FACE)
    Case UP_FACE
        SetCharacterFrameVisible (UP_FACE)
    Case DOWN_FACE
        SetCharacterFrameVisible (DOWN_FACE)
    End Select
End Function

'Called in Run Function in Main Module
Public Function Update()
    Me.move
    characterGroup.Left = xPos
    characterGroup.Top = yPos
    xTile = Round(xPos / 32, 0) + 1
    yTile = Round(yPos / 32, 0) + 1
End Function

'Functions that allow user to move the character
Public Function MoveLeft()
    Dim nextX As Integer
    Dim nextY As Integer

    nextX = xPos - SPEED
    nextY = yPos

    nextXTile = Round(nextX / 32, 0) + 1
    nextYTile = Round(nextY / 32, 0) + 1

    'Add collision detection
    If IsCollision = False Then
        'Moving left
        xPos = xPos - SPEED

        'Not moving up and down
        yPos = yPos
    End If
End Function

Public Function MoveRight()
    Dim nextX As Integer
    Dim nextY As Integer

    nextX = xPos + SPEED
    nextY = yPos

    nextXTile = Round(nextX / 32, 0) + 1
    nextYTile = Round(nextY / 32, 0) + 1

    'Add collision detection
    If IsCollision = False Then
        'Moving left
        xPos = xPos + SPEED

        'Not moving up and down
        yPos = yPos
    End If
End Function

Public Function MoveUp()
    Dim nextX As Integer
    Dim nextY As Integer

    nextX = xPos
    nextY = yPos - SPEED

    nextXTile = Round(nextX / 32, 0) + 1
    nextYTile = Round(nextY / 32, 0) + 1

    'Add collision detection
    If IsCollision = False Then
        'Moving left
        xPos = xPos

        'Not moving up and down
        yPos = yPos - SPEED
    End If

End Function

Public Function MoveDown()
    Dim nextX As Integer
    Dim nextY As Integer

    nextX = xPos
    nextY = yPos + SPEED

    nextXTile = Round(nextX / 32, 0) + 1
    nextYTile = Round(nextY / 32, 0) + 1

    'Add collision detection
    If IsCollision = False Then
        xPos = xPos

        'Moving down
        yPos = yPos + SPEED
    End If

End Function

'COLLISION DETECTION
'Runs everytime main entity moves
Private Function IsCollision() As Boolean
'Collision only for moving character
    If entityID = MAIN_CHAR Then

        Dim tiles() As Variant
        Dim currentTile As String
        tiles = currentMap.getTiles
        currentTile = tiles(nextYTile, nextXTile)

        Select Case currentTile

        Case Is = ""
            IsCollision = False
            collidingWith = ""
            Exit Function
        Case Is = "wall"
            collidingWith = "wall"
            IsCollision = True
            Exit Function
        Case Is = "enemy"
            collidingWith = "wall"
            IsCollision = True
            Exit Function
        Case Is = "door1"
            collidingWith = "door"
            IsCollision = False
            Call doorSoundEffect

            Set currentMap = gameMapHandler.getMap(2)
            gameMapHandler.mapSwitch (2)
            mainGameDialog.SetDialogBoxText ("")
            dialogCounter = 0
            mainGameDialog.SetDialogBoxText ("Dungeon Level 1!")
            Me.setXTile (19)
            Me.setYTile (2)
            Exit Function
            'door to dungeon1 scene
        Case Is = "door2"
            collidingWith = "door"
            Call doorSoundEffect

            Set currentMap = gameMapHandler.getMap(1)
            gameMapHandler.mapSwitch (1)
            mainGameDialog.SetDialogBoxText ("")
            dialogCounter = 0
            mainGameDialog.SetDialogBoxText ("Main Room. Buy potions, weapons, or armour from the shop.")
            Me.setXTile (5)
            Me.setYTile (5)
            Me.setDirection = 4
            IsCollision = False
            Exit Function
            'door back to main
        Case Is = "door3"
            collidingWith = "door"
            IsCollision = False
            Call doorSoundEffect
            Set currentMap = gameMapHandler.getMap(3)
            gameMapHandler.mapSwitch (3)
            mainGameDialog.SetDialogBoxText ("")
            dialogCounter = 0
            mainGameDialog.SetDialogBoxText ("Dungeon Level 2!")
            Me.setXTile (19)
            Me.setYTile (2)
            Exit Function
            'door to dungeon2
        Case Is = "door4"
            collidingWith = "door"
            IsCollision = False
            Call doorSoundEffect
            Set currentMap = gameMapHandler.getMap(4)
            gameMapHandler.mapSwitch (4)
            mainGameDialog.SetDialogBoxText ("")
            dialogCounter = 0
            mainGameDialog.SetDialogBoxText ("Dungeon Level 3!")
            Me.setXTile (19)
            Me.setYTile (2)
            Exit Function
            'door to dungeon3
            Exit Function
        Case Is = "shop"
            collidingWith = "shop"
            IsCollision = True
            mainGameDialog.SetDialogBoxText ("")
            dialogCounter = 0
            mainGameDialog.SetDialogBoxText ("Welcome To the Shop! Click 'Interact' ")
            Exit Function
        End Select

        'Collisions with Enemies
        If InStr(currentTile, BOSS_CHAR) <> 0 Then
            collidingWith = Replace(currentTile, BOSS_CHAR, "", 1, 1)
            mainGameDialog.SetDialogBoxText ("")
            dialogCounter = 0
            mainGameDialog.SetDialogBoxText ("You have encountered a " & BOSS_CHAR & ". Click 'Interact' to Battle")
            IsCollision = True
            Exit Function
        ElseIf InStr(currentTile, SKELETON_CHAR) <> 0 Then
            collidingWith = Replace(currentTile, SKELETON_CHAR, "", 1, 1)
            mainGameDialog.SetDialogBoxText ("")
            dialogCounter = 0
            mainGameDialog.SetDialogBoxText ("You have encountered a " & SKELETON_CHAR & ". Click 'Interact' to Battle")
            IsCollision = True
            Exit Function
        ElseIf InStr(currentTile, BAT_CHAR) <> 0 Then
            collidingWith = Replace(currentTile, BAT_CHAR, "", 1, 1)
            mainGameDialog.SetDialogBoxText ("")
            dialogCounter = 0
            mainGameDialog.SetDialogBoxText ("You have encountered a " & BAT_CHAR & ". Click 'Interact' to Battle")
            IsCollision = True
            Exit Function
        ElseIf InStr(currentTile, GOBLIN_CHAR) <> 0 Then
            collidingWith = Replace(currentTile, GOBLIN_CHAR, "", 1, 1)
            mainGameDialog.SetDialogBoxText ("")
            dialogCounter = 0
            mainGameDialog.SetDialogBoxText ("You have encountered a " & GOBLIN_CHAR & ". Click 'Interact' to Battle")
            IsCollision = True
            Exit Function
        End If

    Else
        IsCollision = False
    End If

End Function


'Got help from: https://www.spreadsheet1.com/how-to-group-shapes.html
'This function gets images from a folder and puts them into an array of images which can be cycled through
Private Function InitializeCharacterImages(ByVal wbPath As String, sheet As Worksheet)

'Load all images
    Dim strState As String
    Dim shp As Shape
    Dim X As Variant

    Set charImg(1) = sheet.Shapes.AddPicture(wbPath & "\Images\Character Sprites\" & entityID & "\SideFace.png", msoFalse, msoTrue, xPos, yPos, entityWidth, entityHeight)
    Set charImg(2) = sheet.Shapes.AddPicture(wbPath & "\Images\Character Sprites\" & entityID & "\SideFace.png", msoFalse, msoTrue, xPos, yPos, entityWidth, entityHeight)
    Set charImg(3) = sheet.Shapes.AddPicture(wbPath & "\Images\Character Sprites\" & entityID & "\BackFace.png", msoFalse, msoTrue, xPos, yPos, entityWidth, entityHeight)
    Set charImg(4) = sheet.Shapes.AddPicture(wbPath & "\Images\Character Sprites\" & entityID & "\FrontFace.png", msoFalse, msoTrue, xPos, yPos, entityWidth, entityHeight)

    'Facing left and right
    charImg(2).Flip (msoFlipHorizontal)

    'Group all Objects
    Set characterGroup = sheet.Shapes.Range(Array(charImg(1).name, charImg(2).name, charImg(3).name, charImg(4).name)).Group

End Function

'This function allows us to switch to a specific frame for the sprite
'This is generalized so we can actually emulate animation
Public Function SetCharacterFrameVisible(visibleIndex As Integer)
    Dim i As Integer

    'If passed 0 through this function, want to set entire sprite to invisible
    For i = LBound(charImg) To UBound(charImg) Step 1
        If visibleIndex >= 1 Then
            If i = visibleIndex Then
                charImg(visibleIndex).Visible = True
            Else
                charImg(i).Visible = False
            End If
        End If
    Next

End Function

'Set the entire character invisible
'Helps when switching through maps
Public Function SetCharacterVisible(visibility As Boolean)
    characterGroup.Visible = visibility
End Function

'Deletes the entity shape
'Does not delete the object instance itself
Public Function Remove()
    Dim shp As Variant
    For Each shp In charImg
        shp.Delete
    Next
End Function

'Based on the entity's current location, sets a cell in the collision worksheets (tiles worksheet)
'This is helpful because when the user is moving around in a map, don't want to be able to go through
'enemies
Public Function setCollisionOnMap()
    Dim ws As Worksheet

    If entityName = "Enemy" Then
        Set ws = Worksheets(currentMap.getTilesWsName)

        'if double normal size
        If entityWidth / TILE_SIZE = 2 And entityHeight / TILE_SIZE = 2 Then
            ws.Cells(yTile + 1, xTile) = entityID & entityUID
            ws.Cells(yTile + 1, xTile + 1) = entityID & entityUID
            'Normal size
        ElseIf entityWidth / TILE_SIZE = 1 And entityHeight / TILE_SIZE = 1 Then
            ws.Cells(yTile, xTile) = entityID & entityUID
        End If
    End If
End Function

'Clears the cell which was filled in setCollisionOnMap Function
Public Function clearCollisionOnMap()
    Dim ws As Worksheet

    If entityName = "Enemy" Then
        Set ws = Worksheets(currentMap.getTilesWsName)
        'if double normal size
        If entityWidth / TILE_SIZE = 2 And entityHeight / TILE_SIZE = 2 Then
            ws.Cells(yTile + 1, xTile).ClearContents
            ws.Cells(yTile + 1, xTile + 1).ClearContents

            'Normal size
        ElseIf entityWidth / TILE_SIZE = 1 And entityHeight / TILE_SIZE = 1 Then
            ws.Cells(yTile, xTile).ClearContents
        End If
    End If
End Function

Public Function updateMap()
    Dim m As Map

    Set m = gameMapHandler.getMapFromUID(currentMap.getUID)
    currentMap.setTiles (m.getTiles)
End Function

' Getters and Setters -------------------------------------------------------------

Property Get getX() As Integer
    getX = xPos
End Property

Property Get getY() As Integer
    getY = yPos
End Property

Property Get getXTile() As Integer
    getXTile = xTile
End Property

Property Get getYTile() As Integer
    getYTile = yTile
End Property

Property Get getHealth() As Integer
    getHealth = health
End Property

Property Get getBaseHealth() As Integer
    getBaseHealth = baseHealth
End Property

Property Get getEntityID() As String
    getEntityID = entityID
End Property

Property Get getCoinsDrop() As Integer
    getCoinsDrop = coinsDropped
End Property

Property Get getEntityUID() As String
    getEntityUID = entityUID
End Property

Property Get getEntityName() As String
    getEntityName = entityName
End Property

Property Get getBeaten() As Boolean
    getBeaten = beaten
End Property

Property Get getBaseDmg() As Integer
    getBaseDmg = baseDamage
End Property

Property Get getCollidingWith() As String
    getCollidingWith = collidingWith
End Property

Public Sub setXTile(xT As Integer)
    xPos = xT * TILE_SIZE
End Sub

Public Sub setYTile(yT As Integer)
    yPos = yT * TILE_SIZE
End Sub

Public Property Get getDirection() As Integer
    getDirection = direction
End Property

Public Property Let setDirection(directionValue As Integer)
    direction = directionValue
End Property

Public Property Let setBaseDmg(dmg As Integer)
    baseDamage = dmg
End Property

Public Sub setHealth(h As Integer)
    health = h
End Sub

Public Sub setBaseHealth(h As Integer)
    baseHealth = h
End Sub

Public Sub setBeaten(b As Boolean)
    beaten = b
End Sub

'Used to update currentMap (if tiles get updated, is useful)
Public Function setCurrentMap(m As Map)
    Set currentMap = m
End Function

Public Function getMap() As Variant
    Set getMap = currentMap
End Function

Public Sub setCoins(coins As Long)
    numCoins = coins
End Sub

Public Property Get getCoins() As Long
    getCoins = numCoins
End Property





