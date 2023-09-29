Attribute VB_Name = "Main_Module"
'-----------------------------------------------------------
'--------------------RPG DUNGEON QUEST----------------------
'-----------------------------------------------------------
'Class: End User Modelling
'University: Richard Ivey
'Project Description: This game allows the user to enter into a world of monsters in excel
'With beautiful graphics, the user can fight enemies, build up their coin stack
'and buy better weapons/armour. The objective of the game is to beat all the enemies
'on the map. ENJOY!

'Module: Main_Module
'Main: Runs when user clicks "Start" on start worksheet
'Run: Game loop, constantly updates all entities every tick
'This module handles all major button clicks (subs assigned to buttons in main worksheet)
'This module also includes various useful functions such as setCollisions
'This module handles setting up the scenes/maps/enemy locations and also checks whether user has won/lost

Option Explicit
Option Base 1

'Sprite tile size
Public Const TILE_SIZE = 32

'Size of game
Public Const GAME_WIDTH As Integer = 640
Public Const GAME_HEIGHT As Integer = 320

'Initial base damage for the user
Public Const INITIAL_BASE_DMG_USER = 30

'Number of enemies per map
Public Const NUM_OF_ENEMIES = 8
Public Const HPBarOriginalWidth = 130

'Folder Names to Extract Sprites
Public Const MAIN_CHAR = "Main"
Public Const BOSS_CHAR = "Boss"
Public Const SKELETON_CHAR = "Skeleton"
Public Const BAT_CHAR = "Bat"
Public Const GOBLIN_CHAR = "Goblin"

'Tiered Health
Public Const ENEMYHEALTH_LEVEL_1 = 250
Public Const ENEMYHEALTH_LEVEL_2 = 600
Public Const ENEMYHEALTH_LEVEL_3 = 1000

'Tiered Base Dmg
Public Const BASEDMG_LEVEL_1 = 30
Public Const BASEDMG_LEVEL_2 = 50
Public Const BASEDMG_LEVEL_3 = 70

'Tiered Coin Drop
Public Const COINDROP_LEVEL_1 = 600
Public Const COINDROP_LEVEL_2 = 700
Public Const COINDROP_LEVEL_3 = 800

'Declaring entities
'There will be an array of entities per dungeon
'All the entities will be housed in a collection object
Public hero As Entity
Dim enemiesDungeon1() As Entity
Dim enemiesDungeon2() As Entity
Dim enemiesDungeon3() As Entity

'Collection of entity objects for easy manipulation
'The collection allows us to access all entities
Public entityCollection As Collection

'Declaring Tile Worksheets
Dim mainMapTilesWs As Worksheet
Dim dungeonTilesWs1 As Worksheet
Dim dungeonTilesWs2 As Worksheet
Dim dungeonTilesWs3 As Worksheet
Dim bossMapTilesWs As Worksheet

'Main Game Map worksheet
Dim mainMapWs As Worksheet

'Declaring variants which will hold the 2D array from tiles worksheets
Dim mainMapTiles() As Variant
Dim dungeonMapTiles1() As Variant
Dim dungeonMapTiles2() As Variant
Dim dungeonMapTiles3() As Variant

'Game map handler will manage all maps (allow for seamless switching, and so on)
Public gameMapHandler As MapHandler

'Counter which stores all the entities' unique ids
Public UIDCounter As Integer

'Other variables. Most are public because they need to be accessed from other modules/class modules
Public coinText As Variant
Public wbPath As String
Private rep_count As Integer
Public dialogCounter As Integer
Public mainGameDialog As DialogBox

Public Sub Main()
'Show loading screen
    LoadingUI.Show (0)

    'Initialize Variables/Objects
    purchasedWeapon = False
    wbPath = ActiveWorkbook.path
    UIDCounter = 0
    Set entityCollection = New Collection
    'Game worksheet and activate
    Set mainMapWs = ThisWorkbook.Worksheets("Main")
    Set coinText = mainMapWs.Shapes("CoinText")
    Set mainGameDialog = Factory.CreateDialogBox(wbPath & "\Images\Dialog Box\DialogBackground.png", mainMapWs, 0, 320, 640, 96, "")
    dialogCounter = 0

    'Setting worksheets and tiles into arrays
    Set mainMapTilesWs = ThisWorkbook.Worksheets("Main_Scene_Tiles")

    '3 Levels of Dungeons
    'There are three different worksheets because in each, there will be different enemies in different locations
    Set dungeonTilesWs1 = ThisWorkbook.Worksheets("Dungeon_Scene_Tiles_1")
    Set dungeonTilesWs2 = ThisWorkbook.Worksheets("Dungeon_Scene_Tiles_2")
    Set dungeonTilesWs3 = ThisWorkbook.Worksheets("Dungeon_Scene_Tiles_3")

    'Getting the data from the worksheet and putting it into a 2D array
    'The 2D array will have x,y coordinates representing the grid in the worksheet
    '2D array is faster to loop through/manipulate than worksheet cells for detecting collision
    mainMapTiles = mainMapTilesWs.Range("Map").Value
    dungeonMapTiles1 = dungeonTilesWs1.Range("Map").Value
    dungeonMapTiles2 = dungeonTilesWs2.Range("Map").Value
    dungeonMapTiles3 = dungeonTilesWs3.Range("Map").Value

    '4 total maps
    'Creating the maps by calling the .addMap function in the map handler class module
    Set gameMapHandler = CreateMapHandler(4)
    Call gameMapHandler.addMap(1, wbPath & "\Images\Maps\Main_Map_Scene_Door.png", mainMapWs, GAME_WIDTH, GAME_HEIGHT, mainMapTiles, "MAIN_ROOM", mainMapTilesWs.name)
    Call gameMapHandler.addMap(2, wbPath & "\Images\Maps\Dungeon Maps\map_1.png", mainMapWs, GAME_WIDTH, GAME_HEIGHT, dungeonMapTiles1, "DUNGEON", dungeonTilesWs1.name)
    Call gameMapHandler.addMap(3, wbPath & "\Images\Maps\Dungeon Maps\map_2.png", mainMapWs, GAME_WIDTH, GAME_HEIGHT, dungeonMapTiles2, "DUNGEON2", dungeonTilesWs2.name)
    Call gameMapHandler.addMap(4, wbPath & "\Images\Maps\Dungeon Maps\map_3.png", mainMapWs, GAME_WIDTH, GAME_HEIGHT, dungeonMapTiles3, "DUNGEON3", dungeonTilesWs3.name)

    'Load every map (set all visible)
    gameMapHandler.LoadAllMaps

    'Creating main hero character
    Set hero = Factory.CreateEntity(MAIN_CHAR, 4, 4, 1, mainMapWs, gameMapHandler.getMap(1), 1, 100, INITIAL_BASE_DMG_USER, False, "Hero")
    entityCollection.Add hero

    'Creating the enemies based on random positions in dungeon 1, 2, and 3
    Call setupEnemies

    'Setting up collisions so the user cannot "move through" walls and enemies
    Call setCollisions

    'Set the initial map (main area) visible, and all the other ones false
    'Map switcher handles this task
    gameMapHandler.mapSwitch (1)

    'Set the coin textbox on the controls UI to the initial hero character's coins (0)
    coinText.TextFrame.Characters.text = hero.getCoins

    'Show main worksheet and start the game loop
    mainMapWs.Activate
    Call WaitASec(2)
    LoadingUI.Hide
    Unload LoadingUI
    
    'Formatting sheet
    
    'mainMapWs.Range("A1:S30").Select
    'ActiveWindow.Zoom = True
    Call Run

End Sub

'Emulates Game Loop
Public Sub Run()

    Do
        DoEvents

        'Constantly Update coinText on WS
        
        coinText.TextFrame.Characters.text = hero.getCoins
        ShopUI.ShopCoinsLabel.Caption = hero.getCoins
        InventoryUI.ShopCoinsLabel.Caption = hero.getCoins

        'Clearing dialog box text with a 'wait' time
        If dialogCounter <= 10 Then
            If mainGameDialog.GetDialogBoxText <> "" Then
                dialogCounter = dialogCounter + 1
            Else
            End If
        Else
            mainGameDialog.SetDialogBoxText ("")
            dialogCounter = 0
        End If

        'Only update entities if they are in the collection
        'If they are beaten, they are deleted from the game to save memory
        Dim clsEntity As Entity

        For Each clsEntity In entityCollection
            If clsEntity.getBeaten = True Then

            Else
                clsEntity.Update
            End If
        Next

        'Constantly check if user has won the game yet
        checkGameWin

        Timeout (0.1)

    Loop

End Sub

'Emulates a wait function for a certain amount of milliseconds
Sub Timeout(duration_ms As Double)
'Declare a start time
    Dim Start_Time As Double

    'Make start time = current time
    Start_Time = Timer

    Do
        DoEvents
    Loop Until (Timer - Start_Time) >= duration_ms
End Sub

'http://codevba.com/help/collection.htm#.W__svOhKjZQ
'Show battle if the enemy exists and is not beaten yet
Public Sub CheckBattle(colliding As String)
    Dim entityID As String
    Dim clsEntity As Entity
    For Each clsEntity In entityCollection
        If clsEntity.getEntityUID = colliding Then
            If clsEntity.getBeaten = False Then
                Call InitBattleScene(hero, clsEntity)
                BattleUI.Show

                Exit Sub
            End If
        End If
    Next
End Sub

'When click interact button, check what the user is colliding with
'And show appropriate dialog/user form
Public Sub InteractButton()
    Call CheckBattle(hero.getCollidingWith)
    If hero.getCollidingWith = "shop" Then
        Call launchShop
    End If
End Sub

'Updates the collision tiles worksheets and feeds it to the user (hero.updateMap)
Public Sub setCollisions()
'3 Levels of Dungeons
    mainMapTiles = mainMapTilesWs.Range("Map").Value
    dungeonMapTiles1 = dungeonTilesWs1.Range("Map").Value
    dungeonMapTiles2 = dungeonTilesWs2.Range("Map").Value
    dungeonMapTiles3 = dungeonTilesWs3.Range("Map").Value

    gameMapHandler.getMap(1).setTiles (mainMapTiles)
    gameMapHandler.getMap(2).setTiles (dungeonMapTiles1)
    gameMapHandler.getMap(3).setTiles (dungeonMapTiles2)
    gameMapHandler.getMap(4).setTiles (dungeonMapTiles3)

    hero.updateMap
End Sub

'Setup all enemies in the three dungeons
Public Sub setupEnemies()

'Initialize arrays of enemies
    ReDim Preserve enemiesDungeon1(1)
    ReDim Preserve enemiesDungeon2(1)
    ReDim Preserve enemiesDungeon3(1)


    'Set up enemies in dungeon1
    Dim i As Integer
    Dim rand As Integer
    Dim randX As Integer
    Dim randY As Integer
    Dim ws As Worksheet

    'Set the worksheet to the dungeon map 1 tiles worksheet
    Set ws = Worksheets(gameMapHandler.getMap(2).getTilesWsName)

    'Fill in the first index
    Set enemiesDungeon1(1) = Factory.CreateEntity(GOBLIN_CHAR, 5, 5, 1, mainMapWs, _
                                                  gameMapHandler.getMap(2), 1, ENEMYHEALTH_LEVEL_1, BASEDMG_LEVEL_1, False, "Enemy", 64, 64, COINDROP_LEVEL_1)

    'Create enemies based on NUM_OF_ENEMIES constant declared at the top of the module
    For i = 2 To (2 + NUM_OF_ENEMIES - 1) Step 1

        'randomly posiition enemies
        randX = Application.WorksheetFunction.RandBetween(3, 17)
        randY = Application.WorksheetFunction.RandBetween(3, 8)

        'Avoid enemies overlapping on top of each other
        Do Until ws.Cells(randY + 1, randX) = "" And ws.Cells(randY + 1, randX + 1) = "" And ws.Cells(randY, randX) = "" And ws.Cells(randY, randX + 1) = ""
            randX = Application.WorksheetFunction.RandBetween(2, 10)
            randY = Application.WorksheetFunction.RandBetween(1, 5)
        Loop

        'Add the enemy to the array
        ReDim Preserve enemiesDungeon1(UBound(enemiesDungeon1) + 1)
        Set enemiesDungeon1(i) = Factory.CreateEntity(GOBLIN_CHAR, randX, randY, 1, mainMapWs, _
                                                      gameMapHandler.getMap(2), 1, ENEMYHEALTH_LEVEL_1, BASEDMG_LEVEL_1, False, "Enemy", 64, 64, COINDROP_LEVEL_1)
    Next

    'Dungeon 2 - same code logic as above
    Set ws = Worksheets(gameMapHandler.getMap(3).getTilesWsName)

    Set enemiesDungeon2(1) = Factory.CreateEntity(BAT_CHAR, 5, 5, 1, mainMapWs, _
                                                  gameMapHandler.getMap(3), 1, ENEMYHEALTH_LEVEL_2, BASEDMG_LEVEL_2, False, "Enemy", 32, 32, COINDROP_LEVEL_2)

    For i = 2 To (2 + NUM_OF_ENEMIES - 1) Step 1

        randX = Application.WorksheetFunction.RandBetween(3, 17)
        randY = Application.WorksheetFunction.RandBetween(3, 8)

        Do Until ws.Cells(randY + 1, randX) = "" And ws.Cells(randY + 1, randX + 1) = "" And ws.Cells(randY, randX) = "" And ws.Cells(randY, randX + 1) = ""
            randX = Application.WorksheetFunction.RandBetween(2, 10)
            randY = Application.WorksheetFunction.RandBetween(1, 5)
        Loop

        ReDim Preserve enemiesDungeon2(UBound(enemiesDungeon2) + 1)
        Set enemiesDungeon2(i) = Factory.CreateEntity(BAT_CHAR, randX, randY, 1, mainMapWs, _
                                                      gameMapHandler.getMap(3), 1, ENEMYHEALTH_LEVEL_2, BASEDMG_LEVEL_2, False, "Enemy", 32, 32, COINDROP_LEVEL_2)
    Next

    'Dungeon 3 - same code logic as above
    Set ws = Worksheets(gameMapHandler.getMap(4).getTilesWsName)

    Set enemiesDungeon3(1) = Factory.CreateEntity(SKELETON_CHAR, 5, 5, 1, mainMapWs, _
                                                  gameMapHandler.getMap(4), 1, ENEMYHEALTH_LEVEL_3, BASEDMG_LEVEL_3, False, "Enemy", 64, 64, COINDROP_LEVEL_3)


    For i = 2 To (2 + NUM_OF_ENEMIES - 1) Step 1

        randX = Application.WorksheetFunction.RandBetween(3, 17)
        randY = Application.WorksheetFunction.RandBetween(3, 8)

        Do Until ws.Cells(randY + 1, randX) = "" And ws.Cells(randY + 1, randX + 1) = "" And ws.Cells(randY, randX) = "" And ws.Cells(randY, randX + 1) = ""
            randX = Application.WorksheetFunction.RandBetween(2, 10)
            randY = Application.WorksheetFunction.RandBetween(1, 5)
        Loop

        ReDim Preserve enemiesDungeon3(UBound(enemiesDungeon3) + 1)
        Set enemiesDungeon3(i) = Factory.CreateEntity(SKELETON_CHAR, randX, randY, 1, mainMapWs, _
                                                      gameMapHandler.getMap(4), 1, ENEMYHEALTH_LEVEL_3, BASEDMG_LEVEL_3, False, "Enemy", 64, 64, COINDROP_LEVEL_3)
    Next

    'Add all enemies to the entityCollection collection object
    For i = LBound(enemiesDungeon1) To UBound(enemiesDungeon1) Step 1
        rand = WorksheetFunction.RandBetween(1, 4)
        entityCollection.Add enemiesDungeon1(i)
        enemiesDungeon1(i).setDirection = rand
    Next
    For i = LBound(enemiesDungeon2) To UBound(enemiesDungeon2) Step 1
        rand = WorksheetFunction.RandBetween(1, 4)
        entityCollection.Add enemiesDungeon2(i)
        enemiesDungeon2(i).setDirection = rand
    Next
    For i = LBound(enemiesDungeon3) To UBound(enemiesDungeon3) Step 1
        rand = WorksheetFunction.RandBetween(1, 4)
        entityCollection.Add enemiesDungeon3(i)
        enemiesDungeon3(i).setDirection = rand
    Next

End Sub

'Stop program
'Delete all objects
'Stop all sounds
'Return to main menu
Public Sub StopProgram()
    Dim ent As Variant
    sndPlaySound vbNullString, ByVal 0
    On Error Resume Next
    For Each ent In entityCollection
        ent.Remove
    Next
    Set entityCollection = Nothing
    gameMapHandler.removeAllMaps
    mainGameDialog.RemoveDialogBox

    Dim sObject As Shape
    For Each sObject In ActiveSheet.Shapes
        'Delete everything else except controls
        If sObject.name = "UIControlsBackPanel" Then
        Else
            sObject.Delete
        End If
    Next
    ThisWorkbook.Worksheets("Start").Activate
    End
End Sub

'Launch the shop
Public Sub launchShop()
    Call coinSound
    ShopUI.Show
End Sub

'Launch the inventory
Public Sub launchInventory()
    Call zipperSound
    Call InventoryUI.UpdateHealthBar
    InventoryUI.Show
End Sub

Public Sub launchHTP()
    HTPUI.Show
End Sub

'CHARACTER MOVEMENT BUTTONS
Public Sub LeftPressed()
    hero.setDirection = 1
    'set isMoving = True
    hero.MoveLeft
End Sub

Public Sub RightPressed()
    hero.setDirection = 2
    'set isMoving = True
    hero.MoveRight
End Sub

Public Sub UpPressed()
    hero.setDirection = 3
    'set isMoving = True
    hero.MoveUp
End Sub

Public Sub DownPressed()
    hero.setDirection = 4
    'set isMoving = True
    hero.MoveDown
End Sub

'Get a specific entity, given an entity UID from the collection
Public Function getEntity(UID As Integer) As Variant
    Dim clsEntity As Entity
    For Each clsEntity In entityCollection
        If clsEntity.getEntityUID = UID Then
            Set getEntity = clsEntity
            Exit Function
        End If
    Next
End Function

'https://stackoverflow.com/questions/14108948/excel-vba-check-if-entry-is-empty-or-not-space
'Check if string is empty
Public Function HasContent(txt As String) As Boolean
    HasContent = (Len(Trim(txt)) > 0)
End Function


'Check each dungeon and each entity in that dungeon
'If all enemies are defeated, then end the game and the user will win
Public Function checkGameWin()
    Dim clsEntity As Variant
    Dim win As Boolean
    Dim counter As Integer
    For Each clsEntity In enemiesDungeon1
        If clsEntity.getEntityName = "Enemy" Then
            If clsEntity.getBeaten = True Then
                counter = counter + 1
            ElseIf clsEntity.getBeaten = False Then

            End If
        End If
    Next

    For Each clsEntity In enemiesDungeon2
        If clsEntity.getEntityName = "Enemy" Then
            If clsEntity.getBeaten = True Then
                counter = counter + 1
            ElseIf clsEntity.getBeaten = False Then

            End If
        End If
    Next

    For Each clsEntity In enemiesDungeon3
        If clsEntity.getEntityName = "Enemy" Then
            If clsEntity.getBeaten = True Then
                counter = counter + 1
            ElseIf clsEntity.getBeaten = False Then

            End If
        End If
    Next

    If counter = (UBound(enemiesDungeon1) + UBound(enemiesDungeon2) + UBound(enemiesDungeon3)) Then
        WinUI.Show
    End If
End Function

