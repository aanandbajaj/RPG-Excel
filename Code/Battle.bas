Attribute VB_Name = "Battle"
'Module: Battle
'Manages entire battle scene
'setCharacterImages: sets both entity images based on who the user is battling (bat, goblin, etc.)
'InitBattleScene: sets all initial variables
'UsePotion: allows user to heal up while battling
'HeroMove: manages damage to enemy based on which move the user picks
'CheckUserWin: Check after every move whether user won, as that will end the battle
'EnemyAttack: manages enemy damage to user

Public user As Entity
Public opponent As Entity
Public continue As Boolean
Public specialAtkUsed As Boolean
Public purchasedWeapon As Boolean

'Declaring constant variables that represent lower and upper limits of battle move damages
Public Const BASIC_WEAPON_DMG_LOWER = 50
Public Const BASIC_WEAPON_DMG_UPPER = 70
Public Const SPECIAL_WEAPON_DMG_LOWER = 100
Public Const SPECIAL_WEAPON_DMG_UPPER = 200

' Sub to set the images of the combatants in a battle
Public Sub setCharacterImages()
    BattleUI.UserImage.Picture = LoadPicture(wbPath & "\Images\Character Sprites\" & user.getEntityID & "\BattleSideFace.jpg")
    BattleUI.OpponentImage.Picture = LoadPicture(wbPath & "\Images\Character Sprites\" & opponent.getEntityID & "\BattleSideFace.jpg")
End Sub

' Sub to initialize battle scene
' Allows for generalization of battle system (can pass through any two entities)
Public Sub InitBattleScene(ByRef userEntity As Variant, ByRef opponentEntity As Variant)

    continue = True

    Set user = userEntity
    Set opponent = opponentEntity

    Call setCharacterImages

    BattleUI.UserNameLabel.Caption = user.getEntityName
    BattleUI.OpponentNameLabel.Caption = opponent.getEntityID
    
    BattleUI.UserHP_Label = "HP " & user.getHealth & "/" & user.getBaseHealth
    BattleUI.OpponentHP_Label = "HP " & opponent.getHealth & "/" & opponent.getBaseHealth
    Debug.Print (opponent.getHealth)
    Debug.Print (opponent.getBaseHealth)

    mainGameDialog.SetDialogBoxText ("")
    dialogCounter = 0

    BattleUI.DialogueBox.Caption = "Oh no! You have encountered the " & opponent.getEntityID & "." & _
                                   vbCrLf & "What will you choose to do?"

End Sub
' Sub to handle the hero using one of the two items
Public Sub usePotion(size As String)

    Select Case size
    Case Is = "Small"
        health = SMALL_HEALTH
        Call healSound
        ' De-increment qty of small potions
        isBought(5) = isBought(5) - 1
        InventoryUI.QtySmallPotion.Caption = isBought(5)
        BattleUI.QtySmallPotion.Caption = isBought(5)
        ' Make small potions un-useable if qty is zero
        If isBought(5) = 0 Then
            BattleUI.InactiveSmallPotion.Visible = True
            InventoryUI.InactiveSmallPotion.Visible = True
        End If
    Case Is = "Big"
        health = BIG_HEALTH
        Call healSound
        ' De-increment qty of big potions
        isBought(6) = isBought(6) - 1
        InventoryUI.QtyBigPotion.Caption = isBought(6)
        BattleUI.QtyBigPotion.Caption = isBought(6)
        ' Make big potions un-useable if qty is zero
        If isBought(6) = 0 Then
            BattleUI.InactiveBigPotion.Visible = True
            InventoryUI.InactiveBigPotion.Visible = True
        End If
    End Select

    Call BattleUI.GoBackButton_Click

    ' Calculate change in health and apply to entity object and battle scene labels
    user.setHealth (WorksheetFunction.Min(user.getHealth + health, user.getBaseHealth))
    Call setUserHPUI

    BattleUI.DialogueBox.Caption = _
    "Game: " & user.getEntityName & " used " & size & " Potion" _
                                   & user.getEntityName & " gained " & health & "HP!"

End Sub

Public Sub heroMove(move As String)

    BattleUI.Punch.Visible = False
    BattleUI.Kick.Visible = False
    BattleUI.WeaponMove.Visible = False
    BattleUI.SpecialMove.Visible = False
    Dim dmg As Integer

    ' Each take is user.getBaseDmg dmg + special effect
    Select Case move
    Case Is = "Punch"
        Call punchKickSound
        dmg = Round((35 - 30) * Rnd() + 30)
        If Rnd() > 0.8 Then
            dmg = dmg + 20
        End If
        BattleUI.DialogueBox.Caption = "Game: " & user.getEntityName & " performs: Heavy Punch! The attack does " & dmg & " damage."

    Case Is = "Kick"
        Call punchKickSound
        dmg = Round(Rnd() * (35 - 20) + 20)
        ' Chance at extra damage at diminishing probability
        If (Rnd() > 0.3) Then
            dmg = dmg * 2
        End If
        If (Rnd() > 0.6) Then
            dmg = dmg * 3
        End If
        If (Rnd() > 0.9) Then
            dmg = dmg * 4
        End If
        BattleUI.DialogueBox.Caption = "Game: " & user.getEntityName & " performs: Kicks of Fury! The attack does " & dmg & " damage."

        'If user chooses to do basic weapon attack
    Case Is = "Weapon"
        Call swordSound
        dmg = Round(Rnd() * (user.getBaseDmg + BASIC_WEAPON_DMG_UPPER - BASIC_WEAPON_DMG_LOWER) + BASIC_WEAPON_DMG_LOWER)
        BattleUI.DialogueBox.Caption = "Game: " & user.getEntityName & " performs: " & InventoryUI.WeaponLabel.Caption & " Slash! The attack does " & dmg & " damage."

        'If user chooses to do special weapon attack (only allowed once per game)
    Case Is = "Special"
        Call swordSound
        specialAtkUsed = True
        dmg = Round(Rnd() * (user.getBaseDmg + SPECIAL_WEAPON_DMG_UPPER - SPECIAL_WEAPON_DMG_LOWER) + SPECIAL_WEAPON_DMG_LOWER)
        BattleUI.DialogueBox.Caption = "Game: " & user.getEntityName & " performs: " & InventoryUI.WeaponLabel.Caption & " Special Attack! The attack does " & dmg & " damage."
    End Select

    'setting opponent HP bar based on user's damage to enemy
    opponent.setHealth (WorksheetFunction.Max((opponent.getHealth - dmg), 0))
    BattleUI.opponentHP_bar.width = ((opponent.getBaseHealth - opponent.getHealth) / opponent.getBaseHealth) * 130
    BattleUI.OpponentHP_Label = "HP " & opponent.getHealth & "/" & opponent.getBaseHealth

    Call Wait

    'Checks whether Enemy/User has won the game after every move
    If CheckUserWin = True Then
        Exit Sub
    End If
    If EnemyAttack = True Then
        Exit Sub
    End If

    'Continue battle sequence
    BattleUI.Layer1.Visible = False
    BattleUI.DialogueBox.Caption = "What will you choose next?"
    BattleUI.ContinueButton.Visible = False

End Sub

Public Function CheckUserWin() As Boolean

' Check for win condition
    If opponent.getHealth = 0 Then
        CheckUserWin = True
        BattleUI.DialogueBox.Caption = "Game: You have defeated " & opponent.getEntityID & "!"
        Call Wait
        BattleUI.DialogueBox.Caption = opponent.getEntityID & ": I let you beat me!"
        opponent.setBeaten (True)
        Call winSound
        Call Wait

        ' Calculate coins and assign to hero and fix labels
        Dim coins As Integer
        coins = Int(WorksheetFunction.RandBetween(opponent.getCoinsDrop, opponent.getCoinsDrop + 50))
        user.setCoins (user.getCoins + coins)
        coinText.TextFrame.Characters.text = user.getCoins

        'Delete opponent from map (clear the collision) & delete the opponent's shapes
        Call opponent.clearCollisionOnMap
        'Reset collisions
        Call setCollisions
        opponent.Remove

        ' Unload battle scene
        BattleUI.Hide
        Unload BattleUI

        'Gained x coins message in dialog box after user form is unloaded
        mainGameDialog.SetDialogBoxText ("")
        dialogCounter = 0
        mainGameDialog.SetDialogBoxText ("You gained " & coins & " coins!")
        Call coinSoundEffect

        Exit Function
    Else
        CheckUserWin = False
    End If
End Function

Public Function EnemyAttack() As Boolean

' Randomly select enemy attack
    atk = WorksheetFunction.RandBetween(1, 2)

    ' Calculate damage based on opponent's base damage
    Select Case atk
    Case Is = 1
        dmg = Round((opponent.getBaseDmg - 10) * Rnd() + 10)

    Case Is = 2
        dmg = Round((opponent.getBaseDmg - 20) * Rnd() + 20)

    End Select

    Call enemyGruntSound
    BattleUI.DialogueBox.Caption = "Game: " & opponent.getEntityID & " performs: an attack! " & "The attack does " & dmg & " damage."

    ' User HP calculation and set the HP bar
    user.setHealth (WorksheetFunction.Max((user.getHealth - dmg), 0))
    Call setUserHPUI

    Call Wait

    ' Check for enemy win condition
    'If enemy wins, then show a few messages and close the battle user form
    If user.getHealth = 0 Then
        EnemyAttack = True
        Call userLoseSound
        BattleUI.DialogueBox.Caption = "Game: You have been defeated by " & opponent.getEntityID & "!"
        Call Wait
        BattleUI.DialogueBox.Caption = opponent.getEntityID & ": I have won! Game over for you!"
        Call Wait
        Call StopProgram
        BattleUI.Hide
        Unload BattleUI
    Else
        EnemyAttack = False
    End If
End Function

'Set User's HP bar
Public Function setUserHPUI()
    BattleUI.userHP_bar.width = ((user.getBaseHealth - user.getHealth) / user.getBaseHealth) * 130
    BattleUI.userHP_bar.Left = 230 - ((user.getBaseHealth - user.getHealth) / user.getBaseHealth) * 130
    BattleUI.UserHP_Label = "HP " & user.getHealth & "/" & user.getBaseHealth
End Function

