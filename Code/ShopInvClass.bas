Attribute VB_Name = "ShopInvClass"
'Module: ShopInvClass
'This module handles purchasing items from the shop and using items from the inventory

' Variables to run shop & inventory
Public isBought(1 To 8) As Integer

' Shop price constants
Public Const IRON_PRICE = 500
Public Const FIRE_PRICE = 1000
Public Const BATTLE_PRICE = 1500
Public Const MASTER_PRICE = 2000
Public Const SMALL_PRICE = 50
Public Const BIG_PRICE = 150
Public Const WARMOUR_PRICE = 2000
Public Const MARMOUR_PRICE = 3000

' Inventory damage constants
Public Const IRON_DMG = 50
Public Const FIRE_DMG = 80
Public Const BATTLE_DMG = 120
Public Const MASTER_DMG = 150
Public Const WARRIOR_HEALTH = 200
Public Const MASTER_HEALTH = 300
Public Const SMALL_HEALTH = 25
Public Const BIG_HEALTH = 50

'Sub that manages buying items from the shop
Public Sub buyItem(item As String, price As Long)

'If not enough coins, exit the sub (don't continue)
    If hero.getCoins < price Then
        MsgBox "You do not have enough coins to purchase this item."
        Exit Sub
    End If

    'Reduce coins that user has
    hero.setCoins (hero.getCoins - price)
    InventoryUI.ShopCoinsLabel.Caption = hero.getCoins
    ShopUI.ShopCoinsLabel.Caption = hero.getCoins
    Call buySound

    'Adjust display of user form based on which items purchased
    Select Case item
    Case Is = "IronSword"
        isBought(1) = 1
        ShopUI.IronSwordSold.Visible = True
        InventoryUI.QtyIronSword.Caption = 1
        InventoryUI.IronSwordInactive.Visible = False

    Case Is = "FireSword"
        isBought(2) = 1
        ShopUI.FireSwordSold.Visible = True
        InventoryUI.QtyFireSword.Caption = 1
        InventoryUI.FireSwordInactive.Visible = False


    Case Is = "Battleaxe"
        isBought(3) = 1
        ShopUI.BattleaxeSold.Visible = True
        InventoryUI.QtyBattleaxe.Caption = 1
        InventoryUI.BattleaxeInactive.Visible = False


    Case Is = "MasterSword"
        isBought(4) = 1
        ShopUI.MasterSwordSold.Visible = True
        InventoryUI.QtyMasterSword.Caption = 1
        InventoryUI.MasterSwordInactive.Visible = False


    Case Is = "WarriorArmour"
        isBought(7) = 1
        ShopUI.WarriorArmourSold.Visible = True
        InventoryUI.QtyWarriorArmour.Caption = 1
        InventoryUI.WarriorArmourInactive.Visible = False

    Case Is = "MasterArmour"
        isBought(8) = 1
        ShopUI.MasterArmourSold.Visible = True
        InventoryUI.QtyMasterArmour.Caption = 1
        InventoryUI.MasterArmourInactive.Visible = False

    Case Is = "SmallPotion"
        isBought(5) = isBought(5) + 1
        InventoryUI.QtySmallPotion.Caption = isBought(5)
        InventoryUI.InactiveSmallPotion.Visible = False

    Case Is = "BigPotion"
        isBought(6) = isBought(6) + 1
        InventoryUI.QtyBigPotion.Caption = isBought(6)
        InventoryUI.InactiveBigPotion.Visible = False
    End Select
End Sub

'When user uses a speciifc item from the inventory
'Increase base damage of user if buy weapon
'Increase health of user if use potion
'Increase base health of user if equip armour
Public Sub useItem(myItem As String, damage As Integer)

    Select Case myItem
    Case Is = "IronSword"
        purchasedWeapon = True
        InventoryUI.WeaponLabel.Caption = "Iron Sword"
        Call equipSound

        InventoryUI.WeaponDamageLabel.Caption = "+" & damage & " Damage"
        InventoryUI.WeaponImage.Picture = LoadPicture(wbPath & "\Shop-Items\" & myItem & ".jpg")
        hero.setBaseDmg = INITIAL_BASE_DMG_USER + IRON_DMG

    Case Is = "FireSword"
        purchasedWeapon = True
        InventoryUI.WeaponLabel.Caption = "Fire Sword"
        Call equipSound

        InventoryUI.WeaponDamageLabel.Caption = "+" & damage & " Damage"
        InventoryUI.WeaponImage.Picture = LoadPicture(wbPath & "\Shop-Items\" & myItem & ".jpg")
        hero.setBaseDmg = INITIAL_BASE_DMG_USER + FIRE_DMG

    Case Is = "BattleAxe"
        purchasedWeapon = True
        InventoryUI.WeaponLabel.Caption = "Battleaxe"
        Call equipSound

        InventoryUI.WeaponDamageLabel.Caption = "+" & damage & " Damage"
        InventoryUI.WeaponImage.Picture = LoadPicture(wbPath & "\Shop-Items\" & myItem & ".jpg")
        hero.setBaseDmg = INITIAL_BASE_DMG_USER + BATTLE_DMG

    Case Is = "MasterSword"
        purchasedWeapon = True
        InventoryUI.WeaponLabel.Caption = "Master Sword"
        Call equipSound

        InventoryUI.WeaponDamageLabel.Caption = "+" & damage & " Damage"
        InventoryUI.WeaponImage.Picture = LoadPicture(wbPath & "\Shop-Items\" & myItem & ".jpg")
        hero.setBaseDmg = INITIAL_BASE_DMG_USER + MASTER_DMG

    Case Is = "WarriorArmour"
        InventoryUI.ArmourLabel.Caption = "Warrior Armour"
        Call equipSound
        InventoryUI.ArmourHealthLabel.Caption = "+" & damage & " Health"
        InventoryUI.ArmourImage.Picture = LoadPicture(wbPath & "\Shop-Items\" & myItem & ".jpg")

        hero.setBaseHealth (WARRIOR_HEALTH)
        If hero.getHealth > WARRIOR_HEALTH Then
            hero.setHealth (WARRIOR_HEALTH)
        End If

    Case Is = "MasterArmour"
        InventoryUI.ArmourLabel.Caption = "Master Armour"
        Call equipSound
        InventoryUI.ArmourHealthLabel.Caption = "+" & damage & " Health"
        InventoryUI.ArmourImage.Picture = LoadPicture(wbPath & "\Shop-Items\" & myItem & ".jpg")

        hero.setBaseHealth (MASTER_HEALTH)

    Case Is = "SmallPotion"
        If isBought(5) > 0 Then
            Call healSound
            isBought(5) = isBought(5) - 1
            InventoryUI.QtySmallPotion.Caption = isBought(5)
            hero.setHealth WorksheetFunction.Min(hero.getHealth + SMALL_HEALTH, hero.getBaseHealth)
            If isBought(5) = 0 Then
                InventoryUI.InactiveSmallPotion.Visible = True
            End If
        ElseIf isBought(5) = 0 Then
            InventoryUI.InactiveSmallPotion.Visible = True
        End If

    Case Is = "BigPotion"
        If isBought(6) > 0 Then
            Call healSound
            isBought(6) = isBought(6) - 1
            Debug.Print (isBought(6))
            InventoryUI.QtyBigPotion.Caption = isBought(6)
            hero.setHealth WorksheetFunction.Min(hero.getHealth + BIG_HEALTH, hero.getBaseHealth)
            If isBought(6) = 0 Then
                InventoryUI.InactiveBigPotion.Visible = True
            End If
        ElseIf isBought(6) = 0 Then
            InventoryUI.InactiveBigPotion.Visible = True
        End If
    End Select

    'Update health bar to show changes in base health and health
    InventoryUI.UpdateHealthBar

End Sub

