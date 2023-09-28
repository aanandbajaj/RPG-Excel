# RPG-Excel

## How to Play
Step 1: Open the .XLSM File </br>
Step 2: Click 'Start Game' </br>
Step 3: Enjoy

*Only works on Windows*

![Home](https://github.com/aanandbajaj/RPG-Excel/blob/master/ReadmeImages/Main_Screen.JPG)
![GUI](https://github.com/aanandbajaj/RPG-Excel/blob/master/ReadmeImages/GUI.JPG)


## Technical Details
This is an RPG Game made in Excel using VBA.

### What is an RPG Game?
Role-Playing Games (RPGs) are games in which players assume the roles of characters in a fictional setting.

* Typically includes a series of “quests”, or smaller tasks that must be completed in order to accomplish the overall goal of the game
* There is typically no time limit
* There is a set of rules and guidelines that must be followed in order to successfully complete the quests
* The game is not a continuous event, rather it plays out as the character progresses through the story

### How the Maps Were Made
Used an application called "Tiled" to build the image of the map (20 x 10)
![Main Map](https://github.com/aanandbajaj/RPG-Excel/blob/master/ReadmeImages/Main_Map_Scene_Door.png)

To implement collisions, I mapped the 20 x 10 grid system on to another Excel spreadsheet, where each cell was a floor, door, wall, etc. Putting the map image and the map tiles together allowed me to implement a complete generalized map system. 
![Main Map Tiles](https://github.com/aanandbajaj/RPG-Excel/blob/master/ReadmeImages/Main_Map_Tiles.JPG)

### Sprite System 
Since games are generally not made on Excel, I had to create a custom sprite system. The process was as followed:
* Load sprite character images at the beginning of the game
* Stack all the sprite images on top of each other 
* Emulate "animation" by rotating through the images (e.g. walk cycle) as the up, down, left, and right arrow keys are pressed

![Walk Cycle](https://github.com/aanandbajaj/RPG-Excel/blob/master/ReadmeImages/WalkCycle.JPG)

### Battle System 
* The battle system was created to emulate a traditional RPG game, so it is turn-based
* There are multiple types of enemies you can fight as the user, including goblins, bats, and skeletons, each with a different level of strength
* Below is an example of how the battle system would look like 

![Battle1](https://github.com/aanandbajaj/RPG-Excel/blob/master/ReadmeImages/Battle.JPG)

![Battle2](https://github.com/aanandbajaj/RPG-Excel/blob/master/ReadmeImages/Battle_Skeleton.JPG)

![Battle3](https://github.com/aanandbajaj/RPG-Excel/blob/master/ReadmeImages/Battle_Moves.JPG)


### The Inventory 
* As you defeat enemies, you get coins, which is an in-game currency
* You can use these coins to upgrade your armour, buy better weapons, and re-stock on shield potion

The Inventory 
![Inventory](https://github.com/aanandbajaj/RPG-Excel/blob/master/ReadmeImages/Inventory.JPG)

The Shop 
![Shop](https://github.com/aanandbajaj/RPG-Excel/blob/master/ReadmeImages/Shop.JPG)
