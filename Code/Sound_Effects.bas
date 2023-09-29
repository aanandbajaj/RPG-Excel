Attribute VB_Name = "Sound_Effects"
'Module: Sound_Effects
'This module handles sounds used throughout the game

' Source:
' https://www.exceltip.com/general-topics-in-vba/playing-wav-files-using-vba-in-microsoft-excel.html

' Function using the winmm.dll library to handle .wav files
Public Declare PtrSafe Function sndPlaySound Lib "winmm.dll" _
       Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                              ByVal uFlags As Long) As Long
Private Declare PtrSafe Function PlaySound Lib "winmm.dll" _
        Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Boolean

Const SND_ASYNC = &H1
Const SND_NOSTOP = &H10

Sub PlayWavFile(WavFileName As String, Wait As Boolean)
    If dir(WavFileName) = "" Then Exit Sub    ' no file to play
    If Wait Then    ' play sound before running any more code
        sndPlaySound WavFileName, 0
        'When wait = false
    Else    ' play sound while code is running
        sndPlaySound WavFileName, 1
    End If
End Sub

' ################## Call subs to trigger a sound effect ########################
Public Sub coinSound()
    Dim sound As String
    sound = wbPath & "\Sound-Effects\enter-shop.wav"    ' https://www.soundjay.com/coin-sounds-1.html
    Call PlayWavFile(sound, False)
End Sub
Public Sub doorSound()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\exit-shop.wav"    ' https://www.freesoundeffects.com/free-sounds/doors-10030/
    Call PlayWavFile(sound, False)
End Sub
Public Sub buySound()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\buy-sound.wav"    ' http://soundbible.com/tags-cash-register.html
    Call PlayWavFile(sound, False)
End Sub
Public Sub zipperSound()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\bag-sound.wav"    'https://freesound.org/people/AntumDeluge/sounds/188044/
    Call PlayWavFile(sound, False)
End Sub
Public Sub coinSoundEffect()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\coin.wav"
    Call PlayWavFile(sound, False)
End Sub
Public Sub winSound()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\win-sound.wav"
    Call PlayWavFile(sound, False)
End Sub
Public Sub userLoseSound()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\gameover-sound.wav"
    Call PlayWavFile(sound, False)
End Sub
Public Sub enemyGruntSound()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\enemy-grunt.wav"
    Call PlayWavFile(sound, False)
End Sub
Public Sub swordSound()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\sword-sound.wav"
    Call PlayWavFile(sound, False)
End Sub
Public Sub punchKickSound()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\punchkick-sound.wav"
    Call PlayWavFile(sound, False)
End Sub

Public Sub equipSound()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\metal-equip.wav"
    Call PlayWavFile(sound, False)
End Sub

Public Sub healSound()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\heal-sound.wav"
    Call PlayWavFile(sound, False)
End Sub

Public Sub doorSoundEffect()
    Dim sound As String
    ' Convert to relative path
    sound = wbPath & "\Sound-Effects\door-sound.wav"
    Call PlayWavFile(sound, False)
End Sub


