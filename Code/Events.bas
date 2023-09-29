Attribute VB_Name = "Events"
' Initialize stats
Public used, Start As Boolean

'Wait function - stops program for a set amount of seconds
Public Sub WaitASec(Sec As Long)
    Dim NowTick As Date
    Dim EndTick As Date
    EndTick = Now + TimeSerial(0, 0, Sec)
    Do
        NowTick = Now
        DoEvents    ' do nothing
    Loop Until NowTick >= EndTick
End Sub

Public Sub Wait()
    Do While continue = False
        DoEvents
    Loop
    continue = False
End Sub









