Attribute VB_Name = "modAllChars"
Option Explicit

' All this does is return a list of "printable" ASCII characters
'
' This definitely is NOT rocket science!
'
' I threw it together because I was working on something unrelated
' and had to cobble this together as a quick utility function
'
' I have no idea what you might want to use this for,
' but it might be a handy exercise for beginning VB programmers
'
' Brian Battles, WS1O   brianb@cmtelephone.com

' constant that contains all the "printable" ASCII characters
Private Const AllVisibleChars As String = "01234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!#$%&()*+,-./:;<=>?@[\]^`{|}~€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ" & "'" & """" & "_"
Public Function ShowAllChars() As String
   
    ' Purpose   : shows all printable ASCII characters
    ' created   : 2/11/2002 By Brian Battles WS1O  brianb@cmtelephone.com
    
    Dim strTemp As String
    Dim I       As Integer
    
    For I = 1 To 255
        If InStr(AllVisibleChars, Chr(I)) Then
            If Format$(strTemp) = "" Then
                strTemp = Chr$(I)
            Else
                strTemp = strTemp & Chr$(I)
            End If
        End If
    Next
    ShowAllChars = strTemp
End Function
Sub Main()
    Debug.Print ShowAllChars
    MsgBox "                      Done..." & vbCrLf & vbCrLf & "         All printable ASCII characters" & vbCrLf & vbCrLf & "shown in your VB IDE's Immediate window", vbInformation, "           " & App.Title & " Finished, Big Deal!"
End Sub
