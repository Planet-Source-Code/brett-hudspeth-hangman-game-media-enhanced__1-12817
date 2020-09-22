Attribute VB_Name = "Utilities"
Option Explicit
Dim strWord As String
Dim intWrong As Integer

Public Function ComingIn(ByVal strW As String)
 strWord = strW
End Function

Public Function GoingOut() As String
 GoingOut = strWord
End Function

Public Function WrongOne(ByVal intI As Integer)
 intWrong = intI
End Function

Public Function Wrong() As Integer
 Wrong = intWrong
End Function
