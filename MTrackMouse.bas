Attribute VB_Name = "MTrackMouse"
Option Explicit

Public colTrackMouse As New Collection
Public Function procTrackMouse(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tmItem As New CTrackMouse
Set tmItem = colTrackMouse.Item("TM" & hWnd)
If Not (tmItem Is Nothing) Then procTrackMouse = tmItem.MessageReceived(wMsg, wParam, lParam)
End Function

