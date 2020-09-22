<div align="center">

## Function to return winsocks current connection state as a string


</div>

### Description

This code will return to you the state of the current winsock function in english rather than a number.. etc, "connected", "resolving host", etc.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BP](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bp.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bp-function-to-return-winsocks-current-connection-state-as-a-string__1-36737/archive/master.zip)





### Source Code

```
Public Function GetSockState(winsock1 as winsock) as String
  If winsock1.State = 0 Then
    GetSockState = "Connection Closed"
  ElseIf winsock1.State = 1 Then
    GetSockState = "Connection Open"
  ElseIf winsock1.State = 2 Then
    GetSockState = "Connection Listening"
  ElseIf winsock1.State = 3 Then
    GetSockState = "Connection Pending"
  ElseIf winsock1.State = 4 Then
    GetSockState = "Resolving Host"
  ElseIf winsock1.State = 5 Then
    GetSockState = "Host Resolved"
  ElseIf winsock1.State = 6 Then
    GetSockState = "Connecting"
  ElseIf winsock1.State = 7 Then
    GetSockState = "Connected"
  ElseIf winsock1.State = 8 Then
    GetSockState = "Peer is closing the connection"
  ElseIf winsock1.State = 9 Then
    GetSockState = "Error"
  End If
End Sub
```

