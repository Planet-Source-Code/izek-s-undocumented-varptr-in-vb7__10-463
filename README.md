<div align="center">

## undocumented varptr in vb7


</div>

### Description

VB6 has 3 undocumented functions. VarPtr, StrPtr, and ObjPtr.StrPtr returns object of a string, VarPtr returns address of any other variable and ObjPtr returns address of an object. In VB.Net these 3 functions are abosolete but their functionality is still avaliable.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Izek\_S](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/izek-s.md)
**Level**          |Intermediate
**User Rating**    |5.0 (45 globes from 9 users)
**Compatibility**  |VB\.NET
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__10-33.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/izek-s-undocumented-varptr-in-vb7__10-463/archive/master.zip)





### Source Code

<br><br><br>
Public Function VarPtr(ByVal o As Object) As Integer<br><br>
Dim GC As System.Runtime.InteropServices.GCHandle = System.Runtime.InteropServices.GCHandle.Alloc(o, System.Runtime.InteropServices.GCHandleType.Pinned)<br><br>
Dim ret As Integer = GC.AddrOfPinnedObject.ToInt32<br><br>
GC.Free()<br><br>
Return ret<br><br>
End Function

