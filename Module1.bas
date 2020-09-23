Attribute VB_Name = "Module1"
' ******************************************************* '
' ********* CPULoad monitor for win95/98/ME/NT4 ********* '
' ** Coded by Simon Thwaites :: simon@lyricscircle.com ** '
' *** Feel free to use this code as you want thought  *** '
' ************ some credits would be nice :) ************ '
' ******************************************************* '
' ** Computer help? ************************************* '
' ** visit pchelp.bz for friendly support & advice  :) ** '
' ******************************************************* '

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Dim rkey As Long

Public Function CPUClose() As Long
RegCloseKey rkey
End Function

Public Function CPUStart() As Long
RegOpenKey &H80000006, "PerfStats\StatData", rkey
End Function

Public Function CPUUsage() As Long
Dim ret As Long
RegQueryValueEx rkey, "KERNEL\CPUUsage", 0&, REG_DWORD, ret, 4
CPUUsage = ret
End Function

