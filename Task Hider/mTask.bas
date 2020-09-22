Attribute VB_Name = "mTask"
Option Explicit

Private Enum LVITEM_Mask
    LVIF_TEXT = &H1
    LVIF_IMAGE = &H2
    LVIF_PARAM = &H4
    LVIF_STATE = &H8
    LVIF_INDENT = &H10
    LVIF_NORECOMPUTE = &H800
End Enum
Private Enum LVITEM_States
    LVIS_FOCUSED = &H1
    LVIS_SELECTED = &H2
    LVIS_CUT = &H4
    LVIS_DROPHILITED = &H8
    LVIS_ACTIVATING = &H20
    LVIS_OVERLAYMASK = &HF00
    LVIS_STATEIMAGEMASK = &HF000
End Enum
Private Type LVITEM
    Mask As LVITEM_Mask
    iItem As Long
    iSubItem As Long
    State As LVITEM_States
    stateMask As LVITEM_States
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type


Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)
Private Const LVM_DELETEITEM As Long = (LVM_FIRST + 8)
Private Const LVM_GETITEMTEXTA As Long = (LVM_FIRST + 45)
Private Const LVM_SETITEMTEXTA As Long = (LVM_FIRST + 46)

Public Sub SetItemText(ByVal Handle As Long, ByVal pStr As String, ByVal Index As Long, Optional ByVal SubIndex As Long = 0)

    Dim hProcess As Long, SharedProcMem As Long, LVISize As Long
    Dim SharedProcMemString  As Long, strSize As Long
    Dim nCount As Long, LenWritten As Long, pId As Long
    Dim LVI As LVITEM
    Dim MemStorage() As Byte
    
    If IsWindowsNT Then
        LVISize = Len(LVI)
        Call GetWindowThreadProcessId(Handle, pId)
        'alloc some shared memory for our Struct
        SharedProcMem = GetMemSharedNT(pId, LVISize, hProcess)
        
        MemStorage = StrConv(pStr & vbNullChar, vbFromUnicode)
        strSize = UBound(MemStorage) + 1
        'alloc some shared memory for our string
        SharedProcMemString = GetMemSharedNT(pId, strSize, hProcess)
        'setup some info
        With LVI
            .iItem = Index
            .iSubItem = SubIndex
            .cchTextMax = strSize
            .pszText = SharedProcMemString 'store our string handle
        End With
        
        'write to memory
        WriteProcessMemory hProcess, ByVal SharedProcMemString, MemStorage(0), strSize, LenWritten
        WriteProcessMemory hProcess, ByVal SharedProcMem, LVI, LVISize, LenWritten
        
        'get the text
        Call SendMessage(Handle, LVM_SETITEMTEXTA, Index, ByVal SharedProcMem)
        
        'clean up
        FreeMemSharedNT hProcess, SharedProcMem, LVISize
        FreeMemSharedNT hProcess, SharedProcMemString, strSize
     End If
     
End Sub

Public Function GetItemText(ByVal Handle As Long, ByVal Index As Long, Optional ByVal SubIndex As Long = 0) As String

    Dim hProcess As Long, SharedProcMem As Long, LVISize As Long
    Dim SharedProcMemString  As Long, strSize As Long
    Dim nCount As Long, LenWritten As Long, pId As Long
    Dim LVI As LVITEM
    Dim MemStorage() As Byte
    
    If IsWindowsNT Then
        LVISize = Len(LVI)
        
        MemStorage = StrConv(String$(255, 0), vbFromUnicode)
        strSize = UBound(MemStorage) + 1
        
        Call GetWindowThreadProcessId(Handle, pId)
        SharedProcMem = GetMemSharedNT(pId, LVISize, hProcess)
        SharedProcMemString = GetMemSharedNT(pId, strSize, hProcess)

        With LVI
            .iItem = Index
            .iSubItem = SubIndex
            .cchTextMax = strSize
            .pszText = SharedProcMemString
        End With
        
        WriteProcessMemory hProcess, ByVal SharedProcMem, LVI, LVISize, LenWritten
        Call SendMessage(Handle, LVM_GETITEMTEXTA, Index, ByVal SharedProcMem)
        ReadProcessMemory hProcess, ByVal SharedProcMemString, MemStorage(0), strSize, LenWritten
        'clean up
        FreeMemSharedNT hProcess, SharedProcMem, LVISize
        FreeMemSharedNT hProcess, SharedProcMemString, strSize

    End If
    
    GetItemText = StrConv(MemStorage, vbUnicode)
    If InStr(1, GetItemText, vbNullChar) Then 'strip nulls
        GetItemText = Left$(GetItemText, InStr(1, GetItemText, vbNullChar) - 1)
    End If
    
End Function

Public Function GetItemCount(ByVal Handle As Long) As Long

    GetItemCount = SendMessage(Handle, LVM_GETITEMCOUNT, 0&, ByVal 0&)

End Function

Public Sub DeleteItem(ByVal Handle As Long, ByVal Index As Long)
    
    Call SendMessage(Handle, LVM_DELETEITEM, Index, ByVal 0&)

End Sub
