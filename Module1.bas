Option Explicit
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private pVtbl As LongPtr
Private func(0 To 2) As LongPtr
Private unk As IUnknown

Private Function Release(This As LongPtr) As Long

    MsgBox "呼ばれたお(・∀・)"

End Function

Sub Main()

    func(2) = VBA.Int(AddressOf Release)
    pVtbl = VarPtr(func(0))
    CopyMemory unk, VarPtr(pVtbl), 4

End Sub
