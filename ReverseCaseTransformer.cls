VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReverseCaseTransformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Lambda.Implementation")

Option Explicit

Implements ITransformer

Private Function ITransformer_Apply(GivenText As String) As String
    
    Dim Result As String
    Dim Counter As Long
    For Counter = 1 To Len(GivenText)
        Dim CurrentChar As String
        CurrentChar = Mid(GivenText, Counter, 1)
        If UCase(CurrentChar) = CurrentChar Then
            CurrentChar = LCase(CurrentChar)
        Else
            CurrentChar = UCase(CurrentChar)
        End If
        Result = Result & CurrentChar
    Next Counter
    
    ITransformer_Apply = Result
    
End Function

