VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SentenceCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Lambda.Implementation")
'This implementation folder has all the implementation of the Interface. This is the boilerplate code you have add
' to add a new transformer.
Option Explicit

Implements ITransformer

Private Function ITransformer_Apply(GivenText As String) As String
    '@TODO: Do your implementation.
End Function
