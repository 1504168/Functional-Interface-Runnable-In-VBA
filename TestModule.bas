Attribute VB_Name = "TestModule"
'@Folder("Lambda.Use")
Option Explicit

'Here I am going to use those code. If you want to add new Tranformation then just implement ITransformer class .

Sub Test()
    LogMessage "ISMAIL HOSEN", New LCaseTransformer, "Lower"
    LogMessage "ismail hosen", New UCaseTransformer, "Upper"
    LogMessage "isMAil HoSen", New ReverseCaseTransformer, "Reverse"
    LogMessage "isMAil HoSen", New ProperCaseTransformer, "Proper"
End Sub

Private Sub LogMessage(GivenText As String, TransformerFunction As ITransformer, TransformerName As String)
    Debug.Print "Convert To " & TransformerName & "case. Input : " & GivenText & " >>"; TransformText(GivenText, TransformerFunction)
End Sub

Private Function TransformText(GivenText As String, TransformerFunction As ITransformer)
   TransformText = TransformerFunction.Apply(GivenText)
End Function
