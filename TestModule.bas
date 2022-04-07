Attribute VB_Name = "TestModule"
'@Folder("Lambda.Use")

'Developer: Md.Ismail Hosen
'Please contact for any project or VBA Automation.
'Email : 1997ismail.hosen@gmail.com
'Whatsapp: +8801515649307
'LinkedIn : https://www.linkedin.com/in/md-ismail-hosen-b77500135/
'Facebook : https://www.facebook.com/mdismail.hosen.7
'Youtube : https://www.youtube.com/channel/UCL-q7_WvISkw0Ox9FRBBzmw


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


Public Sub ArrayExample()
    
    Dim InputArray(1 To 2, 1 To 3) As Variant
    InputArray(1, 1) = "Developer: Md.Ismail Hosen"
    InputArray(1, 2) = "Email : 1997ismail.hosen@gmail.com"
    InputArray(1, 3) = "LinkedIn : https://www.linkedin.com/in/md-ismail-hosen-b77500135/"
    InputArray(2, 1) = "Facebook : https://www.facebook.com/mdismail.hosen.7"
    InputArray(2, 2) = "Youtube : https://www.youtube.com/channel/UCL-q7_WvISkw0Ox9FRBBzmw"
    InputArray(2, 3) = "Please contact for any project or VBA Automation."
    
    Dim Result As Variant
    Result = ApplyTransformationToAll(InputArray, New ProperCaseTransformer)
    
End Sub

Public Function ApplyTransformationToAll(GivenArray As Variant, TransformerFunction As ITransformer)
    
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(GivenArray, 2)
    Dim FirstRowIndex As Long
    FirstRowIndex = LBound(GivenArray, 1)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(GivenArray, 1) To UBound(GivenArray, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(GivenArray, 2) To UBound(GivenArray, 2)
            Dim ApplyOnText As String
            ApplyOnText = CStr(GivenArray(CurrentRowIndex, CurrentColumnIndex))
            GivenArray(CurrentRowIndex, CurrentColumnIndex) = TransformerFunction.Apply(ApplyOnText)
        Next CurrentColumnIndex
    Next CurrentRowIndex
    
    ApplyTransformationToAll = GivenArray
    
End Function

