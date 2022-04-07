# Functional Interface (Runnable) In VBA
This Repository show you how to implement Runnable in VBA. In other programming language(Java,Javascript,Even in Power Query) you can pass code
as argument and that code will be used. So Let's say we have a very generic function which do simply like Text Transformation. So you want to take
the text and a transformer function which you will use on the text and return that. Now you don't want to hardcode the process about the 
transformation. You want to give flexibility to the client to pass whatever transformation mechanism they want to use and you want to use that.
So if they just simply want to convert it into upper case then they can do that. If they want to do lower case or proper case then they can do that.
Your code just depends on the transformer function which can be anything. Or Maybe user wants to do reverse case(Make lower case to upper and upper
to lower case) or sentence case. So that's kind of flexibility I am talking about. Our function is very generic function which doesn't depends on
solid logic but depends on abstraction. Here come's the Functional Interface(https://www.geeksforgeeks.org/functional-interfaces-java/). You can
write code which you can pass and use that.

For simple text it will not be so much useful. As VBA doesn't support Lambda (I am not talking about lambda function from worksheet) that's why 
this is a good way to have the same kind of behaviour as Lambda would give you. But you have to add some boilerplate code for this.

# Use 
Let's say you have a list of text(Array) and you want to apply the same transformation for each item
then you can just pass the array and this Transformation functional interface instance(Concrete Implementation) and then it will apply for each element.

# Example Code
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
