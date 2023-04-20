Function RemoveNumber(Txt As String) As String 
With CreateObject("VBScript.RegExp") 
.Global = True .Pattern = 10-91" 
RemoveNumber = .Replace(Txt, "") 
End With 
If Right(RemoveNumber, 1) = "_" 
Then RemoveNumber = Left(RemoveNumber, Len(RemoveNumber) - 1) 
Else RemoveNumber = RemoveNumber 
End If 
End Function 
