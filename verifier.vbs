Option Explicit
Dim val,number,digit

'Function to diplay the specific value of a numkey pad into the textfield. 
Function inputValue(number)
   document.getElementById("num-ok").disabled =false
   digit=Len(document.getElementById("input-fields").value)
   val=document.getElementById("input-fields").value  
   If digit<16 Then
     val=val & CStr(number)
   End If
   document.getElementById("input-fields").value=val
End Function

'Function to display account number inputted in a MessageBox
Function getMessage()
   document.getElementById("num-ok").disabled =true 
   val=document.getElementById("input-fields").value
   digit=Len(document.getElementById("input-fields").value)
   if digit>=12 AND digit<=16 Then
     MsgBox "Your Account Number is :"&val
	 ClearFields()
   Else
     MsgBox "Minimum account number is 12"
   End If
End Function

'Function to clear all inputted number/digit
Sub ClearFields()
   document.getElementById("num-ok").disabled =true
   document.getElementById("input-fields").value=""  
End Sub

'Function to delete the last inputted number/digit
Function DeleteLastDigit()
   digit=Len(document.getElementById("input-fields").value)
   number=document.getElementById("input-fields").value  
   Do 
     val=Left(number,Len(number)-1)
     digit=digit-1
   Loop While digit>0
   document.getElementById("input-fields").value = val
End Function

  
