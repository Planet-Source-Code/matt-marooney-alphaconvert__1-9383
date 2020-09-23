<div align="center">

## AlphaConvert


</div>

### Description

Converts two character strings of A-Z or AA-ZZ into integer equivalents. Example: User inputs "A", returns 1, user inputs AE, returns 110 and so on.
 
### More Info
 
Two character string variable is passed to this function

Returns Integer equiv of the string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Marooney](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-marooney.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-marooney-alphaconvert__1-9383/archive/master.zip)





### Source Code

```
' User inputs a string of 2 characters, uppercase or lowercase.
'Function returns the combined integer value of the string (ex. A = 1, B=2...
'AA = 27, AB = 28...ect.)
Function GetNumber(UserInput As String) As Integer
Dim UpperCaseArray(1, 26) As String
Dim LowerCaseArray(1, 26) As String
Dim UpperCaseString As String
Dim LowerCaseString As String
Dim FirstNum As Integer
Dim SecondNum As Integer
UpperCaseString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
LowerCaseString = "abcdefghijklmnopqrstuvwxyz"
'Assign string characters to array cells
For x = 1 To Len(UpperCaseString)
    UpperCaseArray(1, x) = Mid(UpperCaseString, x, 1)
    LowerCaseArray(1, x) = Mid(LowerCaseString, x, 1)
Next x
If Len(UserInput) = 1 Then ' check for single character input
  For y = 1 To Len(UpperCaseString)
      'If the input from the user is A-Z or a-z the Function returns 1-26
      If Mid(UserInput, 1, 1) = UpperCaseArray(1, y) Then
        GetNumber = y
      End If
      If Mid(UserInput, 1, 1) = LowerCaseArray(1, y) Then
        GetNumber = y
      End If
  Next y
Else
  'If User Input has two characters...
  'Check first character...store numerical value in FirstNum
  For z = 1 To Len(UpperCaseString)
      If Mid(UserInput, 1, 1) = UpperCaseArray(1, z) Then
        FirstNum = z
      End If
      If Mid(UserInput, 1, 1) = LowerCaseArray(1, z) Then
        FirstNum = z
      End If
  Next z
  'Check second character
  'Store numerical value in SecondNum
  For w = 1 To Len(UpperCaseString)
      If Mid(UserInput, 2, 1) = UpperCaseArray(1, w) Then
        SecondNum = w
      End If
      If Mid(UserInput, 2, 1) = LowerCaseArray(1, w) Then
        SecondNum = w
      End If
  Next w
  'Algorithm for adding the values for the first character to that
  'of the second character to determine which set of 26 the user
  'selected.
  'i.e. if user enters "AA" then this loop determines that the first
  'character is equal to one. the loop returns 26 + 1, or 27. So, the
  'value of user input of "AA" is 27. And so on and so forth...
  'If the value entered is "BA", the algorithm returns 52 + 1, or 53
  'This loop will return the values for up to "IZ"
  'To extend to ZZ, merely change number of iterations in this loop to 26
  For V = 1 To 9
    If FirstNum = V Then
      GetNumber = ((26 * V) + SecondNum)
    End If
  Next V
End If
End Function
```

