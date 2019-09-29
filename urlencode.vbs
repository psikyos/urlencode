'https://www.cnblogs.com/gudai/articles/125044.html
'From the HTML standard: 
'http://www.w3.org/TR/html4/interact/forms.html#h-17.13.4.1
'application/x-www-form-urlencoded  
Public Function wj_URLEncode(strURL) 
Dim I 
Dim tempStr 
For I = 1 To Len(strURL) 
    If Asc(Mid(strURL, I, 1)) < 0 Then 
       tempStr = "%" & Right(CStr(Hex(Asc(Mid(strURL, I, 1)))), 2) 
       tempStr = "%" & Left(CStr(Hex(Asc(Mid(strURL, I, 1)))), Len(CStr(Hex(Asc(Mid(strURL, I, 1))))) - 2) & tempStr 
       wj_URLEncode = wj_URLEncode & tempStr 
    ElseIf (Asc(Mid(strURL, I, 1)) >= 65 And Asc(Mid(strURL, I, 1)) <= 90) Or (Asc(Mid(strURL, I, 1)) >= 97 And Asc(Mid(strURL, I, 1)) <= 122) Then 
       wj_URLEncode = wj_URLEncode & Mid(strURL, I, 1) 
    Else 
       wj_URLEncode = wj_URLEncode & "%" & Hex(Asc(Mid(strURL, I, 1))) 
    End If 
Next 
End Function