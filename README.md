# AutoNumber


```vba
Public Function AutoNumber(TableName As String, FieldName As String) As Long
     Dim rs As New ADODB.Recordset
     Dim cn As New ADODB.Connection
     Dim stSQL As String, x As Long
     stSQL = "SELECT " & FieldName & " FROM " & TableName & " ORDER BY " & FieldName
     If cn.State = adStateOpen Then cn.Close
    Set cn = CurrentProject.AccessConnection
    With rs
         .Open stSQL, cn, , , adCmdText
          If .EOF Then
               x = 1
          Else
               If rs(0) > 1 Then
                    x = 1
               Else
                    x = rs(0)
                    Do Until .EOF
                         .MoveNext
                         If .EOF Then
                              x = x + 1
                              Exit Do
                         ElseIf x + 1 <> rs(0) Then
                              x = x + 1
                              Exit Do
                         Else
                              x = rs(0)
                         End If
                    Loop
               End If
          End If
          .Close
     End With
     AutoNumberReturn = x
End Function
```

## Call
```vba
TextBox.Value = AutoNumber("TableName", "PrimaryKey")

'example
txtId.Value = AutoNumber("tbBook","BookId")
```
