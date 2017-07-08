Option Compare Database
Option Explicit

Public dbName As String

Public db As DAO.Database
Public rs As DAO.Recordset


Sub InitVariables()
    dbName = CurrentDb.Name
    Set db = DAO.OpenDatabase(dbName, False, False)
    Set rs = db.OpenRecordset("Customers", dbOpenDynaset, dbSeeChanges)
End Sub

Sub LoopingRecordset()
    InitVariables
    
    Do Until rs.EOF
        Debug.Print "Name: " & rs![Last Name] & ", " & rs![First Name]
        rs.MoveNext
    Loop
End Sub

Sub FindingFirstInstance()
    InitVariables
    
    rs.FindFirst ("[Last Name] = 'Lee'")
    Debug.Print rs![First Name]
End Sub

Sub AddingRecordToRecordset()
    InitVariables
    
'    rs.AddNew
'    rs![Company] = "Company DD"
'    rs![Last Name] = "Jenkins"
'    rs![First Name] = "Leroy"
'    rs![E-mail Address] = "leroy.jenkins@wow.com"
'    rs![Job Title] = "Legend"
'    rs![Business Phone] = 1231234567
'    rs![Address] = "123 Main St."
'    rs![City] = "Sweet City"
'    rs![State/Province] = "FL"
'    rs![Zip/Postal Code] = 12345
'    rs![Country/Region] = "USA"
'    rs![Web Page] = "www.sweetsitebro.net"
'    rs![Notes] = "This is a note."
    
    Debug.Print rs.Fields.Count
End Sub

'want to tweak this a little
Sub LoopingFieldsInRecordset()
    InitVariables
    
    Dim i As Integer
    
    'rs.fields are zero based, count - 1 to avoid out of bounds error
    'had to do count - 2 because the last field has attachment
    Debug.Print "{"
    For i = 0 To rs.Fields.Count - 2
        If i < rs.Fields.Count - 2 Then
            Debug.Print vbTab & rs.Fields(i).Name & ": " & rs(i) & ","
        Else
            Debug.Print vbTab & rs.Fields(i).Name & ": " & rs(i)
        End If
    Next
    Debug.Print "}"

End Sub

'IsNull seems more accurate than IsEmpty for checking field values
Sub IsNullTest()
    InitVariables
    Debug.Print rs![Last Name]
    If IsNull(rs![First Name]) Then
        Debug.Print "is null"
    Else
        Debug.Print "is NOT null"
    End If
End Sub



