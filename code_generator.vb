Sub main()

    Dim database_name As String
    database_name = "temp.db"
    
    Dim sheet_name As String
    sheet_name = "Player"
    
    Dim file_name As String
    file_name = "c:\Youtube\python\players\team.csv"
    
    Dim MaxNumberOfFieldsPerRow As Integer
    MaxNumberOfFieldsPerRow = 5
    
    
    Call LetsGo(database_name, sheet_name, file_name, MaxNumberOfFieldsPerRow)
    
    
End Sub


Sub LetsGo(database_name As String, sheet_name As String, file_name As String, MaxNumberOfFieldsPerRow As Integer)
    
    Debug.Print "import sqlite3"
    Debug.Print "import csv"
    Debug.Print
    
    'build the create table command
    Call create_add_table_to_database(sheet_name)
    
    'build the load CSV file into SQLite3 database/table we just created
    Call create_csv_load_table(sheet_name, MaxNumberOfFieldsPerRow)
    
    'build __main__
    Call build__main(database_name, sheet_name, file_name)

End Sub

Sub create_add_table_to_database(sheet_name As String)

    Dim ws As Worksheet
    Dim comma As String
    Dim hasPK As Boolean
    Dim pk As String
    Dim field() As String
    Dim numOfColumns As Integer
    Dim row As Integer
   
    Set ws = Worksheets(sheet_name)
    
    'get the number of fields in table
    numOfColumns = get_number_of_columns(sheet_name)
    
    'does this table have a Primary key
    hasPK = does_table_have_PK(sheet_name)
    
    'build the primary key string of fields, comma separate
    pk = build_primary_key(hasPK, ws)
    
    'build column_name datatype
    field = build_column_names(hasPK, numOfColumns, ws)
    
    'def add_player_to_database(conn, cur):
    Debug.Print "def add_" & sheet_name & "_to_database(conn, cur):"
    Debug.Print
    
    'cur.execute("""
    Debug.Print vbTab & "cur.execute(" & Chr(34) & Chr(34) & Chr(34)
    
    'create table if not exists player
    Debug.Print TabOver(2) & "CREATE TABLE if not exists " & sheet_name
    Debug.Print TabOver(2) & "("
    
    'loop over all the fields
    For i = 1 To numOfColumns
        Debug.Print TabOver(3) & field(i)
    Next i
    
    If (pk <> "") Then
        Debug.Print TabOver(3) & "PRIMARY KEY (" & pk & ")"
    End If
    
    Debug.Print TabOver(2) & ")"
    Debug.Print TabOver(2) & Chr(34) & Chr(34) & Chr(34) & ")"
    Debug.Print vbTab; "conn.commit()"
    Debug.Print
    Debug.Print
    
End Sub

Function build_primary_key(hasPK As Boolean, ws As Worksheet) As String

    Dim row As Integer
    Dim pk As String
    
    row = 1
    pk = ""
    
    If (hasPK = True) Then
        While ws.Cells(row, 1) <> ""
            
            If (hasPK = True) Then
                If (LCase$(ws.Cells(row, 3)) = "pk") Then
                    If (pk = "") Then
                        pk = ws.Cells(row, 1)
                    Else
                        pk = pk & "," & ws.Cells(row, 1)
                    End If
                End If
            End If
            row = row + 1
        Wend
    End If
    
    If (hasPK = True) Then
        build_primary_key = pk
    Else
        build_primary_key = ""
    End If
    
End Function


Function build_column_names(hasPK As Boolean, numOfColumns As Integer, ws As Worksheet) As String()

    Dim row As Integer
    Dim pk As String
    Dim comma As String
    Dim field(1 To 250) As String
    
    row = 1
    comma = ","
    While ws.Cells(row, 1) <> ""
    
        'if this table DOES NOT HAVE a primary key
        If (hasPK = False) Then
            If (row = numOfColumns) Then
                comma = ""
            End If
        End If
        
        field(row) = ws.Cells(row, 1) & vbTab & vbTab & ws.Cells(row, 2) & comma
        row = row + 1
    Wend
    
    build_column_names = field
End Function

Sub create_csv_load_table(sheet_name As String, max_num_field_per_row As Integer)
    
    
    Dim ws As Worksheet
    Set ws = Worksheets(sheet_name)
    
    Dim numOfColumns As Integer
    numOfColumns = get_number_of_columns(sheet_name)
    
    Dim field(1 To 250) As String
    Dim fieldnames(1 To 250) As String
    Dim question(1 To 250) As String
    
    Dim index As Integer
    Dim Key As Integer
    Dim cnt As Integer
    Dim comma As String
    Dim i As Integer
    
    
    index = 1
    Key = 0
    cnt = 0
    comma = ","
    
    For i = 1 To numOfColumns
        
        If (i < numOfColumns) Then
            field(index) = field(index) & "entry[" & Key & "]" & comma
            fieldnames(index) = fieldnames(index) + ws.Cells(i, 1) & comma
            question(index) = question(index) + "?" & comma
            cnt = cnt + 1
        Else
             field(index) = field(index) & "entry[" & Key & "]" & ")"
             fieldnames(index) = fieldnames(index) + ws.Cells(i, 1) & ")"
             question(index) = question(index) + "?" & ")"
             cnt = cnt + 1
         End If
        
        Key = Key + 1
        If (cnt = max_num_field_per_row) Then
            If (i < numOfColumns) Then
                index = index + 1
                cnt = 0
            End If
        End If
    Next i
    
    
    
    Debug.Print "def load_" & sheet_name & "(conn, cur, filename):"
    Debug.Print
    Debug.Print vbTab & "with open(filename,""r"") as fh:"
    Debug.Print
    Debug.Print vbTab & vbTab & "reader = csv.reader(fh, delimiter=',')"
    Debug.Print
    Debug.Print vbTab & vbTab & "# if the file DOES NOT have a header row use this line"
    Debug.Print vbTab & vbTab & "next(csv.reader(fh), None)  # skip first row"
    Debug.Print
    
    Debug.Print
    Debug.Print vbTab & vbTab & "stmt = " & Chr(34) & "insert into " & sheet_name & "(" & Chr(34)
    For i = 1 To index
            Debug.Print vbTab & vbTab & "stmt += " & Chr(34) & fieldnames(i) & Chr(34)
    Next i
    
    
    For i = 1 To index
        If (i = 1) Then
            Debug.Print vbTab & vbTab & "stmt += " & Chr(34) & " values (";
            Debug.Print question(i) & Chr(34)
        Else
            Debug.Print vbTab & vbTab & "stmt += " & Chr(34) & question(i) & Chr(34)
        End If
    Next i
    
    
    Debug.Print
    Debug.Print vbTab & vbTab & "for entry in reader:"
    Debug.Print
    Debug.Print vbTab & vbTab & vbTab & "try:"
    Debug.Print
    
    
    For i = 1 To index
        If (i = 1) Then
            Debug.Print TabOver(4) & "record = (";
            Debug.Print field(i)
        Else
            Debug.Print TabOver(6) & field(i)
        End If
    Next i
        
    
    
    Debug.Print
    Debug.Print TabOver(4) & "cur.execute(stmt,record)"
    Debug.Print
    Debug.Print TabOver(4) & "conn.commit()"
    Debug.Print
    
    Debug.Print TabOver(3) & "except Exception as err:"
    Debug.Print
    Debug.Print TabOver(4) & "print(f'Line:{reader.line_num}, Record: {record}')"
    Debug.Print TabOver(4) & "print(f'Exception: {err}')"
    Debug.Print

End Sub

Sub build__main(database_name As String, sheet_name As String, filename As String)

    Debug.Print
    Debug.Print
    Debug.Print
    Debug.Print "if __name__ == " & Chr(34) & "__main__" & Chr(34) & ":"
    Debug.Print
    Debug.Print vbTab & "conn     = sqlite3.connect(" & Chr(34) & database_name & Chr(34) & ")"
    Debug.Print vbTab & "cur      = conn.cursor()"
    Debug.Print
    Debug.Print vbTab & "add_" & sheet_name & "_to_database(conn, cur)"
    
    
    pos = InStr(filename, "\\")
    If (pos = 0) Then
        filename = Replace(filename, "\", "\\", 1)
    End If
    
    
    Debug.Print vbTab & "load_" & sheet_name & "(conn, cur, """ & filename & """)"
    Debug.Print vbTab & "conn.close()"
    Debug.Print
    
End Sub

Function TabOver(count As Integer) As String

    Dim s As String
    Dim i As Integer
    
    For i = 1 To count
        s = s & vbTab
    Next i
    
    TabOver = s
End Function

Function get_number_of_columns(sheetName As String) As Integer

    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)
    
    Dim row As Integer
    Dim count As Integer
    row = 1
    count = 0
    
    While ws.Cells(row, 1) <> ""
        count = count + 1
        row = row + 1
    Wend
    
    get_number_of_columns = count
    
End Function


Function does_table_have_PK(sheetName As String) As Boolean

    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)
    
    
    Dim rows As Integer
    Dim count As Integer
    Dim i As Integer
    Dim has_pk As Boolean
    
    rows = get_number_of_columns(sheetName)
    count = 0
    
    has_pk = False
    
    
    For i = 1 To rows
        If (LCase$(ws.Cells(i, 3)) = "pk") Then
            has_pk = True
            Exit For
        End If
    Next i
    
    If has_pk = True Then
        does_table_have_PK = has_pk
    Else
        does_table_have_PK = False
    End If
    
End Function



