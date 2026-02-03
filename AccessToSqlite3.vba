Sub ExportSQLiteSQL(FilePath As String)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim rel As DAO.Relation
    Dim f As DAO.Field
    Dim sql As String
    Dim fso As Object
    Dim txtFile As Object
    Dim TableSQL As Collection
    Dim TableName As String
    Dim TableRelSQL As Collection
    
    Set db = CurrentDb
    Set TableSQL = New Collection
    Set TableRelSQL = New Collection
    
    ' ---- Parcours des tables ----
    For Each tdf In db.TableDefs
        ' Ignorer les tables système
        If Left(tdf.Name, 4) <> "MSys" Then
            sql = "CREATE TABLE " & tdf.Name & " (" & vbCrLf
            For Each fld In tdf.Fields
                sql = sql & "    " & fld.Name & " "
                Select Case fld.Type
                    Case dbText, dbMemo
                        sql = sql & "TEXT"
                    Case dbLong, dbInteger
                        If fld.Attributes And dbAutoIncrField Then
                            sql = sql & "INTEGER PRIMARY KEY AUTOINCREMENT"
                        Else
                            sql = sql & "INTEGER"
                        End If
                    Case dbDouble, dbCurrency
                        sql = sql & "REAL"
                    Case dbBoolean
                        sql = sql & "INTEGER"
                    Case dbDate
                        sql = sql & "TEXT"
                    Case Else
                        sql = sql & "TEXT"
                End Select
                
                ' Virgule sauf pour le dernier champ
                If fld.OrdinalPosition < tdf.Fields.Count - 1 Then sql = sql & ","
                sql = sql & vbCrLf
            Next fld
            sql = sql & ");" & vbCrLf & vbCrLf
            TableSQL.Add sql
        End If
    Next tdf
    
    ' ---- Parcours des relations pour générer FOREIGN KEY ----
    For Each rel In db.Relations
        ' Ignorer les relations système
        If Left(rel.Name, 4) <> "MSys" Then
            Dim fkFields As String, pkFields As String
            fkFields = ""
            pkFields = ""
            For Each f In rel.Fields
                fkFields = fkFields & f.ForeignName & ", "
                pkFields = pkFields & f.Name & ", "
            Next f
            fkFields = Left(fkFields, Len(fkFields) - 2)
            pkFields = Left(pkFields, Len(pkFields) - 2)
            
            Dim relSQL As String
            relSQL = "ALTER TABLE " & rel.ForeignTable & vbCrLf & _
                     "ADD FOREIGN KEY (" & fkFields & ") REFERENCES " & rel.Table & " (" & pkFields & ");" & vbCrLf & vbCrLf
            TableRelSQL.Add relSQL
        End If
    Next rel
    
    ' ---- Écriture dans le fichier ----
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(FilePath, True)
    
    ' Activer les clés étrangères pour SQLite
    txtFile.WriteLine "PRAGMA foreign_keys = ON;" & vbCrLf
    
    ' Écrire les tables
    For Each sql In TableSQL
        txtFile.WriteLine sql
    Next sql
    
    ' Écrire les relations
    For Each sql In TableRelSQL
        txtFile.WriteLine sql
    Next sql
    
    txtFile.Close
    
    MsgBox "Script SQLite exporté vers : " & FilePath
End Sub
