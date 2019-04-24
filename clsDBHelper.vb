Option Strict Off
Imports System.Data.SqlClient
Public Class clsDBHelper

    Public _SQLConn As SqlConnection

    Public DataTable As New DataTable

    Public Sub New(ByRef AccessConn As SqlConnection)

        _SQLConn = AccessConn

    End Sub

    Public Function ExecuteStoredProc(ByVal strStoredProc As String, ByVal _parameters() As SqlParameter) As Boolean


        Try

            Dim cm As SqlCommand = New SqlCommand(strStoredProc, _SQLConn)
            cm.CommandType = CommandType.StoredProcedure

            If (Not (_parameters) Is Nothing) Then
                For Each p As SqlParameter In _parameters
                    cm.Parameters.Add(p)

                Next
            End If

            cm.ExecuteNonQuery()


            Return True

        Catch ex As Exception

            Throw New ApplicationException(ex.Message)

            Return False

        End Try
    End Function


    Public Function TableExist(ByVal tblName As String) As Boolean

        '  Specify restriction to get table definition schema
        '  For reference on GetSchema see:
        '  http://msdn2.microsoft.com/en-us/library/ms254934(VS.80).aspx
        Dim restrictions As String() = New String(2) {}
        restrictions(2) = tblName
        Dim dbTbl As DataTable = _SQLConn.GetSchema("Tables", restrictions)

        If (dbTbl.Rows.Count = 0) Then
            ' Table does not exist
            Return False
        Else
            ' Table exists
            Return True
        End If

        dbTbl.Dispose()

    End Function


    ''' <summary>
    ''' Returns a complete dataset from ; delimited queries
    ''' </summary>
    ''' <param name="strQuery"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataSet(ByVal strQuery As String) As System.Data.DataSet

        'NOTE:  Since the JET engine does not support multiple queries, we need to create separate data adapters and add them to the dataset..

        'trim the last ; off if it exists...
        strQuery = strQuery.Trim()
        If (strQuery.Substring(strQuery.Length - 1, 1) = ";") Then strQuery = Left(strQuery, Len(strQuery) - 1)


        Dim strQueries() As String = Split(strQuery, ";")
        Dim da(strQueries.Length - 1) As SqlDataAdapter

        Dim ds As New System.Data.DataSet

        For x As Int32 = 0 To strQueries.Length - 1

            '2. Create the command object, passing in the SQL string    
            Dim cmd As New SqlCommand(strQueries(x), _SQLConn)

            '3. Create the DataAdapter
            da(x) = New SqlDataAdapter()
            da(x).SelectCommand = cmd

            '4. Populate the DataSet and close the connection

            da(x).Fill(ds, "Table" & x.ToString())

        Next

        'Return the DataSet
        Return ds


    End Function



    Public Function GetRecordsetDT(ByVal strQuery$) As DataTable
        'Function fills public "DataTable" with results of recordset.  Returns false on error.

        Dim da As SqlDataAdapter

        Dim dt As New DataTable


        da = New SQLDataAdapter(strQuery, _SQLConn)
        da.Fill(dt)

        Return dt





    End Function


    Public Function GetRecordset(ByVal strQuery$) As Boolean
        'Function fills public "DataTable" with results of recordset.  Returns false on error.

        Dim da As SQLDataAdapter
        GetRecordset = False
        If Not DataTable Is Nothing Then DataTable.Clear()


        da = New SQLDataAdapter(strQuery, _SQLConn)
        da.Fill(DataTable)

        GetRecordset = True
        da = Nothing


    End Function

    'Public Sub FillGrid(ByRef grid As ComponentFactory.Krypton.Toolkit.KryptonDataGridView, ByVal ParamArray strFieldList() As Object)
    '    Dim x As Integer, i As Integer
    '    Dim objCells() As Object

    '    grid.Rows.Clear()
    '    With DataTable

    '        ReDim objCells(UBound(strFieldList))
    '        For x = 0 To .Rows.Count - 1

    '            For i = 0 To UBound(strFieldList)

    '                If strFieldList(i) <> "" Then

    '                    objCells(i) = IIf(IsDBNull(.Rows(x).Item(strFieldList(i))), "", .Rows(x).Item(strFieldList(i)))

    '                End If
    '            Next i

    '            grid.Rows.Add(objCells)

    '        Next x

    '    End With




    'End Sub

    Public Function GetLastIdentity() As Integer

        Dim objCommand As SQLCommand
        GetLastIdentity = -1



        objCommand = _SQLConn.CreateCommand()
        objCommand.CommandText = "SELECT @@IDENTITY AS NewID;"

        GetLastIdentity = CInt(objCommand.ExecuteScalar())

        objCommand = Nothing




    End Function


    Public Function ExecuteScalar(ByVal strQuery$) As Object
        'Use ExecuteScalar method to return a single value (for example, an aggregate value) 
        'from the database. This requires less code than using the ExecuteReader method, and then performing 
        'the operations necessary to generate the single value using the data returned by a SqlDataReader.

        'The function returns field object
        Dim objCommand As SQLCommand
        ExecuteScalar = Nothing


        objCommand = _SQLConn.CreateCommand()
        objCommand.CommandText = strQuery

        ExecuteScalar = objCommand.ExecuteScalar()

        objCommand = Nothing




    End Function

    Public Function FixForApostrophe(ByVal strFieldContents As String) As String
        ' replace each apostrophe with two apostrophes

        If Trim(strFieldContents) = "" Then Return ""

        FixForApostrophe = Replace(strFieldContents, Chr(39), "''")

    End Function


    Public Function BuildUpdateQry(ByVal strTableName As String, ByVal ParamArray objFieldsAndValues() As Object) As String
        'Function returns a formatted insert query only. Makes NO connection to the database. 
        Dim qry As String
        Dim x As Int16
        qry = ("Update " _
                    + (strTableName + " Set "))
        x = 0
        Do While (x <= (objFieldsAndValues.GetLength(0) - 1))

            If (objFieldsAndValues((x + 1)) Is Nothing) Then
                objFieldsAndValues((x + 1)) = ""
            End If
            Dim str As String = objFieldsAndValues((x + 1)).GetType.Name
            str = str.ToUpper
            Select Case (str)
                Case "BOOLEAN"
                    If (CType(objFieldsAndValues((x + 1)), Boolean) = True) Then
                        qry = (qry _
                                    + (objFieldsAndValues(x) + "= 1, "))
                    Else
                        qry = (qry _
                                    + (objFieldsAndValues(x) + "= 0, "))
                    End If
                Case "INT32", "INT16", "INT64", "DOUBLE", "SINGLE"
                    qry = (qry _
                                + (objFieldsAndValues(x) & ("=" _
                                + (objFieldsAndValues((x + 1).ToString()) & ", "))))
                Case "STRING", "DATETIME"
                    If (objFieldsAndValues((x + 1)).ToString = "NULL") Then
                        qry = (qry _
                                    + (objFieldsAndValues(x) + "= NULL, "))
                    Else
                        qry = (qry _
                                    + (objFieldsAndValues(x) + ("= '" _
                                    + (FixForApostrophe(objFieldsAndValues((x + 1)).ToString) + "', "))))
                    End If

                Case "GUID"

                    If (objFieldsAndValues((x + 1)).ToString = "NULL") Then
                        qry = (qry _
                                    + (objFieldsAndValues(x) + "= NULL, "))
                    Else
                        qry = (qry _
                                    + (objFieldsAndValues(x) + ("= {" _
                                    + (FixForApostrophe(objFieldsAndValues((x + 1)).ToString) + "}, "))))
                    End If


                Case Else
                    qry = (qry _
                                + (objFieldsAndValues(x) + ("=" _
                                + (objFieldsAndValues((x + 1)) + ", "))))
            End Select
            x = (x + 2)
        Loop
        Return qry.Substring(0, (qry.Length - 2))
    End Function



    Public Function BuildInsertQry(ByVal strTableName As String, ByVal ParamArray strFieldsAndValues() As Object) As String
        'Function returns a formatted insert query only.  Makes NO connection to the database.

        Dim x As Int16
        BuildInsertQry = "Insert Into " & strTableName & " ("
        For x = 0 To UBound(strFieldsAndValues) Step 2

            BuildInsertQry += strFieldsAndValues(x) & ","

        Next
        BuildInsertQry = Left(BuildInsertQry, Len(BuildInsertQry) - 1) & ") VALUES ("

        For x = 1 To UBound(strFieldsAndValues) Step 2

            Select Case UCase(strFieldsAndValues(x).GetType.Name)

                Case "BOOLEAN" : BuildInsertQry += IIf(strFieldsAndValues(x), 1, 0) & ","
                Case "INT64", "INT32", "INT16" : BuildInsertQry += strFieldsAndValues(x) & ","
                Case "STRING", "DATETIME" : BuildInsertQry += "'" & FixForApostrophe(strFieldsAndValues(x)) & "',"
                Case "GUID" : BuildInsertQry &= "{" & strFieldsAndValues(x).ToString() & "},"
                Case "DBNULL" : BuildInsertQry &= "NULL, "
                Case Else : BuildInsertQry += strFieldsAndValues(x) & ","

            End Select

        Next
        BuildInsertQry = Left(BuildInsertQry, Len(BuildInsertQry) - 1) & ")"


    End Function

    Public Function ExecuteNonQuery(ByVal strQuery$) As Boolean
        'No error trapping here.  We want errors to bubble through to the calling function.

        Dim objCommand As SqlCommand


        objCommand = New SqlCommand(strQuery, _SQLConn)
        objCommand.ExecuteNonQuery()
        ExecuteNonQuery = True


        objCommand = Nothing





    End Function


    Protected Overrides Sub Finalize()

        Try

            DataTable = Nothing
          

        Catch ex As Exception
        End Try


        MyBase.Finalize()
    End Sub
End Class
