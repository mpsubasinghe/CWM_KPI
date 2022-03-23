Imports System.Data.SqlClient
Imports System.Data.OleDb

Module Module1

    Public DownLoad_A_Code As String = ""

    Public Areacode As String = ""
    Public AreaName As String = ""
    Public Head As String = ""
    Public RegMan As String = ""
    Public SupsCode As String = ""
    Public Region1 As String = ""
    Public Stockist1 As String = ""


    Public comcode As String = ""
    Public PDAInv As String = ""
    Public Sectors As String = ""
    Public Route As String = ""
    Public RetailerName As String = ""
    Public RepName As String = ""
    Public StkName As String = ""
    Public ItemName As String = ""
    Public SEQ As String = ""


    Public SectorID As Integer = 0
    Public DATINV As String = ""
    Public RouteID As String = ""
    Public SectorName As String = ""


    Public CATE As String = ""

    ' Public AS400Str As String = "Provider=IBMDA400;Data Source=192.168.190.2;User Id=" & Trim(Login.UIDtxt.Text) & ";Password=" & Trim(Login.PWDtxt.Text) & ";"
    'Public AS400Str As String = ""
    'Public InvNo As String
    'Public Dattime As String
    'Public prg As Integer
    'Public mon As String
    'Public SecID As String
    'Public Route As String
    'Public online As Boolean'
    Dim directory As String = My.Application.Info.DirectoryPath

    ' Public aceesscon As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & directory & "\database\DBCMarketing.mdb;Jet OLEDB:Database Password=redrock;")
    Public aceesscon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + directory + "\Database\DBCMarketing.mdb;Jet OLEDB:Database Password=redrock;")
    'Public aceesscon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & directory & "\database\DBCMarketing1.mdb;Jet OLEDB:System Database=system.mdw;User ID=admin;Password=redrock;")
    'Public aceesscon As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & directory & "\database\SFA.mdb;Jet OLEDB:Database Password=dbclmis321;")


    ' Public cona As New OleDbConnection("Provider=IBMDA400;Data Source=192.168.190.2;User Id=CALS;Password=CALS123;")
    'Public cona As New OleDbConnection(AS400Str)
    ' Public con As New SqlConnection("Data Source=SQL5005.Smarterasp.net;Initial Catalog=DB_9AAA2F_Nisiposha;User Id=DB_9AAA2F_Nisiposha_admin;Password=mano1234;")

    Public con As New SqlConnection("Data Source=SQL5006.Smarterasp.net;Initial Catalog=DB_9AAA2F_Consumer;User Id=DB_9AAA2F_Consumer_admin;Password=mano1234;Connection Timeout=360")


    ' Public con As New SqlConnection("Data source = localhost\SQLEXPRESS;Initial catalog =Consumer1;integrated security = true")
 
    Function GetDataSQL(ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, ByVal p4 As String, ByVal p5 As String, ByVal pn1 As String, ByVal pn2 As String, ByVal pn3 As String, ByVal pn4 As String, ByVal pn5 As String, ByVal sql As String) As System.Data.DataSet
        Try

            ' con.Open()
            Dim queryString As String = sql
            Dim dbCommand As System.Data.IDbCommand = New System.Data.SqlClient.SqlCommand
            dbCommand.CommandText = queryString
            dbCommand.Connection = con

            Dim dbParam_p1 As System.Data.IDataParameter = New System.Data.SqlClient.SqlParameter
            dbParam_p1.ParameterName = pn1
            dbParam_p1.Value = p1
            dbParam_p1.DbType = System.Data.DbType.String
            dbCommand.Parameters.Add(dbParam_p1)

            Dim dbParam_p2 As System.Data.IDataParameter = New System.Data.SqlClient.SqlParameter
            dbParam_p2.ParameterName = pn2
            dbParam_p2.Value = p2
            dbParam_p2.DbType = System.Data.DbType.String
            dbCommand.Parameters.Add(dbParam_p2)

            Dim dbParam_p3 As System.Data.IDataParameter = New System.Data.SqlClient.SqlParameter
            dbParam_p3.ParameterName = pn3
            dbParam_p3.Value = p3
            dbParam_p3.DbType = System.Data.DbType.String
            dbCommand.Parameters.Add(dbParam_p3)

            Dim dataAdapter As System.Data.IDbDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            dataAdapter.SelectCommand = dbCommand
            Dim dataSet As System.Data.DataSet = New System.Data.DataSet
            dataAdapter.Fill(dataSet)

            'con.Close()
            Return dataSet

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try

    End Function


    Function GetDataACC(ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, ByVal p4 As String, ByVal p5 As String, ByVal pn1 As String, ByVal pn2 As String, ByVal pn3 As String, ByVal pn4 As String, ByVal pn5 As String, ByVal sql As String) As System.Data.DataSet
        ' MsgBox(sql)
        Dim queryString As String = sql
        Dim dbCommand As System.Data.IDbCommand = New System.Data.OleDb.OleDbCommand
        dbCommand.CommandText = queryString
        dbCommand.Connection = aceesscon

        Dim dbParam_p1 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
        dbParam_p1.ParameterName = pn1
        dbParam_p1.Value = p1
        dbParam_p1.DbType = System.Data.DbType.String
        dbCommand.Parameters.Add(dbParam_p1)

        Dim dbParam_p2 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
        dbParam_p2.ParameterName = pn2
        dbParam_p2.Value = p2
        dbParam_p2.DbType = System.Data.DbType.String
        dbCommand.Parameters.Add(dbParam_p2)

        Dim dbParam_p3 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
        dbParam_p3.ParameterName = pn3
        dbParam_p3.Value = p3
        dbParam_p3.DbType = System.Data.DbType.String
        dbCommand.Parameters.Add(dbParam_p3)

        Dim dbParam_p4 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
        dbParam_p4.ParameterName = pn4
        dbParam_p4.Value = p4
        dbParam_p4.DbType = System.Data.DbType.String
        dbCommand.Parameters.Add(dbParam_p4)

        Dim dbParam_p5 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
        dbParam_p5.ParameterName = pn5
        dbParam_p5.Value = p5
        dbParam_p5.DbType = System.Data.DbType.String
        dbCommand.Parameters.Add(dbParam_p5)

        Dim dataAdapter As System.Data.IDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter
        dataAdapter.SelectCommand = dbCommand
        Dim dataSet As System.Data.DataSet = New System.Data.DataSet
        dataAdapter.Fill(dataSet)

        Return dataSet

    End Function
    'Function InsertDataACC(ByVal p1 As String, ByVal p2 As String, ByVal p3 As String, ByVal p4 As String, ByVal p5 As String, ByVal pn1 As String, ByVal pn2 As String, ByVal pn3 As String, ByVal pn4 As String, ByVal pn5 As String, ByVal sql As String) As System.Data.DataSet
    '     MsgBox(sql)
    '    Dim queryString As String = sql
    '    Dim dbCommand As System.Data.IDbCommand = New System.Data.OleDb.OleDbCommand
    '    dbCommand.CommandText = queryString
    '    dbCommand.Connection = aceesscon

    '    Dim dbParam_p1 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
    '    dbParam_p1.ParameterName = pn1
    '    dbParam_p1.Value = p1
    '    dbParam_p1.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p1)

    '    Dim dbParam_p2 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
    '    dbParam_p2.ParameterName = pn2
    '    dbParam_p2.Value = p2
    '    dbParam_p2.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p2)

    '    Dim dbParam_p3 As System.Data.IDataParameter = New System.Data.OleDb.OleDbParameter
    '    dbParam_p3.ParameterName = pn3
    '    dbParam_p3.Value = p3
    '    dbParam_p3.DbType = System.Data.DbType.String
    '    dbCommand.Parameters.Add(dbParam_p3)

    '    Dim dataAdapter As System.Data.IDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter
    '    dataAdapter.SelectCommand = dbCommand
    '    Dim dataSet As System.Data.DataSet = New System.Data.DataSet
    '    dataAdapter.Fill(dataSet)

    '       Return dataSet

    'End Function

End Module

