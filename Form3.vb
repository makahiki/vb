Imports Npgsql

Public Class Form3
    'Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    'データベース接続
    '    Dim conn As NpgsqlConnection = New NpgsqlConnection(getdatabaseconfig())
    '    conn.Open()

    '    Dim sql As String
    '    sql = ""
    '    sql &= " select *"
    '    sql &= " from coin_bill_stocks"
    '    sql &= " order by id asc"

    '    ' SQL実行
    '    Dim command As NpgsqlCommand = New NpgsqlCommand(sql, conn)
    '    Dim adapter As NpgsqlDataAdapter = New NpgsqlDataAdapter(command)
    '    Dim dt As DataTable = New DataTable()
    '    adapter.Fill(dt)

    '    DataGridView1.Columns.Add("coin_bill_name", "金額")
    '    DataGridView1.Columns.Add("coin_bill_stock_count", "在庫")

    '    For Each row As DataRow In dt.Rows
    '        Dim coin_name As String = row("coin_bill_name")
    '        Dim coin_count As Integer = row("coin_bill_stock_count")
    '        DataGridView1.Rows.Add(coin_name, coin_count)
    '    Next

    '    conn.Close()
    'End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim sql As String
        sql = ""
        sql &= " select *"
        sql &= " from coin_bill_stocks"
        sql &= " order by id asc"

        Dim dt As DataTable = New DataTable()

        Using db As New dbConnection()

            dt = db.getDtSql(sql)

        End Using

        DataGridView1.Columns.Add("coin_bill_name", "金額")
        DataGridView1.Columns.Add("coin_bill_stock_count", "在庫")

        For Each row As DataRow In dt.Rows
            Dim coin_name As String = row("coin_bill_name")
            Dim coin_count As Integer = row("coin_bill_stock_count")
            DataGridView1.Rows.Add(coin_name, coin_count)
        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim sql As String

        sql = ""
        sql &= " select *"
        sql &= " from coin_bill_stocks"
        sql &= " order by id asc"

        Dim dt As DataTable = New DataTable()



        '集計
        If CheckBox1.Checked = True Then
            sql = ""
            sql &= " select"
            sql &= " coin_bill_name,"
            sql &= " coin_bill_stock_count,"
            sql &= " sum(coin_bill_name*coin_bill_stock_count) as total"
            sql &= " from coin_bill_stocks"
            sql &= " group by id, coin_bill_name, coin_bill_stock_count"
            sql &= " order by id asc"
        End If

        Using db As New dbConnection()

            dt = db.getDtSql(sql)

        End Using

        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        If CheckBox1.Checked = True Then
            DataGridView1.Columns.Add("coin_bill_name", "金額")
            DataGridView1.Columns.Add("coin_bill_stock_count", "在庫")
            DataGridView1.Columns.Add("total", "合計")
        Else
            DataGridView1.Columns.Add("coin_bill_name", "金額")
            DataGridView1.Columns.Add("coin_bill_stock_count", "在庫")
        End If

        If CheckBox1.Checked = True Then
            For Each row As DataRow In dt.Rows
                Dim coin_name As String = row("coin_bill_name")
                Dim coin_count As Integer = row("coin_bill_stock_count")
                Dim total As Integer = row("total")
                DataGridView1.Rows.Add(coin_name, coin_count, total)
            Next
        Else
            For Each row As DataRow In dt.Rows
                Dim coin_name As String = row("coin_bill_name")
                Dim coin_count As Integer = row("coin_bill_stock_count")
                DataGridView1.Rows.Add(coin_name, coin_count)
            Next
        End If

    End Sub

    ' データベース接続用
    'Private Function getdatabaseconfig()

    '    Dim builder As New Npgsql.NpgsqlConnectionStringBuilder
    '    With builder
    '        .Host = "localhost"
    '        .Database = "vending_machine"
    '        .Username = "postgres"
    '        .Password = "postgres"
    '        .Port = 5432
    '    End With

    '    Return builder.ConnectionString

    'End Function

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    'データベース接続
    '    Dim conn As NpgsqlConnection = New NpgsqlConnection(getdatabaseconfig())
    '    conn.Open()

    '    Dim sql As String

    '    sql = ""
    '    sql &= " select *"
    '    sql &= " from coin_bill_stocks"
    '    sql &= " order by id asc"

    '    '集計
    '    If CheckBox1.Checked = True Then
    '        sql = ""
    '        sql &= " select"
    '        sql &= " coin_bill_name,"
    '        sql &= " coin_bill_stock_count,"
    '        sql &= " sum(coin_bill_name*coin_bill_stock_count) as total"
    '        sql &= " from coin_bill_stocks"
    '        sql &= " group by id, coin_bill_name, coin_bill_stock_count"
    '        sql &= " order by id asc"
    '    End If

    '    Dim command As NpgsqlCommand = New NpgsqlCommand(sql, conn)
    '    Dim adapter As NpgsqlDataAdapter = New NpgsqlDataAdapter(command)
    '    Dim dt As DataTable = New DataTable()
    '    adapter.Fill(dt)

    '    DataGridView1.Rows.Clear()
    '    DataGridView1.Columns.Clear()

    '    If CheckBox1.Checked = True Then
    '        DataGridView1.Columns.Add("coin_bill_name", "金額")
    '        DataGridView1.Columns.Add("coin_bill_stock_count", "在庫")
    '        DataGridView1.Columns.Add("total", "合計")
    '    Else
    '        DataGridView1.Columns.Add("coin_bill_name", "金額")
    '        DataGridView1.Columns.Add("coin_bill_stock_count", "在庫")
    '    End If

    '    If CheckBox1.Checked = True Then
    '        For Each row As DataRow In dt.Rows
    '            Dim coin_name As String = row("coin_bill_name")
    '            Dim coin_count As Integer = row("coin_bill_stock_count")
    '            Dim total As Integer = row("total")
    '            DataGridView1.Rows.Add(coin_name, coin_count, total)
    '        Next
    '    Else
    '        For Each row As DataRow In dt.Rows
    '            Dim coin_name As String = row("coin_bill_name")
    '            Dim coin_count As Integer = row("coin_bill_stock_count")
    '            DataGridView1.Rows.Add(coin_name, coin_count)
    '        Next
    '    End If

    'End Sub
End Class