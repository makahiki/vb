Imports Npgsql
Imports System.IO
Imports System.Text

Public Class Form4

    Dim dt As DataTable = New DataTable()
    Dim hash As New Hashtable
    Dim aa As Integer() = New Integer() {}

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim sql As String
        sql = ""
        sql &= " select"
        sql &= " c.dealer_name as 業者名,"
        sql &= " a.item_name as 商品名,"
        sql &= " d.stock_count as 在庫,"
        sql &= " CEIL(sum(CEIL(20-stock_count)/4)) as 補充（箱）,"
        sql &= " CEIL(sum(CEIL(20-stock_count)/4))*4 - (20-stock_count) as あまり"
        sql &= " FROM mst_items as a "
        sql &= " left join mst_prices as b"
        sql &= " on a.mst_price_id = b.id"
        sql &= " left join mst_dealers as c"
        sql &= " on a.mst_dealer_id = c.id"
        sql &= " left join trn_stocks as d"
        sql &= " on a.id = d.mst_item_id"
        sql &= " where stock_count <=10"
        sql &= " group by a.id, a.item_name, c.dealer_name, d.stock_count"
        sql &= " order by random()"
        'sql &= " order by a.item_name asc"

        'SQL実行
        Using db As New dbConnection()
            dt = db.getDtSql(sql)
        End Using

        Dim rowCount As Integer = 0

        For Each row As DataRow In dt.Rows

            hash.Add(row("商品名"), row("あまり"))
            rowCount += 1
        Next

        DataGridView1.DataSource = dt

        'DataGridView1.Columns.Add("dealer_name", "業者")
        'DataGridView1.Columns.Add("item_name", "商品名")
        'DataGridView1.Columns.Add("stock_count", "在庫")
        'DataGridView1.Columns.Add("box_count", "補充（箱)")
        'DataGridView1.Columns.Add("stock_remainder", "あまり")

        'For Each row As DataRow In dt.Rows
        '    Dim dealername As String = row("dealer_name")
        '    Dim itemname As String = row("item_name")
        '    Dim stockcount As Integer = row("stock_count")
        '    Dim boxcount As Integer = row("box_count")
        '    'Dim stockremainder As Integer = row("stock_remainder")

        '    ReDim Preserve aa(UBound(aa) + 1)
        '    aa(UBound(aa)) = row("stock_remainder")

        '    DataGridView1.Rows.Add(dealername, itemname, stockcount, boxcount)
        'Next

    End Sub

    'CSV出力
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim dv = New DataView(dt)
        dv.Sort = "業者名"
        dt = dv.ToTable

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        Using sfd As SaveFileDialog = New SaveFileDialog
            'デフォルトのファイル名を指定します
            sfd.FileName = "output.csv"
            If sfd.ShowDialog() = DialogResult.OK Then

                Using writer As New StreamWriter(sfd.FileName, False, Encoding.GetEncoding("shift_jis"))

                    Dim rowCount As Integer = dt.Rows.Count

                    ' ユーザによる行追加が許可されている場合は、最後の新規入力用の
                    ' 1行分を差し引く
                    'If (DataGridView1.AllowUserToAddRows = True) Then
                    '    rowCount = rowCount - 1
                    'End If

                    '見出し
                    Dim strList1 As New List(Of String)
                    For i = 0 To dt.Columns.Count - 1
                        strList1.Add(dt.Columns(i).ToString)
                    Next
                    Dim strary1 As String() = strList1.ToArray
                    Dim strCsvData1 As String = String.Join(",", strary1)
                    writer.WriteLine(strCsvData1)

                    ' 行
                    For i As Integer = 0 To dt.Rows.Count - 1
                        ' リストの初期化
                        Dim strList As New List(Of String)

                        ' 列
                        For j As Integer = 0 To dt.Columns.Count - 1
                            'strList.Add(dt(j)(i).Value.ToString())
                            strList.Add(dt(i)(j).ToString())
                        Next

                        Dim strArray As String() = strList.ToArray() ' 配列へ変換

                        ' CSV 形式に変換
                        Dim strCsvData As String = String.Join(",", strArray)

                        writer.WriteLine(strCsvData)

                    Next
                End Using
            End If
        End Using

    End Sub

    '補充
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim sql As String

        sql = ""
        sql &= " update trn_stocks"
        sql &= " set stock_count = 20"

        ' SQL実行
        Using db As New dbConnection()
            db.getDtSql(sql)
        End Using

        sql = ""
        sql &= " select"
        sql &= " c.dealer_name as 業者名,"
        sql &= " a.item_name as 商品名,"
        sql &= " d.stock_count as 在庫,"
        sql &= " CEIL(sum(CEIL(20-stock_count)/4)) as 補充（箱）"
        'sql &= " CEIL(sum(CEIL(20-stock_count)/4))*4 - (20-stock_count) as あまり"
        sql &= " FROM mst_items as a "
        sql &= " left join mst_prices as b"
        sql &= " on a.mst_price_id = b.id"
        sql &= " left join mst_dealers as c"
        sql &= " on a.mst_dealer_id = c.id"
        sql &= " left join trn_stocks as d"
        sql &= " on a.id = d.mst_item_id"
        sql &= " group by a.id, a.item_name, c.dealer_name, d.stock_count"
        sql &= " order by a.item_name asc"

        ' SQL実行
        Using db As New dbConnection()
            dt = db.getDtSql(sql)
        End Using

        dt.Columns.Add("あまり")

        Dim rowCount As Integer = 0
        Dim columnCount As Integer = dt.Columns.Count - 1

        For Each row As DataRow In dt.Rows
            dt(rowCount)(columnCount) = hash(row("商品名"))
            rowCount += 1
        Next

        DataGridView1.DataSource = dt

    End Sub

    '補充
    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    'データベース接続
    '    'Dim conn As NpgsqlConnection = New NpgsqlConnection(getdatabaseconfig())
    '    'conn.Open()

    '    Dim sql As String

    '    'Dim cmd As New NpgsqlCommand

    '    sql = ""
    '    sql &= " update trn_stocks"
    '    sql &= " set stock_count = 20"

    '    'cmd.Connection = conn
    '    'cmd.CommandText = sql
    '    'cmd.ExecuteNonQuery()

    '    Using db As New dbConnection()
    '        dt = db.getDtSql(sql)
    '    End Using

    '    sql = ""
    '    sql &= " select"
    '    sql &= " c.dealer_name ,a.item_name, d.stock_count,"
    '    sql &= " CEIL(sum(CEIL(20-stock_count)/4)) as box_count"
    '    'sql &= " CEIL(sum(CEIL(20-stock_count)/4))*4 - (20-stock_count) as stock_remainder"
    '    sql &= " FROM mst_items as a "
    '    sql &= " left join mst_prices as b"
    '    sql &= " on a.mst_price_id = b.id"
    '    sql &= " left join mst_dealers as c"
    '    sql &= " on a.mst_dealer_id = c.id"
    '    sql &= " left join trn_stocks as d"
    '    sql &= " on a.id = d.mst_item_id"
    '    'sql &= " where stock_count <=10"
    '    sql &= " group by a.id, a.item_name, c.dealer_name, d.stock_count"

    '    '' SQL実行
    '    'Dim command As NpgsqlCommand = New NpgsqlCommand(sql, conn)
    '    'Dim adapter As NpgsqlDataAdapter = New NpgsqlDataAdapter(command)
    '    'Dim dt As DataTable = New DataTable()
    '    'adapter.Fill(dt)

    '    Using db As New dbConnection()
    '        dt = db.getDtSql(sql)
    '    End Using


    '    DataGridView1.Rows.Clear()
    '    DataGridView1.Columns.Clear()

    '    Dim i As Integer

    '    DataGridView1.Columns.Add("dealer_name", "業者")
    '    DataGridView1.Columns.Add("item_name", "商品名")
    '    DataGridView1.Columns.Add("stock_count", "在庫")
    '    DataGridView1.Columns.Add("box_count", "補充（箱)")
    '    DataGridView1.Columns.Add("stock_remainder", "あまり")

    '    For Each row As DataRow In dt.Rows
    '        Dim dealername As String = row("dealer_name")
    '        Dim itemname As String = row("item_name")
    '        Dim stockcount As Integer = row("stock_count")
    '        Dim boxcount As Integer = row("box_count")
    '        'Dim stockremainder As Integer = row("stock_remainder")

    '        'ReDim Preserve aa(UBound(aa) + 1)
    '        'aa(UBound(aa)) = row("stock_remainder")

    '        DataGridView1.Rows.Add(dealername, itemname, stockcount, boxcount, aa(i))
    '        i = i + 1
    '    Next

    '    'conn.Close()
    'End Sub

    '' データベース接続用
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

End Class