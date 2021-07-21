Imports Npgsql
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Public Class Form2

    'Private Find1 As String
    'Private Find2 As String
    'Private Find3 As String
    'Private Find4 As Integer
    'Private Find5 As Integer
    'Private Time1 As String
    'Private Time2 As String

    'Dim dt As DataTable = New DataTable()

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        'データベース接続
        'Dim conn As NpgsqlConnection = New NpgsqlConnection(getdatabaseconfig())
        'conn.Open()

        Dim sql As String
        sql = ""
        sql &= " select"
        sql &= " a.item_name as 商品名,"
        sql &= " b.price as 金額,"
        sql &= " c.dealer_name as 業者名,"
        sql &= " d.buy_count 購入数,"
        sql &= " d.buy_date as 購入日,"
        sql &= " a.is_hot_or_cold as 温冷"
        sql &= " from mst_items as a"
        sql &= " left join mst_prices as b"
        sql &= " on a.mst_price_id = b.id"
        sql &= " left join mst_dealers as c"
        sql &= " on a.mst_dealer_id = c.id"
        sql &= " left join trn_claims as d"
        sql &= " on d.mst_item_id =a.id"
        sql &= " where d.id is not null"

        Dim dt As DataTable = New DataTable()

        'SQL実行
        Using db As New dbConnection()
            dt = db.getDtSql(sql)
        End Using

        Dim i As Integer = 0

        dt.Columns.Add("温冷2", GetType(String))
        'dt.Columns("温冷2").DataType = GetType(String)


        For Each row As DataRow In dt.Rows
            If row("温冷") = 1 Then
                dt.Rows(i).Item("温冷2") = "冷"
            Else
                dt.Rows(i).Item("温冷2") = "温"
            End If
            i += 1
        Next

        dt.Columns.Remove("温冷")

        'dt.Columns("温冷2").SetOrdinal(0)
        'dt.Columns("商品名").SetOrdinal(1)


        DataGridView1.DataSource = dt

        '' SQL実行
        'Dim command As NpgsqlCommand = New NpgsqlCommand(sql, conn)
        'Dim adapter As NpgsqlDataAdapter = New NpgsqlDataAdapter(command)
        'Dim dt As DataTable = New DataTable()
        'adapter.Fill(dt)

        'DataGridView1.Columns.Add("item_name", "商品名")
        'DataGridView1.Columns.Add("buy_price", "金額")
        'DataGridView1.Columns.Add("dealer_name", "販売元")
        'DataGridView1.Columns.Add("buy_count", "購入数")
        'DataGridView1.Columns.Add("buy_date", "購入日")
        'DataGridView1.Columns.Add("is_hot_or_cold", "温冷")

        'For Each row As DataRow In dt.Rows
        '    Dim itemname As String = row("item_name")
        '    Dim price As Integer = row("price")
        '    Dim dealer As String = row("dealer_name")
        '    Dim buycount As Integer = row("buy_count")
        '    Dim buydate As String = row("buy_date")
        '    Dim hotorcold As String = row("is_hot_or_cold")

        '    If row("is_hot_or_cold") = 1 Then
        '        hotorcold = "cold"
        '    Else
        '        hotorcold = "hot"
        '    End If

        '    DataGridView1.Rows.Add(itemname, price, dealer, buycount, buydate, hotorcold)
        'Next



        'conn.Close()

        Dim dtCombo As New DataTable
        dtCombo.Columns.Add("id")
        dtCombo.Columns.Add("name")

        Dim dtRowCombo As DataRow
        dtRowCombo = dtCombo.NewRow
        dtRowCombo("id") = ""
        dtRowCombo("name") = ""
        dtCombo.Rows.Add(dtRowCombo)

        dtRowCombo = dtCombo.NewRow
        dtRowCombo("id") = "1"
        dtRowCombo("name") = "cold"
        dtCombo.Rows.Add(dtRowCombo)


        dtRowCombo = dtCombo.NewRow
        dtRowCombo("id") = "2"
        dtRowCombo("name") = "hot"
        dtCombo.Rows.Add(dtRowCombo)

        ComboBox1.DataSource = dtCombo
        ComboBox1.DisplayMember = "name"
        ComboBox1.ValueMember = "id"

        Dim dtCombo2 As New DataTable
        dtCombo2.Columns.Add("id")
        dtCombo2.Columns.Add("name")

        Dim dtRowCombo2 As DataRow
        dtRowCombo2 = dtCombo2.NewRow
        dtRowCombo2("id") = ""
        dtRowCombo2("name") = "無制限"
        dtCombo2.Rows.Add(dtRowCombo2)

        dtRowCombo2 = dtCombo2.NewRow
        dtRowCombo2("id") = "1"
        dtRowCombo2("name") = "50"
        dtCombo2.Rows.Add(dtRowCombo2)


        dtRowCombo2 = dtCombo2.NewRow
        dtRowCombo2("id") = "2"
        dtRowCombo2("name") = "100"
        dtCombo2.Rows.Add(dtRowCombo2)

        dtRowCombo2 = dtCombo2.NewRow
        dtRowCombo2("id") = "3"
        dtRowCombo2("name") = "500"
        dtCombo2.Rows.Add(dtRowCombo2)

        ComboBox2.DataSource = dtCombo2
        ComboBox2.DisplayMember = "name"
        ComboBox2.ValueMember = "id"



    End Sub



    '検索ボタン
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Dim conn As NpgsqlConnection = New NpgsqlConnection(getdatabaseconfig())
        'conn.Open()

        Dim Split() As String = {}

        Dim find1 As String = TextBox1.Text
        Dim find2 As String = TextBox2.Text
        Dim find3 As String = TextBox3.Text
        Dim find4 As Integer = ComboBox1.SelectedIndex
        Dim find5 As Integer = ComboBox2.SelectedIndex
        Dim time1 As String = DateTimePicker1.Value.ToString("yyyy/MM/dd")
        Dim time2 As String = DateTimePicker2.Value.ToString("yyyy/MM/dd")


        'Dim ww = Find3.Split(" ")
        Dim rep As String = Find3.Replace("　", " ")

        'Dim Split() = rep.Split(" ")
        If rep <> "" Then
            Split = rep.Split(" ")
        End If

        Dim count As Integer = 0


        Dim sql As String
        sql = ""
        sql &= " select"
        sql &= " a.item_name, b.price, c.dealer_name, d.buy_count, d.buy_date, is_hot_or_cold"
        sql &= " from mst_items as a"
        sql &= " left join mst_prices as b"
        sql &= " on a.mst_price_id = b.id"
        sql &= " left join mst_dealers as c"
        sql &= " on a.mst_dealer_id = c.id"
        sql &= " left join trn_claims as d"
        sql &= " on d.mst_item_id =a.id"
        sql &= " where 1=1 "

        If Find1 <> "" Then
            sql &= " and a.item_name Like'%" & Find1 & "%' "
        End If
        If Find2 <> "" Then
            sql &= " and"
            sql &= " c.dealer_name like'%" & Find2 & "%'"
        End If

        For Each value As String In Split

            If count = 0 Then
                sql &= " And "
                sql &= "(a.item_name Like '%" & value & "%' "
                sql &= "or "
                sql &= "c.dealer_name Like '%" & value & "%' )"
                count = count + 1
            Else
                sql &= "and"
                sql &= " (a.item_name LIKE '%" & value & "%' "
                sql &= "or "
                sql &= "c.dealer_name Like '%" & value & "%') "
            End If

        Next

        sql &= " and"
        sql &= " buy_date BETWEEN '" & Time1 & " 00:00:00' AND '" & Time2 & " 23:59:59' "

        '温冷
        If ComboBox1.SelectedValue <> "" Then
            sql &= " and a.is_hot_or_cold = '" & ComboBox1.SelectedValue & "' "
        End If

        '表示件数
        If Find5 = 1 Then
            sql &= " limit 50"
        ElseIf Find5 = 2 Then
            sql &= " limit 100"
        ElseIf Find5 = 3 Then
            sql &= " FETCH FIRST 500 ROWS ONLY"
        End If

        '集計
        If CheckBox1.Checked = True Then
            sql = ""
            sql &= " select"
            sql &= " a.item_name,"
            sql &= " buy_price,"
            'sql &= " sum(buy_price*buy_count) as total_price,"
            sql &= " c.dealer_name,"
            'sql &= " sum(d.buy_count) as total_buy_count,"
            sql &= " to_char(d.buy_date, 'yyyy/mm/dd') as buy_date,"
            sql &= " is_hot_or_cold,"
            sql &= " sum(buy_price*buy_count) as total_price,"
            sql &= " sum(d.buy_count) as total_buy_count"
            sql &= " from mst_items as a"
            sql &= " left join mst_prices as b"
            sql &= " on a.mst_price_id = b.id"
            sql &= " left join mst_dealers as c"
            sql &= " on a.mst_dealer_id = c.id"
            sql &= " left join trn_claims as d"
            sql &= " on d.mst_item_id = a.id"
            sql &= " where 1 = 1"
            sql &= " and to_char(d.buy_date, 'yyyy/mm/dd hh24:mm:ss') BETWEEN '" & Time1 & " 00:00:00' AND '" & Time2 & " 23:59:59' "
            sql &= " group by a.item_name, buy_price, c.dealer_name, d.buy_count, to_char(d.buy_date, 'yyyy/mm/dd'), is_hot_or_cold"
        End If

        Dim dt As DataTable = New DataTable()

        'SQL実行
        Using db As New dbConnection()
            dt = db.getDtSql(sql)
        End Using

        DataGridView1.DataSource = dt

        'Dim command As NpgsqlCommand = New NpgsqlCommand(sql, conn)
        'Dim adapter As NpgsqlDataAdapter = New NpgsqlDataAdapter(command)
        'Dim dt As DataTable = New DataTable()
        'adapter.Fill(dt)

        'DataGridView1.Rows.Clear()
        'DataGridView1.Columns.Clear()

        'If CheckBox1.Checked = True Then
        '    DataGridView1.Columns.Add("item_name", "商品名")
        '    DataGridView1.Columns.Add("buy_price", "金額")
        '    DataGridView1.Columns.Add("dealer_name", "販売元")
        '    DataGridView1.Columns.Add("buy_date", "購入日")
        '    DataGridView1.Columns.Add("is_hot_or_cold", "温冷")
        '    DataGridView1.Columns.Add("total_price", "合計金額")
        '    DataGridView1.Columns.Add("total_buy_count", "合計購入数")
        'Else
        '    DataGridView1.Columns.Add("item_name", "商品名")
        '    DataGridView1.Columns.Add("buy_price", "金額")
        '    DataGridView1.Columns.Add("dealer_name", "販売元")
        '    DataGridView1.Columns.Add("buy_count", "購入数")
        '    DataGridView1.Columns.Add("buy_date", "購入日")
        '    DataGridView1.Columns.Add("is_hot_or_cold", "温冷")
        'End If

        'If CheckBox1.Checked = True Then
        '    For Each row As DataRow In dt.Rows
        '        Dim itemname As String = row("item_name")
        '        Dim price As Integer = row("buy_price")
        '        Dim dealer As String = row("dealer_name")
        '        Dim buydate As String = row("buy_date")
        '        Dim hotorcold As String = row("is_hot_or_cold")

        '        If row("is_hot_or_cold") = 1 Then
        '            hotorcold = "cold"
        '        Else
        '            hotorcold = "hot"
        '        End If

        '        Dim totalprice As Integer = row("total_price")
        '        Dim totalbuycount As Integer = row("total_buy_count")
        '        DataGridView1.Rows.Add(itemname, price, dealer, buydate, hotorcold, totalprice, totalbuycount)
        '    Next
        'Else
        '    For Each row As DataRow In dt.Rows
        '        Dim itemname As String = row("item_name")
        '        Dim price As Integer = row("price")
        '        Dim dealer As String = row("dealer_name")
        '        Dim buycount As Integer = row("buy_count")
        '        Dim buydate As String = row("buy_date")
        '        Dim hotorcold As String = row("is_hot_or_cold")

        '        If row("is_hot_or_cold") = 1 Then
        '            hotorcold = "cold"
        '        Else
        '            hotorcold = "hot"
        '        End If

        '        DataGridView1.Rows.Add(itemname, price, dealer, buycount, buydate, hotorcold)
        '    Next
        'End If

        'conn.Close()

    End Sub

    'CSV出力
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
        Using sfd As SaveFileDialog = New SaveFileDialog
            'デフォルトのファイル名を指定します
            sfd.FileName = "output.csv"
            If sfd.ShowDialog() = DialogResult.OK Then
                'Using writer As New StreamWriter("\aaa\bbb.csv", False, Encoding.GetEncoding("shift_jis"))
                Using writer As New StreamWriter(sfd.FileName, False, Encoding.GetEncoding("shift_jis"))

                    Dim rowCount As Integer = DataGridView1.Rows.Count

                    ' ユーザによる行追加が許可されている場合は、最後の新規入力用の
                    ' 1行分を差し引く
                    If (DataGridView1.AllowUserToAddRows = True) Then
                        rowCount = rowCount - 1
                    End If

                    '見出し
                    Dim strList1 As New List(Of String)
                    For i = 0 To DataGridView1.ColumnCount - 1
                        strList1.Add(DataGridView1.Columns(i).HeaderText.ToString)
                    Next
                    Dim strary1 As String() = strList1.ToArray
                    Dim strCsvData1 As String = String.Join(",", strary1)
                    writer.WriteLine(strCsvData1)

                    ' 行
                    For i As Integer = 0 To rowCount - 1
                        ' リストの初期化
                        Dim strList As New List(Of String)

                        ' 列
                        For j As Integer = 0 To DataGridView1.Columns.Count - 1
                            strList.Add(DataGridView1(j, i).Value.ToString())
                        Next
                        Dim strArray As String() = strList.ToArray() ' 配列へ変換

                        ' CSV 形式に変換
                        Dim strCsvData As String = String.Join(",", strArray)

                        writer.WriteLine(strCsvData)
                    Next
                    'MessageBox.Show("CSV ファイルを出力しました")
                End Using
            End If
        End Using
    End Sub

    'Dim where As String = ""
    'If TextBox3.Text <> "" Then
    '    ' 入力されている場合
    '    ' 入力値を代入
    '    Find3 = TextBox3.Text
    '    ' 全角空白を半角空白に置換
    '    Dim rep As String = Find3.Replace("　", " ")
    '    ' 半角空白を区切り文字とさせて、区切った結果をリストに設定
    '    Dim Split() = rep.Split(" ")

    '    ' WHEREの条件用文字列を成型
    '    where &= " AND "
    '    For i As Integer = 0 To UBound(Split)

    '        If i > 0 Then
    '            where &= " AND "
    '        End If
    '        where &= " ( "
    '        where &= "   a.item_name LIKE '%" & Split(i) & "%' "
    '        where &= "   OR"
    '        where &= "   c.dealer_name LIKE '%" & Split(i) & "%' "
    '        where &= " ) "
    '    Next

    'End If

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