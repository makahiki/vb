Imports Npgsql

Public Class Form5

    Private ItemId As Integer
    Private AddId As Integer
    Private Cold As String
    Private Hot As Integer

    ' データベース接続用
    Private Function getdatabaseconfig()

        Dim builder As New NpgsqlConnectionStringBuilder
        With builder
            .Host = "localhost"
            .Database = "vending_machine"
            .Username = "postgres"
            .Password = "postgres"
            .Port = 5432
        End With

        Return builder.ConnectionString

    End Function
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'データベース接続
        Dim conn As NpgsqlConnection = New NpgsqlConnection(getdatabaseconfig())
        conn.Open()

        Dim sql As String
        sql = ""
        sql &= " select"
        sql &= " a.id, a.item_name, b.price, c.dealer_name, a.is_hot_or_cold"
        sql &= " from mst_items as a"
        sql &= " left join mst_prices as b"
        sql &= " on a.mst_price_id = b.id"
        sql &= " left join mst_dealers as c"
        sql &= " on a.mst_dealer_id = c.id"
        sql &= " order by a.id asc"

        ' SQL実行
        Dim command As NpgsqlCommand = New NpgsqlCommand(sql, conn)
        Dim adapter As NpgsqlDataAdapter = New NpgsqlDataAdapter(command)
        Dim dt As DataTable = New DataTable()
        adapter.Fill(dt)

        Dim row As DataRow = dt.Rows(0)
        ItemId = row("id")

        '最大値取得
        '''''''''''''''''''''''''''''''''''''''''''''''''
        Dim sql2 As String
        sql2 = ""
        sql2 &= " select"
        sql2 &= " max(id)"
        sql2 &= " from mst_items"

        ' SQL実行
        Dim command2 As NpgsqlCommand = New NpgsqlCommand(sql2, conn)
        Dim adapter2 As NpgsqlDataAdapter = New NpgsqlDataAdapter(command2)
        Dim dt2 As DataTable = New DataTable()
        adapter2.Fill(dt2)

        Dim row2 As DataRow = dt2.Rows(0)
        row2("max") += 1
        AddId = row2("max")
        Label7.Text = AddId

        ''''''''''''''''''''''''''''''''''''''''''''''''''

        '金額コンボ
        Dim dtCombo As New DataTable
        dtCombo.Columns.Add("id")
        dtCombo.Columns.Add("price")

        Dim dtRowCombo As DataRow
        dtRowCombo = dtCombo.NewRow
        dtRowCombo("id") = "1"
        dtRowCombo("price") = "100"
        dtCombo.Rows.Add(dtRowCombo)

        dtRowCombo = dtCombo.NewRow
        dtRowCombo("id") = "2"
        dtRowCombo("price") = "120"
        dtCombo.Rows.Add(dtRowCombo)

        dtRowCombo = dtCombo.NewRow
        dtRowCombo("id") = "3"
        dtRowCombo("price") = "150"
        dtCombo.Rows.Add(dtRowCombo)

        ComboBox1.DataSource = dtCombo
        ComboBox1.DisplayMember = "price"
        ComboBox1.ValueMember = "id"

        '販売元コンボ
        Dim dtCombo2 As New DataTable
        dtCombo2.Columns.Add("id")
        dtCombo2.Columns.Add("dealer")

        Dim dtRowCombo2 As DataRow
        dtRowCombo2 = dtCombo2.NewRow
        dtRowCombo2("id") = "1"
        dtRowCombo2("dealer") = "サントリー"
        dtCombo2.Rows.Add(dtRowCombo2)

        dtRowCombo2 = dtCombo2.NewRow
        dtRowCombo2("id") = "2"
        dtRowCombo2("dealer") = "Asahi飲料"
        dtCombo2.Rows.Add(dtRowCombo2)

        dtRowCombo2 = dtCombo2.NewRow
        dtRowCombo2("id") = "3"
        dtRowCombo2("dealer") = "KIRIN"
        dtCombo2.Rows.Add(dtRowCombo2)

        ComboBox2.DataSource = dtCombo2
        ComboBox2.DisplayMember = "dealer"
        ComboBox2.ValueMember = "id"

        '温冷コンボ
        Dim dtCombo3 As New DataTable
        dtCombo3.Columns.Add("id")
        dtCombo3.Columns.Add("hot_or_cold")

        Dim dtRowCombo3 As DataRow
        dtRowCombo3 = dtCombo3.NewRow
        dtRowCombo3("id") = "1"
        dtRowCombo3("hot_or_cold") = "cold"
        dtCombo3.Rows.Add(dtRowCombo3)

        dtRowCombo3 = dtCombo3.NewRow
        dtRowCombo3("id") = "2"
        dtRowCombo3("hot_or_cold") = "hot"
        dtCombo3.Rows.Add(dtRowCombo3)

        ComboBox3.DataSource = dtCombo3
        ComboBox3.DisplayMember = "hot_or_cold"
        ComboBox3.ValueMember = "id"

        '初期表示
        Label3.Text = ItemId & "の更新"
        TextBox1.Text = row("item_name")
        ComboBox1.Text = row("price")
        ComboBox2.Text = row("dealer_name")
        'ComboBox3.Text = row("is_hot_or_cold")

        If row("is_hot_or_cold") = 1 Then
            ComboBox3.Text = "cold"
        Else
            ComboBox3.Text = "hot"
        End If

        conn.Close()
    End Sub

    '更新
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'データベース接続
        Dim conn As NpgsqlConnection = New NpgsqlConnection(getdatabaseconfig())
        conn.Open()

        Dim sql As String

        Dim newItem As String = TextBox1.Text
        Dim newPrice As Integer = ComboBox1.SelectedItem("id")
        Dim newDealer As Integer = ComboBox2.SelectedItem("id")
        Dim newHotcold As Integer = ComboBox3.SelectedItem("id")

        Dim cmd As New NpgsqlCommand

        sql = ""
        sql &= " update mst_items"
        sql &= " set item_name =" & "'" & newItem & "'"
        sql &= ", mst_price_id =" & newPrice
        sql &= ", mst_dealer_id =" & newDealer
        sql &= ", is_hot_or_cold =" & newHotcold
        sql &= " where id =" & ItemId

        ' SQL実行
        cmd.Connection = conn
        cmd.CommandText = sql
        cmd.ExecuteNonQuery()

        conn.Close()
    End Sub

    '次へ前へ関数
    Private Sub Change(cId)
        'データベース接続
        Dim conn As NpgsqlConnection = New NpgsqlConnection(getdatabaseconfig())
        conn.Open()

        Dim sql As String
        sql = ""
        sql &= " select"
        sql &= " a.id, a.item_name, b.price, c.dealer_name, a.is_hot_or_cold"
        sql &= " from mst_items as a"
        sql &= " left join mst_prices as b"
        sql &= " on a.mst_price_id = b.id"
        sql &= " left join mst_dealers as c"
        sql &= " on a.mst_dealer_id = c.id"
        sql &= " order by a.id asc"

        ' SQL実行
        Dim command As NpgsqlCommand = New NpgsqlCommand(sql, conn)
        Dim adapter As NpgsqlDataAdapter = New NpgsqlDataAdapter(command)
        Dim dt As DataTable = New DataTable()
        adapter.Fill(dt)

        Dim row As DataRow = dt.Rows(ItemId)

        ItemId = row("id")

        '表示
        Label3.Text = ItemId & "の更新"
        TextBox1.Text = row("item_name")
        ComboBox1.Text = row("price")
        ComboBox2.Text = row("dealer_name")
        'ComboBox3.Text = row("is_hot_or_cold")

        If row("is_hot_or_cold") = 1 Then
            ComboBox3.Text = "cold"
        Else
            ComboBox3.Text = "hot"
        End If

    End Sub

    '次へ
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'データベース接続
        Dim conn As NpgsqlConnection = New NpgsqlConnection(getdatabaseconfig())
        conn.Open()

        Dim sql As String
        sql = ""
        sql &= " select"
        sql &= " max(id)"
        sql &= " from mst_items"

        ' SQL実行
        Dim command As NpgsqlCommand = New NpgsqlCommand(sql, conn)
        Dim adapter As NpgsqlDataAdapter = New NpgsqlDataAdapter(command)
        Dim dt As DataTable = New DataTable()
        adapter.Fill(dt)

        Dim row As DataRow = dt.Rows(0)

        If ItemId < row("max") Then
            Change(ItemId)
        End If

    End Sub

    '前へ
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ItemId > 1 Then
            ItemId -= 2
            Change(ItemId)
        End If

    End Sub

    '追加
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'データベース接続
        Dim conn As NpgsqlConnection = New NpgsqlConnection(getdatabaseconfig())
        conn.Open()

        Dim sql As String

        Dim newItem As String = TextBox1.Text
        Dim newPrice As Integer = ComboBox1.SelectedItem("id")
        Dim newDealer As Integer = ComboBox2.SelectedItem("id")
        Dim newHotcold As Integer = ComboBox3.SelectedItem("id")

        Dim cmd As New NpgsqlCommand

        sql = ""
        sql &= " insert"
        sql &= " into mst_items"
        sql &= " (id, item_name, mst_price_id, mst_dealer_id, is_hot_or_cold)"
        sql &= " values"
        sql &= " (" & AddId
        sql &= " ,'" & newItem
        sql &= "', " & newPrice
        sql &= ", " & newDealer
        sql &= ", " & newHotcold
        sql &= ")"

        ' SQL実行
        cmd.Connection = conn
        cmd.CommandText = sql
        cmd.ExecuteNonQuery()

        sql = ""
        sql &= " insert"
        sql &= " into trn_stocks"
        sql &= " (id, mst_item_id, stock_count)"
        sql &= " values"
        sql &= " (" & AddId
        sql &= " ," & AddId
        sql &= " ,10)"

        ' SQL実行
        cmd.Connection = conn
        cmd.CommandText = sql
        cmd.ExecuteNonQuery()

        AddId += 1
        Label7.Text = AddId

        conn.Close()
    End Sub
End Class