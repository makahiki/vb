Imports Npgsql

Public Class Form1

    Private ItemId1 As Integer
    Private ItemId2 As Integer
    Private ItemId3 As Integer
    Private ItemId4 As Integer
    Private ItemId5 As Integer
    Private ItemId6 As Integer
    Private ItemId7 As Integer
    Private ItemId8 As Integer
    Private ItemId9 As Integer

    Private ItemPrice1 As Integer
    Private ItemPrice2 As Integer
    Private ItemPrice3 As Integer
    Private ItemPrice4 As Integer
    Private ItemPrice5 As Integer
    Private ItemPrice6 As Integer
    Private ItemPrice7 As Integer
    Private ItemPrice8 As Integer
    Private ItemPrice9 As Integer

    Private ChangeId1 As Integer
    Private ChangeId2 As Integer
    Private ChangeId3 As Integer
    Private ChangeId4 As Integer

    Private ChangeName1 As Integer
    Private ChangeName2 As Integer
    Private ChangeName3 As Integer
    Private ChangeName4 As Integer

    Private ChangeStock1 As Integer
    Private ChangeStock2 As Integer
    Private ChangeStock3 As Integer
    Private ChangeStock4 As Integer

    Private Coin1 As Integer
    Private Coin2 As Integer
    Private Coin3 As Integer
    Private Coin4 As Integer
    Private Coin5 As Integer

    Private Change As Integer

    Private htInput As New Hashtable

    ' 関数
    Private Sub BuyItem(ByVal Id As Integer, ByVal Price As Integer)


        Dim sql As String
        sql = ""
        sql &= " SELECT a.item_name, b.price, d.stock_count "
        sql &= " FROM mst_items as a"
        sql &= " left join mst_prices as b"
        sql &= " on a.mst_price_id = b.id"
        sql &= " left join mst_dealers as c"
        sql &= " on a.mst_dealer_id = c.id"
        sql &= " left join trn_stocks as d"
        sql &= " on mst_item_id = a.id"
        sql &= " where a.id =" & Id.ToString

        Dim dt As DataTable = New DataTable()

        Using db As New dbConnection()

            dt = db.getDtSql(sql)

        End Using

        ' ボタン用変数
        Dim row As DataRow

        ' ボタン
        row = dt.Rows(0)

        ' 購入
        If Label11.Text >= row("price") And row("stock_count") >= 1 Then
            Label2.Text = Label11.Text - row("price")
            Change = Label2.Text
        ElseIf Label11.Text >= row("price") And row("stock_count") <= 0 Then
            MessageBox.Show("売り切れ", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Else
            MessageBox.Show("足りません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim cmd As New NpgsqlCommand

        'おつり
        'For Each ht As DictionaryEntry In htInput
        '    sql = ""
        '    sql &= " update coin_bill_stocks"
        '    sql &= " set coin_bill_stock_count = coin_bill_stock_count + " & ht.Value
        '    sql &= " where coin_bill_name =" & ht.Key


        '    Using db As New dbConnection()

        '        Try
        '            db.trnStart()
        '            db.executeSql(sql)
        '            db.commit()
        '        Catch ex As Exception
        '            db.rollback()
        '            Throw
        '        End Try

        '    End Using



        'Next

        ' 在庫減らす
        sql = ""
        sql &= "update trn_stocks set stock_count=stock_count-1 where mst_item_id= " & Id.ToString

        Using db As New dbConnection()

            Try
                db.trnStart()
                db.executeSql(sql)
                db.commit()
            Catch ex As Exception
                db.rollback()
                Throw
            End Try

        End Using

        ' 実績登録
        sql = ""
        sql &= "insert into trn_claims "
        sql &= "(mst_item_id, buy_count, buy_price, buy_date) "
        sql &= "values(" & Id & ",'1'," & Price & ",now()) "

        Using db As New dbConnection()

            Try
                db.trnStart()
                db.executeSql(sql)
                db.commit()
            Catch ex As Exception
                db.rollback()
                Throw
            End Try

        End Using

        'htInputCreate()

        If row("stock_count") <= 10 Then
            MessageBox.Show("補充してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        Label11.Text = "0"

    End Sub


    Private Sub ChangeItem()


        Dim sql As String
        sql = ""
        sql &= " select id, coin_bill_name, coin_bill_stock_count"
        sql &= " from coin_bill_stocks"
        sql &= " order by id asc"

        Dim dt As DataTable = New DataTable()

        Using db As New dbConnection()

            dt = db.getDtSql(sql)

        End Using

        Dim row As DataRow

        '10円
        row = dt.Rows(0)
        ChangeId1 = row("id")
        ChangeName1 = row("coin_bill_name")
        ChangeStock1 = row("coin_bill_stock_count")

        '50円
        row = dt.Rows(1)
        ChangeId2 = row("id")
        ChangeName2 = row("coin_bill_name")
        ChangeStock2 = row("coin_bill_stock_count")

        '100円
        row = dt.Rows(2)
        ChangeName3 = row("coin_bill_name")
        ChangeId3 = row("id")
        ChangeStock3 = row("coin_bill_stock_count")

        '500円
        row = dt.Rows(3)
        ChangeName4 = row("coin_bill_name")
        ChangeId4 = row("id")
        ChangeStock4 = row("coin_bill_stock_count")

        While Change > 0
            If Change >= 500 And ChangeStock4 > 0 Then
                Change = Change - ChangeName4
                ChangeStock4 = ChangeStock4 - 1
            ElseIf Change >= 100 And ChangeStock3 > 0 Then
                Change = Change - ChangeName3
                ChangeStock3 = ChangeStock3 - 1
            ElseIf Change >= 50 And ChangeStock2 > 0 Then
                Change = Change - ChangeName2
                ChangeStock2 = ChangeStock2 - 1
            ElseIf Change >= 10 And ChangeStock1 > 0 Then
                Change = Change - ChangeName1
                ChangeStock1 = ChangeStock1 - 1
            Else
                MessageBox.Show("おつりなし", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit While
            End If

        End While

        Dim cmd As New NpgsqlCommand

        sql = UpdateChange(ChangeId1, ChangeStock1)
        Using db As New dbConnection()

            Try
                db.trnStart()
                db.executeSql(sql)
                db.commit()
            Catch ex As Exception
                db.rollback()
                Throw
            End Try

        End Using

        sql = UpdateChange(ChangeId2, ChangeStock2)
        Using db As New dbConnection()

            Try
                db.trnStart()
                db.executeSql(sql)
                db.commit()
            Catch ex As Exception
                db.rollback()
                Throw
            End Try

        End Using

        sql = UpdateChange(ChangeId3, ChangeStock3)
        Using db As New dbConnection()

            Try
                db.trnStart()
                db.executeSql(sql)
                db.commit()
            Catch ex As Exception
                db.rollback()
                Throw
            End Try

        End Using

        sql = UpdateChange(ChangeId4, ChangeStock4)
        Using db As New dbConnection()

            Try
                db.trnStart()
                db.executeSql(sql)
                db.commit()
            Catch ex As Exception
                db.rollback()
                Throw
            End Try

        End Using

    End Sub

    Private Function UpdateChange(ByVal cId As Integer, ByVal cStock As Integer)
        Dim sql As String

        sql = ""
        sql &= " update coin_bill_stocks"
        sql &= " set coin_bill_stock_count"
        sql &= " ="
        sql &= " " & cStock
        sql &= " where id=" & cId

        Return sql
    End Function


    Private Sub htInputCreate()
        htInput.Clear()
        htInput.Add("10円", 0)
        htInput.Add("50円", 0)
        htInput.Add("100円", 0)
        htInput.Add("500円", 0)
        htInput.Add("1000円", 0)
    End Sub

    '10円
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label11.Text = Label11.Text + 10
        htInput(10) = htInput(10) + 1
    End Sub

    '50円
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label11.Text = Label11.Text + 50
        htInput(50) = htInput(50) + 1
    End Sub

    '100円
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label11.Text = Label11.Text + 100
        htInput(100) = htInput(100) + 1
    End Sub

    '500円
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Label11.Text = Label11.Text + 500
        htInput(500) = htInput(500) + 1
    End Sub

    '1000円
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Label11.Text = Label11.Text + 1000
        htInput(1000) = htInput(1000) + 1
    End Sub

    'ボタン表示
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        'データベース接続


        Dim sql As String
        sql = ""
        sql &= " select a.item_name, b.price, a.id, a.is_hot_or_cold"
        sql &= " from mst_items as a"
        sql &= " left join mst_prices as b"
        sql &= " on a.mst_price_id = b.id"
        sql &= " left join mst_dealers as c"
        sql &= " on a.mst_dealer_id = c.id"
        sql &= " order by random()"

        Dim dt As DataTable = New DataTable()

        Using db As New dbConnection()

            dt = db.getDtSql(sql)

        End Using

        ' ボタン表示用変数
        Dim row As DataRow
        Dim en As String = "円"

        ' ボタン1
        row = dt.Rows(0)
        ItemId1 = row("id")
        ItemPrice1 = row("price")
        Button14.Text = row("item_name") & vbCrLf & Math.Floor(row("price")) & en
        If row("is_hot_or_cold") = 1 Then
            Button14.BackColor = Color.Blue
        ElseIf row("is_hot_or_cold") = 2 Then
            Button14.BackColor = Color.Red
        End If

        ' ボタン2
        row = dt.Rows(1)
        ItemId2 = row("id")
        ItemPrice2 = row("price")
        Button6.Text = row("item_name") & vbCrLf & Math.Floor(row("price")) & en
        If row("is_hot_or_cold") = 1 Then
            Button6.BackColor = Color.Blue
        ElseIf row("is_hot_or_cold") = 2 Then
            Button6.BackColor = Color.Red
        End If

        ' ボタン3
        row = dt.Rows(2)
        ItemId3 = row("id")
        ItemPrice3 = row("price")
        Button7.Text = row("item_name") & vbCrLf & Math.Floor(row("price")) & en
        If row("is_hot_or_cold") = 1 Then
            Button7.BackColor = Color.Blue
        ElseIf row("is_hot_or_cold") = 2 Then
            Button7.BackColor = Color.Red
        End If

        ' ボタン4
        row = dt.Rows(3)
        ItemId4 = row("id")
        ItemPrice4 = row("price")
        Button8.Text = row("item_name") & vbCrLf & Math.Floor(row("price")) & en
        If row("is_hot_or_cold") = 1 Then
            Button8.BackColor = Color.Blue
        ElseIf row("is_hot_or_cold") = 2 Then
            Button8.BackColor = Color.Red
        End If

        ' ボタン5
        row = dt.Rows(4)
        ItemId5 = row("id")
        ItemPrice5 = row("price")
        Button9.Text = row("item_name") & vbCrLf & Math.Floor(row("price")) & en
        If row("is_hot_or_cold") = 1 Then
            Button9.BackColor = Color.Blue
        ElseIf row("is_hot_or_cold") = 2 Then
            Button9.BackColor = Color.Red
        End If

        ' ボタン6
        row = dt.Rows(5)
        ItemId6 = row("id")
        ItemPrice6 = row("price")
        Button12.Text = row("item_name") & vbCrLf & Math.Floor(row("price")) & en
        If row("is_hot_or_cold") = 1 Then
            Button12.BackColor = Color.Blue
        ElseIf row("is_hot_or_cold") = 2 Then
            Button12.BackColor = Color.Red
        End If

        ' ボタン7
        row = dt.Rows(6)
        ItemId7 = row("id")
        ItemPrice7 = row("price")
        Button11.Text = row("item_name") & vbCrLf & Math.Floor(row("price")) & en
        If row("is_hot_or_cold") = 1 Then
            Button11.BackColor = Color.Blue
        ElseIf row("is_hot_or_cold") = 2 Then
            Button11.BackColor = Color.Red
        End If

        ' ボタン8
        row = dt.Rows(7)
        ItemId8 = row("id")
        ItemPrice8 = row("price")
        Button10.Text = row("item_name") & vbCrLf & Math.Floor(row("price")) & en
        If row("is_hot_or_cold") = 1 Then
            Button10.BackColor = Color.Blue
        ElseIf row("is_hot_or_cold") = 2 Then
            Button10.BackColor = Color.Red
        End If

        ' ボタン9
        row = dt.Rows(8)
        ItemId9 = row("id")
        ItemPrice9 = row("price")
        Button15.Text = row("item_name") & vbCrLf & Math.Floor(row("price")) & en
        If row("is_hot_or_cold") = 1 Then
            Button15.BackColor = Color.Blue
        ElseIf row("is_hot_or_cold") = 2 Then
            Button15.BackColor = Color.Red
        End If

    End Sub

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

    'クリア
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        htInputCreate()
        Label2.Text = "0"
        Label11.Text = "0"

    End Sub

    'ボタン1
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        BuyItem(ItemId1, ItemPrice1)
        ChangeItem()
    End Sub

    'ボタン2
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        BuyItem(ItemId2, ItemPrice2)
        ChangeItem()
    End Sub

    'ボタン3
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        BuyItem(ItemId3, ItemPrice3)
        ChangeItem()
    End Sub

    'ボタン4
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        BuyItem(ItemId4, ItemPrice4)
        ChangeItem()
    End Sub

    'ボタン5
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        BuyItem(ItemId5, ItemPrice5)
        ChangeItem()
    End Sub

    'ボタン6
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        BuyItem(ItemId6, ItemPrice6)
        ChangeItem()
    End Sub

    'ボタン7
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        BuyItem(ItemId7, ItemPrice7)
        ChangeItem()
    End Sub

    'ボタン8
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        BuyItem(ItemId8, ItemPrice8)
        ChangeItem()
    End Sub

    'ボタン9
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        BuyItem(ItemId9, ItemPrice9)
        ChangeItem()
    End Sub

    '購入一覧
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Dim f2 As New Form2
        f2.Show()
    End Sub

    '硬貨在庫
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim f3 As New Form3
        f3.Show()
    End Sub

    '商品在庫
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim f4 As New Form4
        f4.Show()
    End Sub

    '商品編集
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim f5 As New Form5
        f5.Show()
    End Sub
End Class
