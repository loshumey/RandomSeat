Public Class Form1
    Dim btn(40) As Integer
    Dim oldkey As Integer
    Dim anskey As String   '---存放目前的標準答案---
    Dim start_key As Integer
    Dim end_key As Integer

    Dim tot_time As Integer

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        For i = 1 To 40
            Controls("btn_" & i).Visible = True
        Next
        For i = 1 To 40
            Controls("Btn_" & i).BackColor = Color.LightGray
            Controls("Btn_" & i).ForeColor = Color.Black
        Next
        Timer1.Interval = 100
        lblans1.Visible = False
        txtinput1.Visible = False
        lblerr1.Visible = False

        lblans2.Visible = False
        txtinput2.Visible = False
        lblerr2.Visible = False

    End Sub



    Private Sub 結束ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 結束ToolStripMenuItem.Click
        End
    End Sub

    '---------Timer1負責亂出下一個按鍵-------------
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Randomize()
        Dim newkey As Integer

        '------將舊字外觀復原------
        Controls("Btn_" & oldkey).BackColor = Color.LightGray
        Controls("Btn_" & oldkey).ForeColor = Color.Black
        Controls("Btn_" & oldkey).Font = New Font("微軟正黑體", 24, FontStyle.Regular)

        '-----亂出下一個字-----
        Do
            newkey = Int(Rnd() * (end_key - start_key + 1)) + start_key
        Loop Until newkey <> oldkey
        lbl_tot.Text = Val(lbl_tot.Text) + 1
        Controls("Btn_" & newkey).BackColor = Color.Blue
        Controls("Btn_" & newkey).ForeColor = Color.Red
        Controls("Btn_" & newkey).Font = New Font("微軟正黑體", 28, FontStyle.Bold)
        anskey = Controls("Btn_" & newkey).Text

        oldkey = newkey

        Timer1.Enabled = False '----先close時鐘-----

    End Sub

    '--------按鍵檢查----------
    Private Sub Btn_1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles _
        Btn_1.KeyPress, Btn_2.KeyPress, Btn_3.KeyPress, Btn_4.KeyPress, Btn_5.KeyPress, Btn_6.KeyPress, Btn_7.KeyPress, Btn_8.KeyPress, Btn_9.KeyPress, Btn_10.KeyPress, _
        Btn_11.KeyPress, Btn_12.KeyPress, Btn_13.KeyPress, Btn_14.KeyPress, Btn_15.KeyPress, Btn_16.KeyPress, Btn_17.KeyPress, Btn_18.KeyPress, Btn_19.KeyPress, Btn_20.KeyPress, _
        Btn_21.KeyPress, Btn_22.KeyPress, Btn_23.KeyPress, Btn_24.KeyPress, Btn_25.KeyPress, Btn_26.KeyPress, Btn_27.KeyPress, Btn_28.KeyPress, Btn_29.KeyPress, Btn_30.KeyPress, _
        Btn_31.KeyPress, Btn_32.KeyPress, Btn_33.KeyPress, Btn_34.KeyPress, Btn_35.KeyPress, Btn_36.KeyPress, Btn_37.KeyPress, Btn_38.KeyPress, Btn_39.KeyPress, Btn_40.KeyPress

        Dim scr As Single = 0

        '----空白鍵結束---------
        If e.KeyChar = Chr(32) Then
            Timer1.Enabled = False
            Timer2.Enabled = False
            scr = Fix(Val(lbl_tot.Text) / tot_time * 60)
            MsgBox("練習結束！" & vbNewLine & vbNewLine & "速度：每分鐘" & scr & "字")
        Else
            '----判斷有沒有按對按鍵---------
            If anskey = UCase(e.KeyChar) Then
                Timer1.Enabled = True
            End If
        End If

    End Sub

    '--------單鍵練習_純字母----------
    Private Sub 字母ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 字母ToolStripMenuItem.Click
        For i = 1 To 40
            Controls("btn_" & i).Visible = True
        Next
        For i = 1 To 40
            Controls("Btn_" & i).BackColor = Color.LightGray
            Controls("Btn_" & i).ForeColor = Color.Black
        Next
        lblans1.Visible = False
        txtinput1.Visible = False
        lblerr1.Visible = False
        lblans2.Visible = False
        txtinput2.Visible = False
        lblerr2.Visible = False
        Btn_1.Focus()
        oldkey = 1
        start_key = 11
        end_key = 40

        lbl_tot.Text = 0
        lbl_time.Text = 0

        中間單排ToolStripMenuItem.Checked = False
        中間下面ToolStripMenuItem.Checked = False
        中間上面ToolStripMenuItem.Checked = False
        數字字母ToolStripMenuItem.Checked = False
        字母ToolStripMenuItem.Checked = True
        status_key.Text = "英文字母"


        Dim anykey = MsgBox("按ENTER鍵開始…" & vbNewLine & "按空白鍵結束")
        Timer1.Enabled = True
        tot_time = 0    '-----計算經過時間------
        Timer2.Enabled = True
    End Sub

    '--------單鍵練習_數字+字母----------
    Private Sub 數字字母ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 數字字母ToolStripMenuItem.Click
        For i = 1 To 40
            Controls("btn_" & i).Visible = True
        Next
        For i = 1 To 40
            Controls("Btn_" & i).BackColor = Color.LightGray
            Controls("Btn_" & i).ForeColor = Color.Black
        Next
        lblans1.Visible = False
        txtinput1.Visible = False
        lblerr1.Visible = False
        lblans2.Visible = False
        txtinput2.Visible = False
        lblerr2.Visible = False
        Btn_1.Focus()
        oldkey = 1
        start_key = 1
        end_key = 40

        lbl_tot.Text = 0
        lbl_time.Text = 0

        中間單排ToolStripMenuItem.Checked = False
        中間下面ToolStripMenuItem.Checked = False
        中間上面ToolStripMenuItem.Checked = False
        數字字母ToolStripMenuItem.Checked = True
        字母ToolStripMenuItem.Checked = False
        status_key.Text = "數字+字母"


        Dim anykey = MsgBox("按Enter鍵開始…" & vbNewLine & "按空白鍵結束")
        Timer1.Enabled = True
        tot_time = 0    '-----計算經過時間------
        Timer2.Enabled = True
    End Sub

    '--------單鍵練習_中間+上面----------
    Private Sub 中間上面ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 中間上面ToolStripMenuItem.Click
        For i = 1 To 40
            Controls("btn_" & i).Visible = True
        Next
        For i = 1 To 40
            Controls("Btn_" & i).BackColor = Color.LightGray
            Controls("Btn_" & i).ForeColor = Color.Black
        Next
        lblans1.Visible = False
        txtinput1.Visible = False
        lblerr1.Visible = False
        lblans2.Visible = False
        txtinput2.Visible = False
        lblerr2.Visible = False
        Btn_1.Focus()
        oldkey = 1
        start_key = 11
        end_key = 30

        lbl_tot.Text = 0
        lbl_time.Text = 0

        中間單排ToolStripMenuItem.Checked = False
        中間下面ToolStripMenuItem.Checked = False
        中間上面ToolStripMenuItem.Checked = True
        數字字母ToolStripMenuItem.Checked = False
        字母ToolStripMenuItem.Checked = False
        status_key.Text = "中間+上排"

        Dim anykey = MsgBox("按Enter鍵開始…" & vbNewLine & "按空白鍵結束")
        Timer1.Enabled = True
        tot_time = 0    '-----計算經過時間------
        Timer2.Enabled = True
    End Sub

    '--------單鍵練習_中間+下面----------
    Private Sub 中間下面ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 中間下面ToolStripMenuItem.Click
        For i = 1 To 40
            Controls("btn_" & i).Visible = True
        Next
        For i = 1 To 40
            Controls("Btn_" & i).BackColor = Color.LightGray
            Controls("Btn_" & i).ForeColor = Color.Black
        Next
        lblans1.Visible = False
        txtinput1.Visible = False
        lblerr1.Visible = False
        lblans2.Visible = False
        txtinput2.Visible = False
        lblerr2.Visible = False
        Btn_1.Focus()
        oldkey = 1
        start_key = 21
        end_key = 40

        lbl_tot.Text = 0
        lbl_time.Text = 0

        中間單排ToolStripMenuItem.Checked = False
        中間下面ToolStripMenuItem.Checked = True
        中間上面ToolStripMenuItem.Checked = False
        數字字母ToolStripMenuItem.Checked = False
        字母ToolStripMenuItem.Checked = False
        status_key.Text = "中間+下排"

        Dim anykey = MsgBox("按Enter鍵開始…" & vbNewLine & "按空白鍵結束")
        Timer1.Enabled = True
        tot_time = 0    '-----計算經過時間------
        Timer2.Enabled = True
    End Sub

    '--------單鍵練習_中間----------
    Private Sub 中間單排ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 中間單排ToolStripMenuItem.Click
        For i = 1 To 40
            Controls("btn_" & i).Visible = True
        Next
        For i = 1 To 40
            Controls("Btn_" & i).BackColor = Color.LightGray
            Controls("Btn_" & i).ForeColor = Color.Black
        Next
        lblans1.Visible = False
        txtinput1.Visible = False
        lblerr1.Visible = False
        lblans2.Visible = False
        txtinput2.Visible = False
        lblerr2.Visible = False

        Btn_1.Focus()
        oldkey = 1
        start_key = 21
        end_key = 30
        lbl_tot.Text = 0
        lbl_time.Text = 0

        中間單排ToolStripMenuItem.Checked = True
        中間下面ToolStripMenuItem.Checked = False
        中間上面ToolStripMenuItem.Checked = False
        數字字母ToolStripMenuItem.Checked = False
        字母ToolStripMenuItem.Checked = False
        status_key.Text = "中間單排"

        Dim anykey = MsgBox("按Enter鍵開始…" & vbNewLine & "按空白鍵結束")
        Timer1.Enabled = True
        tot_time = 0    '-----計算經過時間------
        Timer2.Enabled = True
    End Sub

    '-------Timer2專門計算經過時間--------
    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Dim m, s As Integer
        tot_time = tot_time + 1
        m = tot_time \ 60
        s = tot_time Mod 60

        lbl_time.Text = m & ":" & s
    End Sub

    '------單行練習_中間單排-----------
    Private Sub 中間單排ToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles 中間單排ToolStripMenuItem1.Click
        Dim ch = "ASDFGHJKL"
        Dim n As Integer = 0
        Randomize()

        For i = 1 To 40
            Controls("btn_" & i).Visible = False
        Next
        lblans1.Visible = True
        lblerr1.Visible = True
        txtinput1.Visible = True
        lblans2.Visible = True
        lblerr2.Visible = True
        txtinput2.Visible = True
        lblerr1.Text = Nothing
        lblans1.Text = Nothing
        txtinput1.Text = Nothing
        lblerr2.Text = Nothing
        lblans2.Text = Nothing
        txtinput2.Text = Nothing
        txtinput1.Focus()

        lbl_time.Text = 0

        status_key.Text = "中間單排"
        Dim anykey = MsgBox("按ENTER鍵開始…" & vbNewLine & "按ENTER結束")

        '------產生參考文字標籤內容------------
        For j = 1 To 10
            For i = 1 To 5
                n = Int(Rnd() * 9)
                lblans1.Text = lblans1.Text & ch(n)
                n = Int(Rnd() * 9)
                lblans2.Text = lblans2.Text & ch(n)
            Next
            lblans1.Text = lblans1.Text & " "
            lblans2.Text = lblans2.Text & " "
        Next

        tot_time = 0    '-----計算經過時間------
        Timer2.Enabled = True
    End Sub

    '---------單排輸入結束--------
    Private Sub txtinput_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtinput1.KeyPress
        Dim scr As Single = 0

        e.KeyChar = UCase(e.KeyChar)
        oldkey = 0
        '----ENTER 結束---------
        If e.KeyChar = Chr(13) Then
            Timer2.Enabled = False
            scr = Fix(Len(txtinput1.Text) / tot_time * 60)
            MsgBox("練習結束！" & vbNewLine & vbNewLine & "速度：每分鐘" & scr & "字")
        End If

        '---倒退鍵---------
        If e.KeyChar = Chr(8) Then
            lblerr1.Text = Strings.Left(lblerr1.Text, Len(txtinput1.Text) - 1)
            oldkey = 8
        End If

    End Sub

    '---------輸入錯誤時呈現紅色文字提醒--------
    Private Sub txtinput_TextChanged(sender As Object, e As System.EventArgs) Handles txtinput1.TextChanged
        Dim ch1, ch2 As String
        If Len(txtinput1.Text) >= 1 And oldkey <> 8 Then
            ch1 = Strings.Right(txtinput1.Text, 1)
            ch2 = Mid(lblans1.Text, Len(txtinput1.Text), 1)
            If ch1 <> ch2 Then
                lblerr1.Text = lblerr1.Text & "X"
            Else
                lblerr1.Text = lblerr1.Text & " "
            End If
        End If

        If Len(txtinput1.Text) > 59 Then
            txtinput2.Focus()
        End If

        lbl_tot.Text = Len(txtinput1.Text) + Len(txtinput2.Text)

    End Sub

    '---------單排練習_中間上排--------
    Private Sub 中間上排ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 中間上排ToolStripMenuItem.Click
        Dim ch = "ASDFGHJKLQWERTYUIOP"
        Dim n As Integer = 0
        Randomize()

        For i = 1 To 40
            Controls("btn_" & i).Visible = False
        Next
        lblans1.Visible = True
        lblerr1.Visible = True
        txtinput1.Visible = True
        lblans2.Visible = True
        lblerr2.Visible = True
        txtinput2.Visible = True
        lblerr1.Text = Nothing
        lblans1.Text = Nothing
        txtinput1.Text = Nothing
        lblerr2.Text = Nothing
        lblans2.Text = Nothing
        txtinput2.Text = Nothing
        txtinput1.Focus()

        lbl_time.Text = 0

        status_key.Text = "中間+上排"
        Dim anykey = MsgBox("按ENTER鍵開始…" & vbNewLine & "按ENTER結束")

        '------產生參考文字標籤內容------------
        For j = 1 To 10
            For i = 1 To 5
                n = Int(Rnd() * 9)
                lblans1.Text = lblans1.Text & ch(n)
                n = Int(Rnd() * 9)
                lblans2.Text = lblans2.Text & ch(n)
            Next
            lblans1.Text = lblans1.Text & " "
            lblans2.Text = lblans2.Text & " "
        Next

        tot_time = 0    '-----計算經過時間------
        Timer2.Enabled = True
    End Sub

    '---------單排練習_中間下排--------
    Private Sub 中間下排ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 中間下排ToolStripMenuItem.Click
        Dim ch = "ASDFGHJKLZXCVBNM"
        Dim n As Integer = 0
        Randomize()

        For i = 1 To 40
            Controls("btn_" & i).Visible = False
        Next
        lblans1.Visible = True
        lblerr1.Visible = True
        txtinput1.Visible = True
        lblans2.Visible = True
        lblerr2.Visible = True
        txtinput2.Visible = True
        lblerr1.Text = Nothing
        lblans1.Text = Nothing
        txtinput1.Text = Nothing
        lblerr2.Text = Nothing
        lblans2.Text = Nothing
        txtinput2.Text = Nothing
        txtinput1.Focus()

        lbl_time.Text = 0

        status_key.Text = "中間+下排"
        Dim anykey = MsgBox("按ENTER鍵開始…" & vbNewLine & "按ENTER結束")

        '------產生參考文字標籤內容------------
        For j = 1 To 10
            For i = 1 To 5
                n = Int(Rnd() * 9)
                lblans1.Text = lblans1.Text & ch(n)
                n = Int(Rnd() * 9)
                lblans2.Text = lblans2.Text & ch(n)
            Next
            lblans1.Text = lblans1.Text & " "
            lblans2.Text = lblans2.Text & " "
        Next

        tot_time = 0    '-----計算經過時間------
        Timer2.Enabled = True
    End Sub

    '---------單排練習_全部字母--------
    Private Sub 全部字母ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 全部字母ToolStripMenuItem.Click
        Dim ch = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim n As Integer = 0
        Randomize()

        For i = 1 To 40
            Controls("btn_" & i).Visible = False
        Next
        lblans1.Visible = True
        lblerr1.Visible = True
        txtinput1.Visible = True
        lblans2.Visible = True
        lblerr2.Visible = True
        txtinput2.Visible = True
        lblerr1.Text = Nothing
        lblans1.Text = Nothing
        txtinput1.Text = Nothing
        lblerr2.Text = Nothing
        lblans2.Text = Nothing
        txtinput2.Text = Nothing
        txtinput1.Focus()

        lbl_time.Text = 0

        status_key.Text = "所有英文字母"
        Dim anykey = MsgBox("按ENTER鍵開始…" & vbNewLine & vbNewLine & "按ENTER結束")

        '------產生參考文字標籤內容------------
        For j = 1 To 10
            For i = 1 To 5
                n = Int(Rnd() * 9)
                lblans1.Text = lblans1.Text & ch(n)
                n = Int(Rnd() * 9)
                lblans2.Text = lblans2.Text & ch(n)
            Next
            lblans1.Text = lblans1.Text & " "
            lblans2.Text = lblans2.Text & " "
        Next

        tot_time = 0    '-----計算經過時間------
        Timer2.Enabled = True
    End Sub

    '---------單排練習_字母+數字--------
    Private Sub 字母數字ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 字母數字ToolStripMenuItem.Click
        Dim ch = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim n As Integer = 0
        Randomize()

        For i = 1 To 40
            Controls("btn_" & i).Visible = False
        Next
        lblans1.Visible = True
        lblerr1.Visible = True
        txtinput1.Visible = True
        lblans2.Visible = True
        lblerr2.Visible = True
        txtinput2.Visible = True
        lblerr1.Text = Nothing
        lblans1.Text = Nothing
        txtinput1.Text = Nothing
        lblerr2.Text = Nothing
        lblans2.Text = Nothing
        txtinput2.Text = Nothing
        txtinput1.Focus()

        lbl_time.Text = 0

        status_key.Text = "字母+數字"
        Dim anykey = MsgBox("按ENTER鍵開始…" & vbNewLine & "按ENTER結束")

        '------產生參考文字標籤內容------------
        For j = 1 To 10
            For i = 1 To 5
                n = Int(Rnd() * 9)
                lblans1.Text = lblans1.Text & ch(n)
                n = Int(Rnd() * 9)
                lblans2.Text = lblans2.Text & ch(n)
            Next
            lblans1.Text = lblans1.Text & " "
            lblans2.Text = lblans2.Text & " "
        Next

        tot_time = 0    '-----計算經過時間------
        Timer2.Enabled = True
    End Sub


    '---------第二個文字框檢查結束--------
    Private Sub txtinput2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtinput2.KeyPress
        Dim scr As Single = 0

        e.KeyChar = UCase(e.KeyChar)
        oldkey = 0
        '----ENTER 結束---------
        If e.KeyChar = Chr(13) Then
            Timer2.Enabled = False
            scr = Fix((Len(txtinput1.Text) + Len(txtinput2.Text)) / tot_time * 60)
            MsgBox("練習結束！" & vbNewLine & vbNewLine & "速度：每分鐘" & scr & "字")
        End If

        '---倒退鍵---------
        If e.KeyChar = Chr(8) Then
            lblerr2.Text = Strings.Left(lblerr2.Text, Len(txtinput2.Text) - 1)
            oldkey = 8
        End If
    End Sub

    '---------第二個文字框檢查是否輸錯--------
    Private Sub txtinput2_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtinput2.TextChanged
        Dim ch1, ch2 As String
        If Len(txtinput2.Text) >= 1 And oldkey <> 8 Then
            ch1 = Strings.Right(txtinput2.Text, 1)
            ch2 = Mid(lblans2.Text, Len(txtinput2.Text), 1)
            If ch1 <> ch2 Then
                lblerr2.Text = lblerr2.Text & "X"
            Else
                lblerr2.Text = lblerr2.Text & " "
            End If
        End If

        lbl_tot.Text = Len(txtinput1.Text) + Len(txtinput2.Text)
    End Sub
End Class
