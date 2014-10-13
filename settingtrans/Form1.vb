Public Class Form1

    Dim HighSlot(4), MidSlot(4), LowSlot(4) As String
    Dim Fixes_Drones As String
    Const FaD As String = "==改装件和无人机=="
    Const Fit As String = "<div style=""clear:both"">" + vbCrLf + "==舰船装配==" + vbCrLf + "</div>"
    Const TableHead As String = "{| class=""article-table"" border=1 style=""width:100% padding:0px; margin:0px;""" + _
        vbCrLf + "|-" + vbCrLf + _
        "! scope=""col"" | 槽位" + vbCrLf + _
        "! scope=""col"" | 10人副本" + vbCrLf + _
        "! scope=""col"" | 20人副本" + vbCrLf + _
        "! scope=""col"" | 40人副本" + vbCrLf + _
        "! scope=""col"" | 决战副本" + vbCrLf + _
        "|-"
    Dim tmpl As String = ""

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        processAllString()
    End Sub

    Private Function processAllString()
        ReDim HighSlot(4), MidSlot(4), LowSlot(4)
        Fixes_Drones = ""

        tmpl = "{{舰船属性|舰船名称=" + TextBox1.Text + "|图片名称=" + TextBox3.Text + _
            "|槽位数量=" + NumericUpDown1.Value.ToString + "高/" + NumericUpDown2.Value.ToString + "中/" + NumericUpDown3.Value.ToString + "低|舰船类型=" + _
            TextBox2.Text + _
            "|护盾=" + NumericUpDown4.Value.ToString + "|装甲=" + NumericUpDown5.Value.ToString + "|结构=" + NumericUpDown6.Value.ToString + _
            "}}" + vbCrLf + vbCrLf
        RichTextBox6.Text += tmpl

        Dim s As String() = RichTextBox1.Lines
        Dim currcat As String = ""
        For i As Integer = 0 To UBound(s)
            Select Case s(i)
                Case "改装件插槽"
                    Fixes_Drones = Fixes_Drones + "===改装件===" + vbCrLf
                Case "无人机"
                    Fixes_Drones = Fixes_Drones + "===无人机===" + vbCrLf
                Case "子系统"
                    Fixes_Drones = Fixes_Drones + "===子系统===" + vbCrLf
                Case ""
                    Continue For
                Case Else
                    Fixes_Drones += "* " + s(i) + vbCrLf
            End Select
        Next
        RichTextBox6.Text += "==改装件和无人机==" + vbCrLf
        RichTextBox6.Text += Fixes_Drones + vbCrLf + vbCrLf
        RichTextBox6.Text += Fit + vbCrLf + TableHead + vbCrLf
        ProcessFits()
    End Function

    Private Function ProcessFits()
        Dim k1 As String() = RichTextBox2.Lines
        Dim currcat As String = ""
        For i As Integer = 0 To UBound(k1)
            Select Case k1(i)
                Case "改装件插槽"
                    Exit For
                Case "无人机"
                    Exit For
                Case "子系统"
                    Exit For
                Case "高能量"
                    HighSlot(0) += "| style=""text-align:center; width:10%;"" | 高槽位" + vbCrLf + "| "
                    currcat = "h"
                Case "中级能量"
                    MidSlot(0) += "|-" + vbCrLf + "| style=""text-align:center; width:10%;"" | 中槽位" + vbCrLf + "| "
                    currcat = "m"
                Case "低能量"
                    LowSlot(0) += "|-" + vbCrLf + "| style=""text-align:center; width:10%;"" | 低槽位" + vbCrLf + "| "
                    currcat = "l"
                Case ""
                    Continue For
                Case Else
                    Select Case currcat
                        Case "h"
                            HighSlot(0) += k1(i) + "<br />"
                        Case "m"
                            MidSlot(0) += k1(i) + "<br />"
                        Case "l"
                            LowSlot(0) += k1(i) + "<br />"
                    End Select
            End Select
        Next
        Dim k2 As String() = RichTextBox3.Lines
        currcat = ""
        For i As Integer = 0 To UBound(k2)
            Select Case k2(i)
                Case "改装件插槽"
                    Exit For
                Case "无人机"
                    Exit For
                Case "高能量"
                    HighSlot(1) += vbCrLf + "| "
                    currcat = "h"
                Case "中级能量"
                    MidSlot(1) += vbCrLf + "| "
                    currcat = "m"
                Case "低能量"
                    LowSlot(1) += vbCrLf + "| "
                    currcat = "l"
                Case ""
                    Continue For
                Case Else
                    Select Case currcat
                        Case "h"
                            HighSlot(1) += k2(i) + "<br />"
                        Case "m"
                            MidSlot(1) += k2(i) + "<br />"
                        Case "l"
                            LowSlot(1) += k2(i) + "<br />"
                    End Select
            End Select
        Next
        Dim k3 As String() = RichTextBox4.Lines
        currcat = ""
        For i As Integer = 0 To UBound(k3)
            Select Case k3(i)
                Case "改装件插槽"
                    Exit For
                Case "无人机"
                    Exit For
                Case "高能量"
                    HighSlot(2) += vbCrLf + "| "
                    currcat = "h"
                Case "中级能量"
                    MidSlot(2) += vbCrLf + "| "
                    currcat = "m"
                Case "低能量"
                    LowSlot(2) += vbCrLf + "| "
                    currcat = "l"
                Case ""
                    Continue For
                Case Else
                    Select Case currcat
                        Case "h"
                            HighSlot(2) += k3(i) + "<br />"
                        Case "m"
                            MidSlot(2) += k3(i) + "<br />"
                        Case "l"
                            LowSlot(2) += k3(i) + "<br />"
                    End Select
            End Select
        Next
        Dim k4 As String() = RichTextBox5.Lines
        currcat = ""
        For i As Integer = 0 To UBound(k4)
            Select Case k4(i)
                Case "改装件插槽"
                    Exit For
                Case "无人机"
                    Exit For
                Case "高能量"
                    HighSlot(3) += vbCrLf + "| "
                    currcat = "h"
                Case "中级能量"
                    MidSlot(3) += vbCrLf + "| "
                    currcat = "m"
                Case "低能量"
                    LowSlot(3) += vbCrLf + "| "
                    currcat = "l"
                Case ""
                    Continue For
                Case Else
                    Select Case currcat
                        Case "h"
                            HighSlot(3) += k4(i) + "<br />"
                        Case "m"
                            MidSlot(3) += k4(i) + "<br />"
                        Case "l"
                            LowSlot(3) += k4(i) + "<br />"
                    End Select
            End Select
        Next
        Dim TableResult(3) As String
        For i As Integer = 0 To 3
            TableResult(0) += HighSlot(i).Remove(HighSlot(i).LastIndexOf("<"))
            TableResult(1) += MidSlot(i).Remove(MidSlot(i).LastIndexOf("<"))
            TableResult(2) += LowSlot(i).Remove(LowSlot(i).LastIndexOf("<"))
        Next
        Dim result = TableResult(0) + vbCrLf + TableResult(1) + vbCrLf + TableResult(2) + vbCrLf + "|}"
        RichTextBox6.Text += result + vbCrLf + vbCrLf + "{{导航}}"
    End Function

    Private Sub NumericUpDown1_Enter(sender As Object, e As EventArgs) Handles NumericUpDown6.Enter, NumericUpDown5.Enter, NumericUpDown4.Enter, NumericUpDown3.Enter, NumericUpDown2.Enter, NumericUpDown1.Enter
        sender.select(0, sender.value.ToString.Length)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox3.Text = TextBox1.Text + ".png"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        NumericUpDown1.Value = 0
        NumericUpDown2.Value = 0
        NumericUpDown3.Value = 0
        NumericUpDown4.Value = 0
        NumericUpDown5.Value = 0
        NumericUpDown6.Value = 0
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        RichTextBox4.Text = ""
        RichTextBox5.Text = ""
        RichTextBox6.Text = ""
    End Sub
End Class
