Public Class Form1
    Dim ConnOBJ As New ADODB.Connection
    Dim RecSetObj As New ADODB.Recordset
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        RecSetObj.MoveLast()
        Call ShowInForm()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ConnOBJ.Provider = "Microsoft.jet.oledb.4.0"
        ConnOBJ.ConnectionString = "C:\Users\jlwan\source\repos\vb\WindowsFormApp.NetFramework.vb\WindowsApp1\Database2.mdb"
        ConnOBJ.Open()
        RecSetObj.Open("Select * FROM SchoolRegister", ConnOBJ,
             ADODB.CursorTypeEnum.adOpenDynamic,
             ADODB.LockTypeEnum.adLockOptimistic)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        For Each TxtBox As TextBox In Me.Controls.OfType(Of TextBox)()
            If TxtBox.Text = "" Then
                MsgBox("you need to fill in all textbox")
                Exit Sub
            End If
        Next
        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            MsgBox("please select your gender")
            Exit Sub
        End If
        Dim StrCriteria As String = "StudentID = " & TextBox1.Text
        RecSetObj.MoveFirst()
        RecSetObj.Find(StrCriteria)
        If RecSetObj.EOF Then
            RecSetObj.AddNew()
            Call SaveInAccess()
            MsgBox("a new record has been added")
        Else
            MsgBox("Duplicate record")
        End If
    End Sub
    Private Sub SaveInAccess()
        RecSetObj.Fields("StudentID").Value = TextBox1.Text
        RecSetObj.Fields("FirstName").Value = TextBox2.Text
        RecSetObj.Fields("LastName").Value = TextBox3.Text
        RecSetObj.Fields("Nationality").Value = TextBox4.Text
        RecSetObj.Fields("Age").Value = TextBox5.Text
        If RadioButton1.Checked Then
            RecSetObj.Fields("Gender").Value = False
        Else
            RecSetObj.Fields("Gender").Value = True
        End If
        RecSetObj.Update()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        RecSetObj.MoveFirst()
        Call ShowInForm()
    End Sub

    Private Sub ShowInForm()
        TextBox1.Text = RecSetObj.Fields("StudentID").Value
        TextBox2.Text = RecSetObj.Fields("FirstName").Value
        TextBox3.Text = RecSetObj.Fields("LastName").Value
        TextBox4.Text = RecSetObj.Fields("Nationality").Value
        TextBox5.Text = RecSetObj.Fields("Age").Value
        If RecSetObj.Fields("Gender").Value Then
            RadioButton2.Checked = True
        Else
            RadioButton1.Checked = True
        End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        RecSetObj.MoveNext()
        If RecSetObj.EOF Then
            RecSetObj.MoveLast()
        Else
            Call ShowInForm()
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        RecSetObj.MovePrevious()
        If RecSetObj.BOF Then
            RecSetObj.MoveFirst()
        Else
            Call ShowInForm()
        End If

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Call Clear()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim StrCriteria As String = "StudentID = " & TextBox1.Text
        RecSetObj.MoveFirst()
        RecSetObj.Find(StrCriteria)
        If RecSetObj.EOF Then
            MsgBox("cannot find the StudentID")
            Call Clear()
        Else
            Call ShowInForm()
        End If
    End Sub

    Private Sub Clear()
        For Each TxtBox As TextBox In Me.Controls.OfType(Of TextBox)()
            TxtBox.Text = ""
        Next
        RadioButton1.Checked = False
        RadioButton2.Checked = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim StrCriteria As String = $"FirstName = '{TextBox2.Text}' "
        RecSetObj.MoveFirst()
        RecSetObj.Find(StrCriteria)
        If RecSetObj.EOF Then
            MsgBox("cannot find the First Name")
            Call Clear()
        Else
            Call ShowInForm()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim StrCriteria As String = $"LastName = '{TextBox3.Text}' "
        RecSetObj.MoveFirst()
        RecSetObj.Find(StrCriteria)
        If RecSetObj.EOF Then
            MsgBox("cannot find the Last Name")
            Call Clear()
        Else
            Call ShowInForm()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        For Each TxtBox As TextBox In Me.Controls.OfType(Of TextBox)()
            If TxtBox.Text = "" Then
                MsgBox("you need to fill in all textbox")
                Exit Sub
            End If
        Next
        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            MsgBox("please select your gender")
            Exit Sub
        End If
        Dim StrCriteria As String = "StudentID = " & TextBox1.Text
        RecSetObj.MoveFirst()
        RecSetObj.Find(StrCriteria)
        If RecSetObj.EOF Then
            RecSetObj.AddNew()
            Call SaveInAccess()
            MsgBox("a new record has benn added")
        Else
            Call SaveInAccess()
            MsgBox("the record has been modified  ")
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        For Each TxtBox As TextBox In Me.Controls.OfType(Of TextBox)()
            If TxtBox.Text = "" Then
                MsgBox("you need to fill in all textbox")
                Exit Sub
            End If
        Next
        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            MsgBox("please select your gender")
            Exit Sub
        End If
        Dim StrCriteria As String = "StudentID = " & TextBox1.Text
        RecSetObj.MoveFirst()
        RecSetObj.Find(StrCriteria)
        If RecSetObj.EOF Then
            MsgBox("no record can be deleted")
        Else
            Dim Confirm = MsgBox("do you really want to delete the record", vbYesNo)
            If Confirm = vbYes Then
                RecSetObj.Delete()
                Call Clear()
            Else
                Exit Sub
            End If
        End If
    End Sub
End Class
