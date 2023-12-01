Imports Microsoft.Office.Interop
Imports System
Imports System.Net

Public Class Form1
    'VARIABEL ALAT MUSIK
    Dim music As Integer
    'VARIABEL ARRAY
    Dim array(99) As String
    Dim jmlArray As Integer
    Dim isiArray As String
    Public score As Integer = 0
    Dim kill As Boolean
    Dim ran As New Random()
    Dim i As Integer
    Dim health As Integer = 100
    'PILIH KEYBOARD
    Private Sub keyboardLink_Click_1(sender As Object, e As EventArgs) Handles keyboardLink.Click
        music = 1
        pianoForm.Visible = True
        GunaPanel1.Visible = False
        onChord.Visible = True
    End Sub
    'PILIH KALIMBA
    Private Sub kalimbaLink_Click(sender As Object, e As EventArgs)
        music = 2
    End Sub
    'FORM LOAD
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        array = {"a", "b", "c", "d", "e", "f", "g", "c1", "d1", "e1", "f1"}
        jmlArray = UBound(array)
        batas.UseTransfarantBackground = 50
        point.Text = "SCORE "
        Label22.Text = "point : " & pointC.Location.ToString
        ProgressBar1.Value = 100
        i = 0
    End Sub

    'START DROP
    Private Async Sub start()
        GunaPictureBox1.BackColor = Color.White
        GunaPictureBox6.Height = GunaPictureBox1.Height
        GunaPictureBox2.BackColor = Color.White
        GunaPictureBox3.BackColor = Color.White
        GunaPictureBox4.BackColor = Color.White
        GunaPictureBox5.BackColor = Color.White
        GunaPictureBox6.BackColor = Color.White
        GunaPictureBox7.BackColor = Color.White
        GunaPictureBox8.BackColor = Color.White
        GunaPictureBox9.BackColor = Color.White
        GunaPictureBox10.BackColor = Color.White
        GunaPictureBox11.BackColor = Color.White
        GunaPictureBox1.Location = New Point(31, -173)
        GunaPictureBox2.Location = New Point(103, -173)
        GunaPictureBox3.Location = New Point(175, -173)
        GunaPictureBox4.Location = New Point(247, -173)
        GunaPictureBox5.Location = New Point(319, -173)
        GunaPictureBox6.Location = New Point(391, -173)
        GunaPictureBox7.Location = New Point(463, -173)
        GunaPictureBox8.Location = New Point(535, -173)
        GunaPictureBox9.Location = New Point(607, -173)
        GunaPictureBox10.Location = New Point(679, -173)
        GunaPictureBox11.Location = New Point(751, -173)

        'PENGAMBILAN CHORD DI ARRAY
        'For i As Integer = 0 To jmlArray

        'Label21.Text = array(i)
        ' isiArray = array(i)
        'Timer1.Start()
        'Await Task.Delay(TimeSpan.FromSeconds(1.3))
        'If kill = True Then
        'Exit For
        ' End If
        ' Next

        While ProgressBar1.Value >= 0
            isiArray = array(ran.Next(11))
            Timer1.Start()
            Label21.Text = isiArray
            Await Task.Delay(TimeSpan.FromSeconds(1.3))
            If kill = True Then
                Exit While
            ElseIf ProgressBar1.Value = 0 Then
                GunaShadowPanel1.Visible = True
                scoreAkhir.Text = point.ToString
            End If
        End While
        'If ProgressBar1.Value = 0 Then
        'GunaShadowPanel1.Visible = True
        'scoreAkhir.Text = point.ToString
        'End If
    End Sub

    'ON RADIO BUTTON
    Private Sub onChord_CheckedChanged(sender As Object, e As EventArgs) Handles onChord.CheckedChanged
        start()
        point.Visible = True
        onChord.Visible = False
        offChord.Visible = True
        batas.Visible = True
        kill = False
    End Sub

    'OFF RADIO BUTTON
    Private Sub offChord_CheckedChanged(sender As Object, e As EventArgs) Handles offChord.CheckedChanged
        onChord.Visible = True
        offChord.Visible = False
        Timer1.Stop()
        Timer2.Stop()
        Timer3.Stop()
        Timer4.Stop()
        Timer5.Stop()
        Timer6.Stop()
        Timer7.Stop()
        Timer8.Stop()
        Timer9.Stop()
        Timer10.Stop()
        Timer11.Stop()
        Timer12.Stop()
        Timer13.Stop()
        kill = True
        point.Visible = False
        batas.Visible = False
    End Sub


    'TAMBAH POINT
    Public Sub pointPls()
        If GunaPictureBox3.Location.Y > 200 And GunaPictureBox3.Location.Y < 300 Then 'BOX3
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox3.Location = New Point(31, -173)
            Timer6.Stop()
        End If
        If GunaPictureBox4.Location.Y > 200 And GunaPictureBox4.Location.Y < 300 Then 'BOX4
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox4.Location = New Point(31, -173)
            Timer7.Stop()
        End If
        If GunaPictureBox5.Location.Y > 200 And GunaPictureBox5.Location.Y < 300 Then 'BOX5
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox5.Location = New Point(31, -173)
            Timer8.Stop()
        End If

        If GunaPictureBox7.Location.Y > 200 And GunaPictureBox7.Location.Y < 300 Then 'BOX7
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox7.Location = New Point(31, -173)
            Timer3.Stop()
        End If
        If GunaPictureBox8.Location.Y > 200 And GunaPictureBox8.Location.Y < 300 Then 'BOX8
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox8.Location = New Point(31, -173)
            Timer9.Stop()
        End If
        If GunaPictureBox9.Location.Y > 200 And GunaPictureBox9.Location.Y < 300 Then 'BOX9
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox9.Location = New Point(31, -173)
            Timer10.Stop()
        End If
        If GunaPictureBox10.Location.Y > 200 And GunaPictureBox10.Location.Y < 300 Then 'BOX10
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox10.Location = New Point(31, -173)
            Timer11.Stop()
        End If
        If GunaPictureBox11.Location.Y > 200 And GunaPictureBox11.Location.Y < 300 Then 'BOX11
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox11.Location = New Point(31, -173)
            Timer12.Stop()
        End If
        If GunaPictureBox1.Location.Y > 200 And GunaPictureBox1.Location.Y < 300 Then 'BOX1
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox1.Location = New Point(31, -173)
            Timer4.Stop()
        End If

        If GunaPictureBox2.Location.Y > 200 And GunaPictureBox2.Location.Y < 300 Then 'BOX2
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox2.Location = New Point(31, -173)
            Timer5.Stop()
        End If

        If GunaPictureBox6.Location.Y > 200 And GunaPictureBox6.Location.Y < 300 Then 'BOX2
            pointC.BaseColor = Color.White
            Label22.Text = "point : " & pointC.Location.ToString
            score += 1
            point.Text = "SCORE : " & score
            GunaPictureBox6.Location = New Point(31, -173)
            Timer2.Stop()
        End If
    End Sub

    'PENAMPIL POINT
    Private Async Sub showPoint()
        Await Task.Delay(TimeSpan.FromSeconds(7))
        score = 0
    End Sub

    'FUNGSI KEYBOARD
    Private Async Sub keyboardLink_KeyDown(sender As Object, e As KeyEventArgs) Handles keyboardLink.KeyDown
        If e.KeyCode = Keys.A Then
            My.Computer.Audio.Play(My.Resources.C, AudioPlayMode.Background)
            If GunaPictureBox1.Location.Y > 190 And GunaPictureBox1.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox1.Location = New Point(31, -173)
                Timer4.Stop()
            End If
        End If
        If e.KeyCode = Keys.S Then
            My.Computer.Audio.Play(My.Resources.D, AudioPlayMode.Background)
            If GunaPictureBox2.Location.Y > 190 And GunaPictureBox2.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox2.Location = New Point(31, -173)
                Timer5.Stop()
            End If
        End If
        If e.KeyCode = Keys.D Then
            My.Computer.Audio.Play(My.Resources.E, AudioPlayMode.Background)
            If GunaPictureBox3.Location.Y > 190 And GunaPictureBox3.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox3.Location = New Point(31, -173)
                Timer6.Stop()
            End If
        End If
        If e.KeyCode = Keys.F Then
            My.Computer.Audio.Play(My.Resources.F, AudioPlayMode.Background)
            If GunaPictureBox4.Location.Y > 190 And GunaPictureBox4.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox4.Location = New Point(31, -173)
                Timer7.Stop()
            End If
        End If
        If e.KeyCode = Keys.G Then
            My.Computer.Audio.Play(My.Resources.G, AudioPlayMode.Background)
            If GunaPictureBox5.Location.Y > 190 And GunaPictureBox5.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox5.Location = New Point(31, -173)
                Timer8.Stop()
            End If
        End If
        If e.KeyCode = Keys.H Then
            My.Computer.Audio.Play(My.Resources.A, AudioPlayMode.Background)
            If GunaPictureBox6.Location.Y > 190 And GunaPictureBox6.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox6.Location = New Point(31, -173)
                Timer2.Stop()
            End If
        End If
        If e.KeyCode = Keys.J Then
            My.Computer.Audio.Play(My.Resources.B, AudioPlayMode.Background)
            If GunaPictureBox7.Location.Y > 190 And GunaPictureBox7.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox7.Location = New Point(31, -173)
                Timer3.Stop()
            End If
        End If
        If e.KeyCode = Keys.K Then
            My.Computer.Audio.Play(My.Resources.C1, AudioPlayMode.Background)
            If GunaPictureBox8.Location.Y > 190 And GunaPictureBox8.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox8.Location = New Point(31, -173)
                Timer9.Stop()
            End If
        End If
        If e.KeyCode = Keys.L Then
            My.Computer.Audio.Play(My.Resources.D1, AudioPlayMode.Background)
            If GunaPictureBox9.Location.Y > 190 And GunaPictureBox9.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox9.Location = New Point(31, -173)
                Timer10.Stop()
            End If
        End If
        If e.KeyCode = Keys.OemSemicolon Then
            My.Computer.Audio.Play(My.Resources.E1, AudioPlayMode.Background)
            If GunaPictureBox10.Location.Y > 190 And GunaPictureBox10.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox10.Location = New Point(31, -173)
                Timer11.Stop()
            End If
        End If
        If e.KeyCode = Keys.OemQuotes Then
            My.Computer.Audio.Play(My.Resources.F1, AudioPlayMode.Background)
            If GunaPictureBox11.Location.Y > 190 And GunaPictureBox11.Location.Y < 259 Then
                pointC.BaseColor = Color.White
                Label22.Text = "point : " & pointC.Location.ToString
                score += 1
                ProgressBar1.Value += 1
                point.Text = "SCORE : " & score
                GunaPictureBox11.Location = New Point(31, -173)
                Timer12.Stop()
            End If
        End If
    End Sub

    'DROP CHORD (lanjut dari ambil array)
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label19.Text = GunaPictureBox3.Location.ToString
        Label20.Text = ProgressBar1.Value
        If isiArray = "a" Then
            Timer2.Start()
        End If
        If isiArray = "b" Then
            Timer3.Start()
        End If
        If isiArray = "c" Then
            Timer4.Start()
        End If
        If isiArray = "d" Then
            Timer5.Start()
        End If
        If isiArray = "e" Then
            Timer6.Start()
        End If
        If isiArray = "f" Then
            Timer7.Start()
        End If
        If isiArray = "g" Then
            Timer8.Start()
        End If
        If isiArray = "c1" Then
            Timer9.Start()
        End If
        If isiArray = "d1" Then
            Timer10.Start()
        End If
        If isiArray = "e1" Then
            Timer11.Start()
        End If
        If isiArray = "f1" Then
            Timer12.Start()
        End If
        If isiArray = "end" Then
            showPoint()
        End If
    End Sub

    'DROP CHORD (2)
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        GunaPictureBox6.Top += 6
        Label19.Text = batas.Location.ToString
        Label20.Text = GunaPictureBox6.Location.ToString
        If GunaPictureBox6.Location.Y = batas.Location.Y Then
            Timer2.Stop()
            GunaPictureBox6.Location = New Point(31, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        GunaPictureBox7.Top += 6
        If GunaPictureBox7.Location.Y = batas.Location.Y Then
            Timer3.Stop()
            GunaPictureBox7.Location = New Point(463, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        GunaPictureBox1.Top += 6
        If GunaPictureBox1.Location.Y = batas.Location.Y Then
            Timer4.Stop()
            GunaPictureBox1.Location = New Point(31, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        GunaPictureBox2.Top += 6
        If GunaPictureBox2.Location.Y = batas.Location.Y Then
            Timer5.Stop()
            GunaPictureBox2.Location = New Point(103, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub Timer6_Tick(sender As Object, e As EventArgs) Handles Timer6.Tick
        GunaPictureBox3.Top += 6
        If GunaPictureBox3.Location.Y = batas.Location.Y Then
            Timer6.Stop()
            GunaPictureBox3.Location = New Point(175, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub Timer7_Tick(sender As Object, e As EventArgs) Handles Timer7.Tick
        GunaPictureBox4.Top += 6
        If GunaPictureBox4.Location.Y = batas.Location.Y Then
            Timer7.Stop()
            GunaPictureBox4.Location = New Point(391, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub Timer8_Tick(sender As Object, e As EventArgs) Handles Timer8.Tick
        GunaPictureBox5.Top += 6
        If GunaPictureBox5.Location.Y = batas.Location.Y Then
            Timer8.Stop()
            GunaPictureBox5.Location = New Point(319, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub Timer9_Tick(sender As Object, e As EventArgs) Handles Timer9.Tick
        GunaPictureBox8.Top += 6
        If GunaPictureBox8.Location.Y = batas.Location.Y Then
            Timer9.Stop()
            GunaPictureBox8.Location = New Point(535, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub Timer10_Tick(sender As Object, e As EventArgs) Handles Timer10.Tick
        GunaPictureBox9.Top += 6
        If GunaPictureBox9.Location.Y = batas.Location.Y Then
            Timer10.Stop()
            GunaPictureBox9.Location = New Point(607, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub Timer11_Tick(sender As Object, e As EventArgs) Handles Timer11.Tick
        GunaPictureBox10.Top += 6
        If GunaPictureBox10.Location.Y = batas.Location.Y Then
            Timer11.Stop()
            GunaPictureBox10.Location = New Point(679, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub Timer12_Tick(sender As Object, e As EventArgs) Handles Timer12.Tick
        GunaPictureBox11.Top += 6
        If GunaPictureBox11.Location.Y = batas.Location.Y Then
            Timer12.Stop()
            GunaPictureBox11.Location = New Point(751, -173)
            ProgressBar1.Value -= 2
        End If
    End Sub

    Private Sub GunaButton2_Click(sender As Object, e As EventArgs) Handles GunaButton2.Click
        GunaShadowPanel1.Visible = False
        scoreAkhir.Text = 0
        point.Text = 0
        score = 0
    End Sub
End Class