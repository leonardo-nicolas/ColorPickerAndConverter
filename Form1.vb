#Region "License"
'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <https://www.gnu.org/licenses/>5.
#End Region



Imports System.ComponentModel
Imports System.Environment
Imports System.Globalization
Imports System.Runtime.CompilerServices
Imports System.Threading

Public Class Form1
    Public Sub New()
        Call InitializeComponent() 'This method should be called first!

        cboDelphiTypes.SelectedIndex = 0
        pcbColor.Cursor = New Cursor(Me.GetType, "dropper.cur")
        SelectingColor = False
        ControlKeyPressed = False
        AltKeyPressed = False
        cboLanguage.Items.Add(CultureInfo.CreateSpecificCulture("pt-br"))
        cboLanguage.Items.Add(CultureInfo.CreateSpecificCulture("en-us"))
    End Sub
    Dim ControlKeyPressed, AltKeyPressed, SelectingColor, UserChangingBoxes As Boolean

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.Control Then ControlKeyPressed = True
        If e.Alt Then AltKeyPressed = True
    End Sub

    Private Sub Form1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        If ControlKeyPressed And e.KeyChar = Convert.ToChar(18) Then 'Tecla R
            Call txtRed.Focus()
        ElseIf ControlKeyPressed AndAlso e.KeyChar = Convert.ToChar(7) Then 'Tecla G
            Call txtGreen.Focus()
        ElseIf ControlKeyPressed AndAlso e.KeyChar = Convert.ToChar(2) Then 'Tecla B
            Call txtBlue.Focus()
        ElseIf ControlKeyPressed And AltKeyPressed AndAlso e.KeyChar = Convert.ToChar(3) Then 'Tecla C
            Call txtCyan.Focus()
        ElseIf ControlKeyPressed And AltKeyPressed AndAlso e.KeyChar = Convert.ToChar(1) Then 'Tecla A
            Call txtAlpha.Focus()
        ElseIf ControlKeyPressed AndAlso e.KeyChar = Convert.ToChar(vbCr) Then 'Tecla M
            Call txtMagenta.Focus()
        ElseIf ControlKeyPressed And AltKeyPressed AndAlso e.KeyChar = Convert.ToChar(25) Then 'Tecla Y
            Call txtYellow.Focus()
        ElseIf ControlKeyPressed AndAlso e.KeyChar = Convert.ToChar(vbVerticalTab) Then 'Tecla K
            Call txtBlack.Focus()
        ElseIf ControlKeyPressed AndAlso e.KeyChar = Convert.ToChar(vbBack) Then 'Tecla H
            Call txtHtmlHex.Focus()
        ElseIf ControlKeyPressed And AltKeyPressed AndAlso e.KeyChar = Convert.ToChar(22) Then 'Tecla V
            Call txtVisualBasicHex.Focus()
        ElseIf ControlKeyPressed And AltKeyPressed AndAlso e.KeyChar = Convert.ToChar(24) Then 'Tecla X
            Call txtOtherLangHex.Focus()
        End If
    End Sub

    Private Sub Form1_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        If e.Control Then ControlKeyPressed = False
        If e.Alt Then AltKeyPressed = False
    End Sub

    Private Sub Form1_click(sender As Object, e As EventArgs) Handles MyBase.Click
        Call Me.Focus()
    End Sub

    Private Sub txtRed_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtYellow.KeyPress, txtRed.KeyPress,
                                                                                    txtMagenta.KeyPress, txtGreen.KeyPress, txtCyan.KeyPress,
                                                                                    txtBlue.KeyPress, txtBlack.KeyPress, txtAlpha.KeyPress
        If (sender Is txtRed Or sender Is txtGreen Or sender Is txtBlue) AndAlso IsNumeric(CType(sender, NumericUpDown).Text) AndAlso CInt(CType(sender, NumericUpDown).Text) > 255 Then
            e.Handled = True
        End If
    End Sub

    Dim AlphaValue As Double

    Private Sub pcbColor_MouseDown(sender As Object, e As MouseEventArgs) Handles pcbColor.MouseDown
        SelectingColor = True
    End Sub

    Private Sub pcbColor_MouseMove(sender As Object, e As MouseEventArgs) Handles pcbColor.MouseMove
        If SelectingColor Then
            Try
                Dim imgBmp As New Bitmap(pcbColor.Image)
                pnlSelectedColor.BackColor = imgBmp.GetPixel(e.X - 3, e.Y + 13)
            Catch
            End Try
        End If
    End Sub



    Private Sub pcbColor_MouseUp(sender As Object, e As MouseEventArgs) Handles pcbColor.MouseUp
        SelectingColor = False
    End Sub

    Private Sub txtHex_LostFocus(sender As Object, e As EventArgs) Handles txtVisualBasicHex.LostFocus, txtOtherLangHex.LostFocus, txtHtmlHex.LostFocus, txtDelphiHex.LostFocus
        UserChangingBoxes = False
        If sender Is txtVisualBasicHex Then
            If Not IsNumeric(CType(sender, TextBox).Text) Then
                Call CType(sender, TextBox).Focus()
            End If
        ElseIf sender Is txtOtherLangHex Then
            If Not CType(sender, TextBox).Text.StartsWith("0x") OrElse Not IsNumeric(CType(sender, TextBox).Text.Replace("0x", "&H")) Then
                Call CType(sender, TextBox).Focus()
            End If
        ElseIf sender Is txtHtmlHex Then
            If Not CType(sender, TextBox).Text.StartsWith("#") OrElse Not IsNumeric(CType(sender, TextBox).Text.Replace("#", "&H")) Then
                Call CType(sender, TextBox).Focus()
            End If
        ElseIf sender Is txtDelphiHex Then
            If Not CType(sender, TextBox).Text.StartsWith(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF")) OrElse Not IsNumeric(CType(sender, TextBox).Text.Replace(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF"), "&H")) Then
                Call CType(sender, TextBox).Focus()
            End If
        End If
    End Sub

    Private Sub TextBoxesAndNumericUpDownes_GotFocus(sender As Object, e As EventArgs) Handles txtYellow.GotFocus, txtVisualBasicHex.GotFocus, txtRed.GotFocus, txtOtherLangHex.GotFocus,
                                                                                               txtMagenta.GotFocus, txtHtmlHex.GotFocus, txtGreen.GotFocus, txtCyan.GotFocus, txtBlue.GotFocus,
                                                                                               txtBlack.GotFocus, txtAlpha.GotFocus, txtDelphiHex.GotFocus
        UserChangingBoxes = True
    End Sub

    Private Sub cboDelphiTypes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDelphiTypes.SelectedIndexChanged
        If Not txtDelphiHex.Focused And cboDelphiTypes.SelectedIndex = 0 Then
            txtDelphiHex.Text = "$00" + pnlSelectedColor.BackColor.B.ToString("X2") + pnlSelectedColor.BackColor.G.ToString("X2") + pnlSelectedColor.BackColor.R.ToString("X2")
        ElseIf Not txtDelphiHex.Focused And cboDelphiTypes.SelectedIndex = 1 Then
            txtDelphiHex.Text = "$FF" + pnlSelectedColor.BackColor.B.ToString("X2") + pnlSelectedColor.BackColor.G.ToString("X2") + pnlSelectedColor.BackColor.R.ToString("X2")
        End If
    End Sub

    Dim PrinterColor As CmykColor

    Private Sub pnlSelectedColor_Paint(sender As Object, e As EventArgs) Handles pnlSelectedColor.BackColorChanged
        PrinterColor = pnlSelectedColor.BackColor.ToCmykColor
        If Not txtRed.Focused Then txtRed.Value = Convert.ToDecimal(pnlSelectedColor.BackColor.R)
        If Not txtGreen.Focused Then txtGreen.Value = Convert.ToDecimal(pnlSelectedColor.BackColor.G)
        If Not txtBlue.Focused Then txtBlue.Value = Convert.ToDecimal(pnlSelectedColor.BackColor.B)
        If Not txtAlpha.Focused Then txtAlpha.Value = Convert.ToDecimal(pnlSelectedColor.BackColor.A)
        If Not txtHtmlHex.Focused Then
            txtHtmlHex.Text = "#" + pnlSelectedColor.BackColor.R.ToString("X2") + pnlSelectedColor.BackColor.G.ToString("X2") + pnlSelectedColor.BackColor.B.ToString("X2")
        End If
        If Not txtVisualBasicHex.Focused Then
            txtVisualBasicHex.Text = "&H" + pnlSelectedColor.BackColor.R.ToString("X2") + pnlSelectedColor.BackColor.G.ToString("X2") + pnlSelectedColor.BackColor.B.ToString("X2")
        End If
        If Not txtOtherLangHex.Focused Then
            txtOtherLangHex.Text = "0x" + pnlSelectedColor.BackColor.R.ToString("X2") + pnlSelectedColor.BackColor.G.ToString("X2") + pnlSelectedColor.BackColor.B.ToString("X2")
        End If
        If Not txtDelphiHex.Focused And cboDelphiTypes.SelectedIndex = 0 Then
            txtDelphiHex.Text = "$00" + pnlSelectedColor.BackColor.B.ToString("X2") + pnlSelectedColor.BackColor.G.ToString("X2") + pnlSelectedColor.BackColor.R.ToString("X2")
        ElseIf Not txtDelphiHex.Focused And cboDelphiTypes.SelectedIndex = 1 Then
            txtDelphiHex.Text = "$FF" + pnlSelectedColor.BackColor.B.ToString("X2") + pnlSelectedColor.BackColor.G.ToString("X2") + pnlSelectedColor.BackColor.R.ToString("X2")
        End If
            If Not txtCyan.Focused Then txtCyan.Value = Convert.ToDecimal(PrinterColor.C * 100)
        If Not txtMagenta.Focused Then txtMagenta.Value = Convert.ToDecimal(PrinterColor.M * 100)
        If Not txtYellow.Focused Then txtYellow.Value = Convert.ToDecimal(PrinterColor.Y * 100)
        If Not txtBlack.Focused Then txtBlack.Value = Convert.ToDecimal(PrinterColor.K * 100)
    End Sub

    Private Sub ButtonsToCopy_Click(sender As Object, e As EventArgs) Handles btnCopyRed.Click, btnCopyYellow.Click, btnCopyVisualBasicHex.Click, btnCopyOthersLangHex.Click,
                                                                                btnCopyMagenta.Click, btnCopyHtmlHex.Click, btnCopyGreen.Click, btnCopyDelphiHex.Click, btnCopyCyan.Click,
                                                                                btnCopyBlue.Click, btnCopyBlack.Click, btnCopyAlpha.Click
        If sender Is btnCopyRed Then
            Call Clipboard.SetText(txtRed.Value.ToString)
        ElseIf sender Is btnCopyGreen Then
            Call Clipboard.SetText(txtGreen.Value.ToString)
        ElseIf sender Is btnCopyBlue Then
            Call Clipboard.SetText(txtBlue.Value.ToString)
        ElseIf sender Is btnCopyAlpha Then
            Call Clipboard.SetText(txtAlpha.Value.ToString)
        ElseIf sender Is btnCopyCyan Then
            Call Clipboard.SetText(txtCyan.Value.ToString)
        ElseIf sender Is btnCopyMagenta Then
            Call Clipboard.SetText(txtMagenta.Value.ToString)
        ElseIf sender Is btnCopyYellow Then
            Call Clipboard.SetText(txtYellow.Value.ToString)
        ElseIf sender Is btnCopyBlack Then
            Call Clipboard.SetText(txtBlack.Value.ToString)
        ElseIf sender Is btnCopyHtmlHex Then
            Call Clipboard.SetText(txtHtmlHex.Text.ToString)
        ElseIf sender Is btnCopyVisualBasicHex Then
            Call Clipboard.SetText(txtVisualBasicHex.Text.ToString)
        ElseIf sender Is btnCopyDelphiHex Then
            Call Clipboard.SetText(txtDelphiHex.Text.ToString)
        ElseIf sender Is btnCopyOthersLangHex Then
            Call Clipboard.SetText(txtOtherLangHex.Text.ToString)
        End If
        MessageBox.Show("Successfully Copied to Clipboard! :D \o/" + NewLine + "Just paste it into your file You're working on! :)")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboLanguage.SelectedIndexChanged
        Dim Culture As String = CType(cboLanguage.SelectedItem, CultureInfo).Name
        Thread.CurrentThread.CurrentCulture = New CultureInfo(culture, True)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(culture, True)
        Dim resx As New ComponentResourceManager(GetType(Form1))

        Me.Text = resx.GetObject("$this.Text", Thread.CurrentThread.CurrentCulture).ToString()
        For Each ctrl As Control In Me.Controls
            Call resx.ApplyResources(ctrl, ctrl.Name, Thread.CurrentThread.CurrentCulture)
        Next ctrl
    End Sub

    Private Sub NumericUpDownes_LostFocus(sender As Object, e As EventArgs) Handles txtYellow.LostFocus, txtRed.LostFocus, txtMagenta.LostFocus,
                                                                                        txtGreen.LostFocus, txtCyan.LostFocus, txtBlue.LostFocus,
                                                                                        txtBlack.LostFocus, txtAlpha.LostFocus
        UserChangingBoxes = False
    End Sub

    Private Sub Form1_HelpButtonClicked(sender As Object, e As CancelEventArgs) Handles MyBase.HelpButtonClicked
        e.Cancel = True
        Dim Message As String = String.Empty
        Message += "Author: Leonardo Nicolas" + NewLine
        Message += "Author's GitHub: https://github.com/leonardon397" + NewLine
        Message += "Author's LinkedIN: https://www.linkedin.com/in/leonardo-nicolas-sales-dias-2a3892149/" + NewLine
        Message += "This project uses CMYK from third party, where the Font's link is: https://www.cyotek.com/blog/converting-colours-between-rgb-and-cmyk-in-csharp" + NewLine
        Message += "Written with VB.NET (Visual Basic DotNet)" + NewLine
        Message += "The CMYK source code has been converted from C# to VB.NET." + NewLine + NewLine
        Message += "This software is free for all developers! And its license is under GPL" + NewLine
        Message += "Colors Picker is Open Source and you can get it from GitHub." + NewLine
        Message += "Repository link: https://github.com/leonardon397/ColorPickerAndConverter"
        Message += "This program is free software: you can redistribute it and/or modify" + NewLine + NewLine
        Message += "License: " + NewLine
        Message += "it under the terms of the GNU General Public License as published by" + NewLine
        Message += "the Free Software Foundation, either version 3 of the License, or" + newline
        message += "(at your option) any later version." + newline
        Message += NewLine
        Message += "Colors Picker is distributed in the hope that it will be useful," + NewLine
        Message += "but WITHOUT ANY WARRANTY; without even the implied warranty of" + NewLine
        Message += "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the" + NewLine
        Message += "GNU General Public License for more details." + NewLine
        Message += NewLine
        Message += "You should have received a copy of the GNU General Public License" + NewLine
        Message += "along with this program.  If not, see <https://www.gnu.org/licenses/>5." + NewLine
        Call MessageBox.Show(Message)
    End Sub


    Private Sub NumericUpDownes_TextChanged(sender As Object, e As EventArgs) Handles txtRed.ValueChanged, txtGreen.ValueChanged, txtBlue.ValueChanged,
                                                                                        txtAlpha.ValueChanged, txtCyan.ValueChanged, txtMagenta.ValueChanged,
                                                                                        txtYellow.ValueChanged, txtBlack.ValueChanged, txtHtmlHex.TextChanged,
                                                                                        txtVisualBasicHex.TextChanged, txtOtherLangHex.TextChanged, txtDelphiHex.TextChanged
        Dim CodeHexColorToFill As Integer = BackColor.ToArgb
        Try
            Dim StrCodeComplete, StrCodeRedColor, strCodeGreenColor, strCodeBlueColor, strCodeAlphaColor As String
            StrCodeRedColor = "FF"
            strCodeGreenColor = "FF"
            strCodeBlueColor = "FF"
            strCodeAlphaColor = "FF"
            If CType(sender, Control).Focused AndAlso (sender Is txtRed Or sender Is txtGreen Or sender Is txtBlue Or sender Is txtAlpha) Then
                CodeHexColorToFill = Color.FromArgb(txtAlpha.Value, txtRed.Value, txtGreen.Value, txtBlue.Value).ToArgb
            ElseIf CType(sender, Control).Focused AndAlso (sender Is txtCyan Or sender Is txtMagenta Or sender Is txtYellow Or sender Is txtBlack) Then
                CodeHexColorToFill = ColorConvert.ConvertCmykToRgb(txtCyan.Value / 100, txtMagenta.Value / 100, txtYellow.Value / 100, txtBlack.Value / 100).ToArgb
            ElseIf CType(sender, Control).Focused AndAlso sender Is txtHtmlHex Then 'HTML and CSS
                If Not txtHtmlHex.Text.StartsWith("#") OrElse Not IsNumeric(txtHtmlHex.Text.Replace("#", "&H")) Then
                    Exit Sub
                End If
                CodeHexColorToFill = ColorTranslator.FromHtml(txtHtmlHex.Text).ToArgb
            ElseIf CType(sender, Control).Focused AndAlso sender Is txtVisualBasicHex Then 'Visual Basic
                If Not IsNumeric(txtVisualBasicHex.Text) Then
                    Exit Sub
                ElseIf txtVisualBasicHex.Text.Replace("&H", "").Length = 2 Then
                    StrCodeComplete = txtVisualBasicHex.Text.Replace("&H", "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                ElseIf txtVisualBasicHex.Text.Replace("&H", "").Length = 4 Then
                    StrCodeComplete = txtVisualBasicHex.Text.Replace("&H", "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                    strCodeGreenColor = StrCodeComplete.Substring(2, 2)
                ElseIf txtVisualBasicHex.Text.Replace("&H", "").Length = 6 Then
                    StrCodeComplete = txtVisualBasicHex.Text.Replace("&H", "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                    strCodeGreenColor = StrCodeComplete.Substring(2, 2)
                    strCodeBlueColor = StrCodeComplete.Substring(4, 2)
                ElseIf txtVisualBasicHex.Text.Replace("&H", "").Length = 8 Then
                    StrCodeComplete = txtVisualBasicHex.Text.Replace("&H", "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                    strCodeGreenColor = StrCodeComplete.Substring(2, 2)
                    strCodeBlueColor = StrCodeComplete.Substring(4, 2)
                    strCodeAlphaColor = StrCodeComplete.Substring(6, 2)
                End If
            ElseIf CType(sender, Control).Focused AndAlso sender Is txtDelphiHex Then 'Delphi or Pascal
                If Not txtDelphiHex.Text.StartsWith(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF")) OrElse Not IsNumeric(txtDelphiHex.Text.Replace(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF"), "&H")) Then
                    Exit Sub
                ElseIf txtDelphiHex.Text.Replace(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF"), "").Length = 2 Then
                    StrCodeComplete = txtDelphiHex.Text.Replace(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF"), "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                ElseIf txtDelphiHex.Text.Replace(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF"), "").Length = 4 Then
                    StrCodeComplete = txtDelphiHex.Text.Replace(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF"), "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                    strCodeGreenColor = StrCodeComplete.Substring(2, 2)
                ElseIf txtDelphiHex.Text.Replace(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF"), "").Length = 6 Then
                    StrCodeComplete = txtDelphiHex.Text.Replace(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF"), "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                    strCodeGreenColor = StrCodeComplete.Substring(2, 2)
                    strCodeBlueColor = StrCodeComplete.Substring(4, 2)
                ElseIf txtDelphiHex.Text.Replace(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF"), "").Length = 8 Then
                    StrCodeComplete = txtDelphiHex.Text.Replace(If(cboDelphiTypes.SelectedIndex = 0, "$00", "$FF"), "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                    strCodeGreenColor = StrCodeComplete.Substring(2, 2)
                    strCodeBlueColor = StrCodeComplete.Substring(4, 2)
                    strCodeAlphaColor = StrCodeComplete.Substring(6, 2)
                End If
            ElseIf CType(sender, Control).Focused AndAlso sender Is txtOtherLangHex Then 'Other languages (Objective C, C++, C#, Java, Pyton, Javascript, Kotlin, Dart, etc)
                If Not txtOtherLangHex.Text.StartsWith("0x") OrElse Not IsNumeric(txtOtherLangHex.Text.Replace("0x", "&H")) Then
                    Exit Sub
                ElseIf txtOtherLangHex.Text.Replace("0x", "").Length = 2 Then
                    StrCodeComplete = txtOtherLangHex.Text.Replace("0x", "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                ElseIf txtOtherLangHex.Text.Replace("0x", "").Length = 4 Then
                    StrCodeComplete = txtOtherLangHex.Text.Replace("0x", "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                    strCodeGreenColor = StrCodeComplete.Substring(2, 2)
                ElseIf txtOtherLangHex.Text.Replace("0x", "").Length = 6 Then
                    StrCodeComplete = txtOtherLangHex.Text.Replace("0x", "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                    strCodeGreenColor = StrCodeComplete.Substring(2, 2)
                    strCodeBlueColor = StrCodeComplete.Substring(4, 2)
                ElseIf txtOtherLangHex.Text.Replace("0x", "").Length = 8 Then
                    StrCodeComplete = txtOtherLangHex.Text.Replace("0x", "").ToUpper
                    StrCodeRedColor = StrCodeComplete.Substring(0, 2)
                    strCodeGreenColor = StrCodeComplete.Substring(2, 2)
                    strCodeBlueColor = StrCodeComplete.Substring(4, 2)
                    strCodeAlphaColor = StrCodeComplete.Substring(6, 2)
                End If
            Else
                Exit Sub
            End If
            If CType(sender, Control).Focused AndAlso (sender Is txtOtherLangHex Or sender Is txtVisualBasicHex Or sender Is txtDelphiHex) Then
                CodeHexColorToFill = Convert.ToInt32(strCodeBlueColor + strCodeGreenColor + StrCodeRedColor, 16)
            End If
        Catch ex As Exception
            Exit Sub
        End Try
        pnlSelectedColor.BackColor = Color.FromArgb(CodeHexColorToFill)
    End Sub

End Class

Public Structure CmykColor
    Private ReadOnly _c As Single
    Private ReadOnly _m As Single
    Private ReadOnly _y As Single
    Private ReadOnly _k As Single

    Public Sub New(ByVal c As Single, ByVal m As Single, ByVal y As Single, ByVal k As Single)
        _c = c
        _m = m
        _y = y
        _k = k
    End Sub

    Public ReadOnly Property C As Single
        Get
            Return _c
        End Get
    End Property

    Public ReadOnly Property M As Single
        Get
            Return _m
        End Get
    End Property

    Public ReadOnly Property Y As Single
        Get
            Return _y
        End Get
    End Property

    Public ReadOnly Property K As Single
        Get
            Return _k
        End Get
    End Property
End Structure

Module ColorConvert
    Function ConvertCmykToRgb(ByVal c As Single, ByVal m As Single, ByVal y As Single, ByVal k As Single) As Color
        Dim r, g, b As Integer
        r = Convert.ToInt32(255 * (1 - c) * (1 - k))
        g = Convert.ToInt32(255 * (1 - m) * (1 - k))
        b = Convert.ToInt32(255 * (1 - y) * (1 - k))
        Return Color.FromArgb(r, g, b)
    End Function

    Function ConvertRgbToCmyk(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As CmykColor
        Dim c, m, y, k, rf, gf, bf As Single
        rf = r / 255.0F
        gf = g / 255.0F
        bf = b / 255.0F
        k = ClampCmyk(1 - Math.Max(Math.Max(rf, gf), bf))
        c = ClampCmyk((1 - rf - k) / (1 - k))
        m = ClampCmyk((1 - gf - k) / (1 - k))
        y = ClampCmyk((1 - bf - k) / (1 - k))
        Return New CmykColor(c, m, y, k)
    End Function

    Private Function ClampCmyk(ByVal value As Single) As Single
        ClampCmyk = If(value < 0 OrElse Single.IsNaN(value), 0, value)
    End Function


    <Extension>
    Public Function ToCmykColor(CurrentRgbColor As Color) As CmykColor
        ToCmykColor = ColorConvert.ConvertRgbToCmyk(CurrentRgbColor.R, CurrentRgbColor.G, CurrentRgbColor.B)
    End Function
End Module