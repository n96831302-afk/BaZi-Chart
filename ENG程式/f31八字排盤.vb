Option Explicit On : Option Infer On : Option Strict On
Imports System.CodeDom
Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO, System.Text
Imports Microsoft.VisualBasic.Devices

Public Class f31八字排盤
	Private Sub frm_MouseClick(sender As Object, e As MouseEventArgs) Handles MyBase.MouseClick
		dgv60甲子.Hide()
	End Sub
	Private Sub frm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
	End Sub
	Private Sub frm_HandleCreated(sender As Object, e As EventArgs) Handles MyBase.HandleCreated
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		'Me.AcceptButton = btn搜尋 ' 不要用
		us部署()
		m易理.us_基本資料_干支()
		m易理.us_基本資料_八卦()
		m易理.us_dgv60甲子(dgv60甲子)
		us起始設定()
	End Sub
	Private Sub frm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		'Console.WriteLine(#02/04/2024#.AddDays(182).ToShortDateString) ' 干支的年中 8/4日
		'---------------------------------------------
		'Me.WindowState = FormWindowState.Maximized
		Dim wa = Screen.PrimaryScreen.WorkingArea, iwaL = wa.Left, iwaW = wa.Width, iwaH = wa.Height
		Me.Location = New Point(0, 0) : Me.Size = New Size(CInt(iwaW * 0.9), iwaH)

		btn排盤.PerformClick()
	End Sub

	Private Sub frm_SizeChanged(sender As Object, e As EventArgs) Handles MyBase.SizeChanged
		If Not Me.Visible Then Exit Sub
		Static Me_WinSta1 As FormWindowState : Dim Me_WinSta0 = Me.WindowState '  1=FormWindowState.Minimized
		If Me_WinSta0 = 1 OrElse Me_WinSta1 = 1 Then Me_WinSta1 = Me_WinSta0 : Exit Sub ' (1)剛才Normal，現在 min；(2)假如剛才是"最小化"，表示現在是 *.Normal
		'---------------------------------------------
		Me.SuspendLayout()
		Dim icsW = Me.ClientSize.Width, icsH = Me.ClientSize.Height
		Dim iH As Int32
		'---------------------------------------------
		iH = txt組合_傳統法.Top + txt組合_傳統法.Height
		txt流年_傳統法.Height = icsH - iH
		txt流年_壬癸己.Height = icsH - iH
		txt流年_六爻法.Height = icsH - iH
		Me.ResumeLayout() : My.Application.DoEvents()
	End Sub
	Dim _i寬度 As Int32 = 530
	Private Sub us位置變換()
		Dim icsW = Me.ClientSize.Width, icsH = Me.ClientSize.Height
		Dim iL, iH, iW As Int32
		iH = txt組合_傳統法.Top + txt組合_傳統法.Height
		txt流年_傳統法.Height = icsH - iH
		txt流年_壬癸己.Height = icsH - iH
		txt流年_六爻法.Height = icsH - iH
		Me.ResumeLayout() : My.Application.DoEvents()
		'---------------------------------------------
		iL = txt統計5.Left + txt統計5.Width : iW = _i寬度
		Dim i大運 = 130, i八字 = 200, i組合 = 200, i流年 = icsH - i大運 - i八字 - i組合
		If txt大運_壬癸己.Left = iL Then
			txt大運_傳統法.Size = New Size(iW, i大運) : txt大運_傳統法.Location = New Point(iL, 0)
			txt八字_傳統法.Size = New Size(iW, i八字) : txt八字_傳統法.Location = New Point(iL, txt大運_傳統法.Height)
			txt組合_傳統法.Size = New Size(iW, i組合) : txt組合_傳統法.Location = New Point(iL, txt八字_傳統法.Top + txt八字_傳統法.Height)
			txt流月_傳統法.Size = txt組合_傳統法.Size : txt流月_傳統法.Location = txt組合_傳統法.Location
			txt流年_傳統法.Size = New Size(iW, i流年) : txt流年_傳統法.Location = New Point(iL, txt組合_傳統法.Top + txt組合_傳統法.Height)

			txt大運_六爻法.Size = New Size(iW, i大運) : txt大運_六爻法.Location = New Point(iL, 0)
			txt八字_六爻法.Size = New Size(iW, i八字) : txt八字_六爻法.Location = New Point(iL, txt大運_六爻法.Height)
			txt組合_六爻法.Size = New Size(iW, i組合) : txt組合_六爻法.Location = New Point(iL, txt八字_六爻法.Top + txt八字_六爻法.Height)
			txt流月_六爻法.Size = txt組合_六爻法.Size : txt流月_六爻法.Location = txt組合_六爻法.Location
			txt流年_六爻法.Size = New Size(iW, i流年) : txt流年_六爻法.Location = New Point(iL, txt組合_六爻法.Top + txt組合_六爻法.Height)

			iL += iW
			txt大運_壬癸己.Size = New Size(iW, i大運) : txt大運_壬癸己.Location = New Point(iL, 0)
			txt八字_壬癸己.Size = New Size(iW, i八字) : txt八字_壬癸己.Location = New Point(iL, txt大運_壬癸己.Height)
			txt組合_壬癸己.Size = New Size(iW, i組合) : txt組合_壬癸己.Location = New Point(iL, txt八字_壬癸己.Top + txt八字_壬癸己.Height)
			txt流月_壬癸己.Size = txt組合_壬癸己.Size : txt流月_壬癸己.Location = txt組合_壬癸己.Location
			txt流年_壬癸己.Size = New Size(iW, i流年) : txt流年_壬癸己.Location = New Point(iL, txt組合_壬癸己.Top + txt組合_壬癸己.Height)
		Else
			txt大運_壬癸己.Size = New Size(iW, i大運) : txt大運_壬癸己.Location = New Point(iL, 0)
			txt八字_壬癸己.Size = New Size(iW, i八字) : txt八字_壬癸己.Location = New Point(iL, txt大運_壬癸己.Height)
			txt組合_壬癸己.Size = New Size(iW, i組合) : txt組合_壬癸己.Location = New Point(iL, txt八字_壬癸己.Top + txt八字_壬癸己.Height)
			txt流月_壬癸己.Size = txt組合_壬癸己.Size : txt流月_壬癸己.Location = txt組合_壬癸己.Location
			txt流年_壬癸己.Size = New Size(iW, i流年) : txt流年_壬癸己.Location = New Point(iL, txt組合_壬癸己.Top + txt組合_壬癸己.Height)

			iL += iW
			txt大運_傳統法.Size = New Size(iW, i大運) : txt大運_傳統法.Location = New Point(iL, 0)
			txt八字_傳統法.Size = New Size(iW, i八字) : txt八字_傳統法.Location = New Point(iL, txt大運_傳統法.Height)
			txt組合_傳統法.Size = New Size(iW, i組合) : txt組合_傳統法.Location = New Point(iL, txt八字_傳統法.Top + txt八字_傳統法.Height)
			txt流月_傳統法.Size = txt組合_傳統法.Size : txt流月_傳統法.Location = txt組合_傳統法.Location
			txt流年_傳統法.Size = New Size(iW, i流年) : txt流年_傳統法.Location = New Point(iL, txt組合_傳統法.Top + txt組合_傳統法.Height)

			txt大運_六爻法.Size = New Size(iW, i大運) : txt大運_六爻法.Location = New Point(iL, 0)
			txt八字_六爻法.Size = New Size(iW, i八字) : txt八字_六爻法.Location = New Point(iL, txt大運_六爻法.Height)
			txt組合_六爻法.Size = New Size(iW, i組合) : txt組合_六爻法.Location = New Point(iL, txt八字_六爻法.Top + txt八字_六爻法.Height)
			txt流月_六爻法.Size = txt組合_六爻法.Size : txt流月_六爻法.Location = txt組合_六爻法.Location
			txt流年_六爻法.Size = New Size(iW, i流年) : txt流年_六爻法.Location = New Point(iL, txt組合_六爻法.Top + txt組合_六爻法.Height)
		End If
	End Sub
	Private Sub us部署()
		Dim iW, iT, iL As Int32
		Dim sz As Size
		'---------------------------------------------
		iT = 0 : iW = 40 : sz = New Size(iW, 25)
		txtYgz.Size = sz : txtYgz.Location = New Point(0, iT)
		txtMgz.Size = sz : txtMgz.Location = New Point(iW, iT)
		txtDgz.Size = sz : txtDgz.Location = New Point(iW * 2, iT)
		txtHgz.Size = sz : txtHgz.Location = New Point(iW * 3, iT)
		txt月gz.Size = sz : txt月gz.Location = New Point(iW * 4, iT)

		iT += 25
		txt旬空.Location = New Point(0, iT) : txt旬空.Size = New Size(210, 65)
		txt地支組合.Location = New Point(txt旬空.Left + txt旬空.Width, 0) : txt地支組合.Size = New Size(140, 165)
		'---------------------------------------------
		iT += txt旬空.Height
		txt出生年.Size = New Size(45, 25) : txt出生年.Location = New Point(0, iT)
		txt大运起年.Size = New Size(45, 25) : txt大运起年.Location = New Point(txt出生年.Left + txt出生年.Width, iT)
		txt大运起岁.Size = New Size(15, 25) : txt大运起岁.Location = New Point(txt大运起年.Left + txt大运起年.Width, iT)
		chk大運_男順女順.Location = New Point(txt大运起岁.Left + txt大运起岁.Width, iT + 3) : chk大運_男順女順.SendToBack()

		iL = chk大運_男順女順.Left + chk大運_男順女順.Width - 8
		btn排盤.Size = New Size(50, 25) : btn排盤.Location = New Point(iL, iT)
		chk男命.Location = New Point(iL, iT + 27)

		iT += 25
		cbo大運年月.Size = New Size(80, 25) : cbo大運年月.Location = New Point(0, iT) : cbo大運年月.BringToFront()
		txt大運增年.Size = New Size(30, 25) : txt大運增年.Location = New Point(80, iT)
		btn自己八字.Size = New Size(30, 25) : btn自己八字.Location = New Point(110, iT)
		'---------------------------------------------
		iT = 145
		lbl傳統法.Location = New Point(0, iT + 3)
		cbo基準為何支1.Location = New Point(50, iT) : cbo基準為何支1.Size = New Size(45, 25)

		iL = cbo基準為何支1.Left + cbo基準為何支1.Width
		btn四格2記憶體.Size = New Size(60, 25) : btn四格2記憶體.Location = New Point(iL, iT)
		iL = btn四格2記憶體.Left + btn四格2記憶體.Width
		btn記憶體2四格.Size = New Size(60, 25) : btn記憶體2四格.Location = New Point(iL, iT)

		iT += 25 : iW = 60 : sz = New Size(iW, 110)
		txt年柱1.Size = sz : txt年柱1.Location = New Point(0, iT)
		txt月柱1.Size = sz : txt月柱1.Location = New Point(iW, iT)
		txt日柱1.Size = sz : txt日柱1.Location = New Point(iW * 2, iT)
		txt時柱1.Size = sz : txt時柱1.Location = New Point(iW * 3, iT)
		txt統計1.Size = New Size(110, 110) : txt統計1.Location = New Point(iW * 4, iT)
		'---------------------------------------------
		iT += txt年柱1.Height
		lbl干化支.Location = New Point(0, iT + 3)
		cbo基準為何支2.Location = New Point(70, iT) : cbo基準為何支2.Size = New Size(90, 25)
		cbo干藏支.Location = New Point(160, iT) : cbo干藏支.Size = New Size(80, 25)
		cbo支藏干.Location = New Point(240, iT) : cbo支藏干.Size = New Size(100, 25)

		iT += 25
		txt年柱2.Size = sz : txt年柱2.Location = New Point(0, iT)
		txt月柱2.Size = sz : txt月柱2.Location = New Point(iW, iT)
		txt日柱2.Size = sz : txt日柱2.Location = New Point(iW * 2, iT)
		txt時柱2.Size = sz : txt時柱2.Location = New Point(iW * 3, iT)
		txt統計2.Size = New Size(110, 110) : txt統計2.Location = New Point(iW * 4, iT)
		'---------------------------------------------
		iT += 110
		lbl壬癸己1.Location = New Point(0, iT + 3)
		txt轉換後干支1.Location = New Point(70, iT) : txt轉換後干支1.Size = New Size(130, 25)
		cbo基準為何支3.Location = New Point(300, iT) : cbo基準為何支3.Size = New Size(50, 25)

		iT += 25
		txt年柱3.Size = sz : txt年柱3.Location = New Point(0, iT)
		txt月柱3.Size = sz : txt月柱3.Location = New Point(iW, iT)
		txt日柱3.Size = sz : txt日柱3.Location = New Point(iW * 2, iT)
		txt時柱3.Size = sz : txt時柱3.Location = New Point(iW * 3, iT)
		txt統計3.Size = New Size(110, 110) : txt統計3.Location = New Point(iW * 4, iT)
		'---------------------------------------------
		iT += 110
		lbl壬癸己2.Location = New Point(0, iT + 3)
		txt轉換後干支2.Location = New Point(70, iT) : txt轉換後干支2.Size = New Size(130, 25)
		cbo基準為何支4.Location = New Point(200, iT) : cbo基準為何支4.Size = New Size(90, 25)
		lsb流年干.Location = New Point(320, iT) ': lsb流月.Size = New Size(25, 140)
		lsb流年干.SelectedIndex = 0

		iT += 25
		txt年柱4.Size = sz : txt年柱4.Location = New Point(0, iT)
		txt月柱4.Size = sz : txt月柱4.Location = New Point(iW, iT)
		txt日柱4.Size = sz : txt日柱4.Location = New Point(iW * 2, iT)
		txt時柱4.Size = sz : txt時柱4.Location = New Point(iW * 3, iT)
		txt統計4.Size = New Size(110, 110) : txt統計4.Location = New Point(iW * 4, iT)
		'---------------------------------------------
		iT += 110
		lbl八行法.Location = New Point(0, iT + 3)
		cbo基準為何支5.Location = New Point(70, iT) : cbo基準為何支5.Size = New Size(90, 25)

		iT += 25
		txt年柱5.Size = sz : txt年柱5.Location = New Point(0, iT)
		txt月柱5.Size = sz : txt月柱5.Location = New Point(iW, iT)
		txt日柱5.Size = sz : txt日柱5.Location = New Point(iW * 2, iT)
		txt時柱5.Size = sz : txt時柱5.Location = New Point(iW * 3, iT)
		txt統計5.Size = New Size(110, 110) : txt統計5.Location = New Point(iW * 4, iT)
		'---------------------------------------------
		Dim icsH = Me.ClientSize.Height
		Dim i大運 = 130, i八字 = 200, i組合 = 200, i流年 = icsH - i大運 - i八字 - i組合
		iL = txt統計5.Left + txt統計5.Width : iW = _i寬度
		txt大運_壬癸己.Size = New Size(iW, i大運) : txt大運_壬癸己.Location = New Point(iL, 0)
		txt八字_壬癸己.Size = New Size(iW, i八字) : txt八字_壬癸己.Location = New Point(iL, txt大運_壬癸己.Height)
		txt組合_壬癸己.Size = New Size(iW, i組合) : txt組合_壬癸己.Location = New Point(iL, txt八字_壬癸己.Top + txt八字_壬癸己.Height)
		txt流月_壬癸己.Size = txt組合_壬癸己.Size : txt流月_壬癸己.Location = txt組合_壬癸己.Location : txt流月_壬癸己.SendToBack()
		txt流年_壬癸己.Size = New Size(iW, i流年) : txt流年_壬癸己.Location = New Point(iL, txt組合_壬癸己.Top + txt組合_壬癸己.Height)

		iL += iW
		txt大運_傳統法.Size = New Size(iW, i大運) : txt大運_傳統法.Location = New Point(iL, 0)
		txt八字_傳統法.Size = New Size(iW, i八字) : txt八字_傳統法.Location = New Point(iL, txt大運_傳統法.Height)
		txt組合_傳統法.Size = New Size(iW, i組合) : txt組合_傳統法.Location = New Point(iL, txt八字_傳統法.Top + txt八字_傳統法.Height)
		txt流月_傳統法.Size = txt組合_傳統法.Size : txt流月_傳統法.Location = txt組合_傳統法.Location : txt流月_傳統法.SendToBack()
		txt流年_傳統法.Size = New Size(iW, i流年) : txt流年_傳統法.Location = New Point(iL, txt組合_傳統法.Top + txt組合_傳統法.Height)

		txt大運_六爻法.Size = New Size(iW, i大運) : txt大運_六爻法.Location = New Point(iL, 0)
		txt八字_六爻法.Size = New Size(iW, i八字) : txt八字_六爻法.Location = New Point(iL, txt大運_六爻法.Height)
		txt組合_六爻法.Size = New Size(iW, i組合) : txt組合_六爻法.Location = New Point(iL, txt八字_六爻法.Top + txt八字_六爻法.Height)
		txt流月_六爻法.Size = txt組合_六爻法.Size : txt流月_六爻法.Location = txt組合_六爻法.Location : txt流月_六爻法.SendToBack()
		txt流年_六爻法.Size = New Size(iW, i流年) : txt流年_六爻法.Location = New Point(iL, txt組合_六爻法.Top + txt組合_六爻法.Height)

		txt大運_六爻法.Hide()
		txt八字_六爻法.Hide()
		txt組合_六爻法.Hide()
		txt流年_六爻法.Hide()
		txt流月_六爻法.Hide()
		'---------------------------------------------
		cbo干藏支.Items.AddRange({"甲=寅亥", "甲=寅卯"})
		cbo干藏支.DropDownStyle = ComboBoxStyle.DropDownList : cbo干藏支.Text = "甲=寅卯"

		cbo支藏干.Items.AddRange({"寅[甲丙戊]", "寅[甲丙己]", "寅[甲戊]"})
		cbo支藏干.DropDownStyle = ComboBoxStyle.DropDownList : cbo支藏干.Text = "寅[甲丙戊]"

		cbo大運年月.Items.AddRange({"大運=月", "大運=年"})
		cbo大運年月.DropDownStyle = ComboBoxStyle.DropDownList : cbo大運年月.Text = "大運=年"
		'---------------------------
		cbo基準為何支1.Items.AddRange({"Dg", "Dz", "Mg", "Mz", "Yg", "Yz", "Hg"， "Hz"})
		cbo基準為何支1.DropDownStyle = ComboBoxStyle.DropDownList : cbo基準為何支1.Text = "Dg"

		cbo基準為何支2.Items.AddRange({"Dz", "Mz", "Yz"， "Hz", "Dg-主氣", "Dg-中氣"})
		cbo基準為何支2.DropDownStyle = ComboBoxStyle.DropDownList : cbo基準為何支2.Text = "Dg-主氣"
		' 六行-干干
		cbo基準為何支3.Items.AddRange({"Dg", "Dz", "Mg", "Mz", "Yg", "Yz", "Hg"， "Hz", "月g", "月z"})
		cbo基準為何支3.DropDownStyle = ComboBoxStyle.DropDownList : cbo基準為何支3.Text = "Dg"
		' 六行-支支
		cbo基準為何支4.Items.AddRange({"Dz", "Mz", "Yz"， "Hz", "月z", "Dg-主氣", "Dg-中氣", "月g-主氣", "月g-中氣"})
		cbo基準為何支4.DropDownStyle = ComboBoxStyle.DropDownList : cbo基準為何支4.Text = "Dg-主氣"
		' 八行-支支
		cbo基準為何支5.Items.AddRange({"Dz", "Mz", "Yz"， "Hz", "月z", "Dg-主氣", "Dg-中氣", "月g-主氣", "月g-中氣"})
		cbo基準為何支5.DropDownStyle = ComboBoxStyle.DropDownList : cbo基準為何支5.Text = "Dg-主氣"
	End Sub

	Dim _s60甲子擴充_(120) As String
	Private Sub us起始設定()
		' 排大運時，若女命，甲子年生，逆排，但剛好在 __s60甲子Idx_(1)，故要擴充矩陣
		For i = 1 To 60
			_s60甲子擴充_(i) = __s60甲子Idx_(i)
		Next
		For i = 1 To 60
			_s60甲子擴充_(i + 60) = __s60甲子Idx_(i)
		Next
		'---------------------------------------------
		_dic串宮_干2支_古法.Clear() : _dic串宮_干2支_六行.Clear()
		With _dic串宮_干2支_古法
			.Add("甲", "寅") : .Add("乙", "卯") : .Add("丙", "巳") : .Add("丁", "午")
			.Add("戊", "戌")
			.Add("己", "丑")
			.Add("庚", "申") : .Add("辛", "酉") : .Add("壬", "亥") : .Add("癸", "子")
		End With

		With _dic串宮_干2支_六行
			.Add("甲", "寅") : .Add("乙", "卯") : .Add("丙", "巳") : .Add("丁", "午")
			.Add("戊", "未") ' 变
			.Add("己", "丑")
			.Add("庚", "申") : .Add("辛", "酉") : .Add("壬", "亥") : .Add("癸", "子")
		End With
	End Sub
	'================================================================
	Private Sub dgv60甲子_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv60甲子.CellClick
		Dim s1 = dgv60甲子.CurrentCell.Value?.ToString
		dgv60甲子.Hide()
		Select Case dgv60甲子.Tag?.ToString
			Case txtYgz.Name : txtYgz.Focus() : txtYgz.Text = s1
			Case txtMgz.Name : txtMgz.Focus() : txtMgz.Text = s1
			Case txtDgz.Name : txtDgz.Focus() : txtDgz.Text = s1
			Case txtHgz.Name : txtHgz.Focus() : txtHgz.Text = s1
			Case txt月gz.Name : txt月gz.Focus() : txt月gz.Text = s1
		End Select
	End Sub
	Private Sub dgv60甲子_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs)
		If e.KeyCode = Keys.Escape Then dgv60甲子.Hide()
	End Sub
	'================================================================
	Public Sub us來自八字案例(s八字 As String, is男 As Boolean, i出生年 As Int32, i大运起年 As Int32, i大运起岁 As Int32, s月宮 As String)
		chk男命.Checked = is男
		txt出生年.Text = i出生年.ToString
		txt大运起年.Text = i大运起年.ToString
		txt大运起岁.Text = i大运起岁.ToString
		'------------------------------------------------
		If s月宮.Equals("○") Then s月宮 = "子"
		Select Case s月宮
			Case "子", "寅", "辰", "午", "申", "戌" : s月宮 = " 甲" + s月宮
			Case "丑", "卯", "巳", "未", "酉", "亥" : s月宮 = " 乙" + s月宮
		End Select

		s八字 += s月宮
		txtYgz.Text = Mid(s八字, 1, 2)
		txtMgz.Text = Mid(s八字, 4, 2)
		txtDgz.Text = Mid(s八字, 7, 2)
		txtHgz.Text = Mid(s八字, 10, 2)
		txt月gz.Text = Mid(s八字, 13, 2)

		btn排盤.PerformClick()
	End Sub
	Private Function uf記憶體取代字(s As String) As String
		s = s.Replace("年", "").Replace("月", "").Replace("日", "").Replace("时", "").Replace("時", "")
		Select Case True
			Case s.Contains("乾："), s.Contains("男：") : chk男命.Checked = True
			Case s.Contains("坤："), s.Contains("女：") : chk男命.Checked = False
		End Select
		s = s.Replace("乾：", "").Replace("坤：", "").Replace("男：", "").Replace("女：", "")
		s = s.Replace("ＯＯ", "甲子").Replace("、", " ").Replace("　", " ")
		s = s.Replace("y", " ").Replace("m", " ").Replace("d", " ").Replace("h", "") ' 台湾921大地震：己卯y癸酉m丙子d己丑h
		Return s
	End Function
	Private Sub btn記憶體2四格_Click(sender As Object, e As EventArgs) Handles btn記憶體2四格.Click
		Dim s = uf記憶體取代字（Clipboard.GetText.Trim）
		If s.Length = 11 Then s += " 甲子"
		If s.Length <> 14 Then MsgBox("字串長度不符規定=" + s) : Exit Sub

		_sbYgz = Mid(s, 1, 2) : If Not __dic60納音.ContainsKey(_sbYgz) Then MsgBox(_sbYgz + " 找不到此干支") : Exit Sub
		_sbMgz = Mid(s, 4, 2) : If Not __dic60納音.ContainsKey(_sbMgz) Then MsgBox(_sbMgz + " 找不到此干支") : Exit Sub
		_sbDgz = Mid(s, 7, 2) : If Not __dic60納音.ContainsKey(_sbDgz) Then MsgBox(_sbDgz + " 找不到此干支") : Exit Sub
		_sbHgz = Mid(s, 10, 2) : If Not __dic60納音.ContainsKey(_sbHgz) Then MsgBox(_sbHgz + " 找不到此干支") : Exit Sub
		_sb月亮gz = Mid(s, 13, 2) : If Not __dic60納音.ContainsKey(_sb月亮gz) Then MsgBox(_sb月亮gz + " 找不到此干支") : Exit Sub
		txtYgz.Text = _sbYgz
		txtMgz.Text = _sbMgz
		txtDgz.Text = _sbDgz
		txtHgz.Text = _sbHgz
		txt月gz.Text = _sb月亮gz
		'------------------------------------------------
		' 粗估 大运起年
		Dim iYnow = Today.Year
		For iY = iYnow - 80 To iYnow - 20 ' 20岁內的年青人，不会想算命
			If __dic西元年_60甲子(iY).Equals(txtYgz.Text) Then
				txt出生年.Text = iY.ToString
				txt大运起年.Text = (iY + 1).ToString
				txt大运起岁.Text = "1"
			End If
		Next
		'------------------------------------------------
		btn排盤.PerformClick()
	End Sub
	Private Sub btn四格2記憶體_Click(sender As Object, e As EventArgs) Handles btn四格2記憶體.Click
		' 方便copy到資料庫中
		Clipboard.SetText(txtYgz.Text + " " + txtMgz.Text + " " + txtDgz.Text + " " + txtHgz.Text)
	End Sub
	Private Sub btn自己八字_Click(sender As Object, e As EventArgs) Handles btn自己八字.Click
		txtYgz.Text = "戊申" : txtMgz.Text = "庚申" : txtDgz.Text = "丙子" : txtHgz.Text = "辛卯" : txt月gz.Text = "甲子"
		chk男命.Checked = True : txt出生年.Text = "1968" : txt大运起年.Text = "1970" : txt大运起岁.Text = "2" : txt出生年.Text = "1968"
		btn排盤.PerformClick()
	End Sub

	Dim _i大运起年, _i大运起岁 As Int32, _i出生年 As Int32 = 1968 ' 我的生年
	Dim _s大運_傳統法, _s基準_傳統法 As String, _is男 As Boolean = False
	Dim _s大運_壬癸己, _s基準_壬癸己 As String
	Dim _sbYgz, _sbYg, _sbYz, _sYo As String
	Dim _sbMgz, _sbMg, _sbMz, _sMo As String
	Dim _sbDgz, _sbDg, _sbDz, _sDo As String
	Dim _sbHgz, _sbHg, _sbHz, _sHo As String
	Dim _sb月亮gz, _sb月亮g, _sb月亮z, _s月亮o As String
	Dim _sb胎元gz, _sb胎元g, _sb胎元z, _s胎元o As String
	Dim _sb命宮gz, _sb命宮g, _sb命宮z, _s命宮o As String
	Dim _sb身宮gz, _sb身宮g, _sb身宮z, _s身宮o As String
	Private Sub btn排盤_Click(sender As Object, e As EventArgs) Handles btn排盤.Click
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		_sbYgz = txtYgz.Text : _sbYg = Mid(_sbYgz, 1, 1) : _sbYz = Mid(_sbYgz, 2, 1)
		_sbMgz = txtMgz.Text : _sbMg = Mid(_sbMgz, 1, 1) : _sbMz = Mid(_sbMgz, 2, 1)
		_sbDgz = txtDgz.Text : _sbDg = Mid(_sbDgz, 1, 1) : _sbDz = Mid(_sbDgz, 2, 1)
		_sbHgz = txtHgz.Text : _sbHg = Mid(_sbHgz, 1, 1) : _sbHz = Mid(_sbHgz, 2, 1)
		'------------------------------------------------
		_sb月亮gz = txt月gz.Text : _sb月亮g = Mid(_sb月亮gz, 1, 1) : _sb月亮z = Mid(_sb月亮gz, 2, 1)
		_sb胎元gz = uf胎元(_sbMgz) : _sb胎元g = Mid(_sb胎元gz, 1, 1) : _sb胎元z = Mid(_sb胎元gz, 2, 1)
		_sb命宮gz = uf命宮(_sbYg, _sbMz, _sbHz) : _sb命宮g = Mid(_sb命宮gz, 1, 1) : _sb命宮z = Mid(_sb命宮gz, 2, 1)
		_sb身宮gz = uf身宮(_sbYg, _sbMz, _sbHz) : _sb身宮g = Mid(_sb身宮gz, 1, 1) : _sb身宮z = Mid(_sb身宮gz, 2, 1)
		'------------------------------------------------
		Dim sY空亡 = "", sM空亡 = "", sD空亡 = "", sH空亡 = "", s月空亡 = ""
		Dim sY出空 = "", sM出空 = "", sD出空 = "", sH出空 = "", s月出空 = ""

		m易理.us_空亡(_sbYgz, sY空亡, sY出空) : m易理.us_空亡(_sbMgz, sM空亡, sM出空)
		m易理.us_空亡(_sbDgz, sD空亡, sD出空) : m易理.us_空亡(_sbHgz, sH空亡, sH出空) : m易理.us_空亡(_sb月亮gz, s月空亡, s月出空)

		txt旬空.Text = "生占：" + _sbYgz + " " + _sbMgz + " " + _sbDgz + " " + _sbHgz + " " + _sb月亮gz + vbNL +
									"空亡：" + sY空亡 + " " + sM空亡 + " " + sD空亡 + " " + sH空亡 + " " + s月空亡 + vbNL +
									"出空：" + sY出空 + " " + sM出空 + " " + sD出空 + " " + sH出空 + " " + s月出空
		'------------------------------------------------
		_sYo = "" : If InStr(sY空亡, _sbYz) > 0 Then : _sYo = "Yo" : ElseIf InStr(sD空亡, _sbYz) > 0 Then : _sYo = "Do" : ElseIf InStr(sM空亡, _sbYz) > 0 Then : _sYo = "Mo" : End If
		_sMo = "" : If InStr(sY空亡, _sbMz) > 0 Then : _sMo = "Yo" : ElseIf InStr(sD空亡, _sbMz) > 0 Then : _sMo = "Do" : ElseIf InStr(sM空亡, _sbMz) > 0 Then : _sMo = "Mo" : End If
		_sDo = "" : If InStr(sY空亡, _sbDz) > 0 Then : _sDo = "Yo" : ElseIf InStr(sD空亡, _sbDz) > 0 Then : _sDo = "Do" : ElseIf InStr(sM空亡, _sbDz) > 0 Then : _sDo = "Mo" : End If
		_sHo = "" : If InStr(sY空亡, _sbHz) > 0 Then : _sHo = "Yo" : ElseIf InStr(sD空亡, _sbHz) > 0 Then : _sHo = "Do" : ElseIf InStr(sM空亡, _sbHz) > 0 Then : _sHo = "Mo" : End If

		_s胎元o = "" : If InStr(sY空亡, _sb胎元z) > 0 Then : _s胎元o = "Yo" : ElseIf InStr(sD空亡, _sb胎元z) > 0 Then : _s胎元o = "Do" : ElseIf InStr(sM空亡, _sb胎元z) > 0 Then : _s胎元o = "Mo" : End If
		_s命宮o = "" : If InStr(sY空亡, _sb命宮z) > 0 Then : _s命宮o = "Yo" : ElseIf InStr(sD空亡, _sb命宮z) > 0 Then : _s命宮o = "Do" : ElseIf InStr(sM空亡, _sb命宮z) > 0 Then : _s命宮o = "Mo" : End If
		_s身宮o = "" : If InStr(sY空亡, _sb身宮z) > 0 Then : _s身宮o = "Yo" : ElseIf InStr(sD空亡, _sb身宮z) > 0 Then : _s身宮o = "Do" : ElseIf InStr(sM空亡, _sb身宮z) > 0 Then : _s身宮o = "Mo" : End If
		'------------------------------------------------
		_i出生年 = CInt(txt出生年.Text)
		_i大运起年 = CInt(txt大运起年.Text)
		_i大运起岁 = CInt(txt大运起岁.Text)
		_is男 = chk男命.Checked
		us串宮压运(_sbYz)
		us方法_傳統法_干() : us方法_傳統法_支()
		us方法_壬癸己_干() : us方法_壬癸己_支() : us方法_壬癸己_六爻()
		us方法_八行法_支()
		us算流月()
	End Sub
	Private Function uf單柱橫排(s柱子 As String, sZ As String) As String
		Dim s = "", s_() As String
		s_ = Split(s柱子, vbNL)
		s = s_(0) : s_(0) = Mid(s, 2, 1) + Mid(s, 1, 1) ' 甲财->财甲 天干要透出
		Select Case s_.Length
			Case 2 : Return s_(0) + " " + sZ + "[" + s_(1) + "]"
			Case 3 : Return s_(0) + " " + sZ + "[" + s_(1) + "+" + s_(2) + "]"
			Case 4 : Return s_(0) + " " + sZ + "[" + s_(1) + "+" + s_(2) + "+" + s_(3) + "]"
		End Select
		Return s
	End Function
	Private Sub us方法_傳統法_干()
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		Dim s = "", sA = "", sYMDHgg, sYMDHgz, s_() As String
		Dim sYgg, sYY As String
		Dim sMgg, sMM As String
		Dim sDgg, sDD As String
		Dim sHgg, sHH As String
		Dim s胎元gg, s胎元 As String
		Dim s命宮gg, s命宮 As String
		Dim s身宮gg, s身宮 As String
		sYgg = _sbYg + __dic支藏干_古法(_sbYz)
		sMgg = _sbMg + __dic支藏干_古法(_sbMz)
		sDgg = _sbDg + __dic支藏干_古法(_sbDz)
		sHgg = _sbHg + __dic支藏干_古法(_sbHz)
		s胎元gg = _sb胎元g + __dic支藏干_古法(_sb胎元z)
		s命宮gg = _sb命宮g + __dic支藏干_古法(_sb命宮z)
		s身宮gg = _sb身宮g + __dic支藏干_古法(_sb身宮z)
		sYMDHgg = sYgg + sMgg + sDgg + sHgg
		sYMDHgz = _sbYgz + _sbMgz + _sbDgz + _sbHgz

		Select Case cbo基準為何支1.Text
			Case "Yg" : sA = _sbYg : Case "Yz" : sA = Mid(__dic支藏干_古法(_sbYz), 1, 1)
			Case "Mg" : sA = _sbMg : Case "Mz" : sA = Mid(__dic支藏干_古法(_sbMz), 1, 1)
			Case "Dg" : sA = _sbDg : Case "Dz" : sA = Mid(__dic支藏干_古法(_sbDz), 1, 1)
			Case "Hg" : sA = _sbHg : Case "Hz" : sA = Mid(__dic支藏干_古法(_sbHz), 1, 1)
		End Select
		_s基準_傳統法 = sA

		s = "" : For Each s1 As Char In sYgg : s += s1 + uf_印綬_干vs干_五行(sA, s1) + vbNL : Next : sYY = s.TrimEnd : txt年柱1.Text = sYY
		s = "" : For Each s1 As Char In sMgg : s += s1 + uf_印綬_干vs干_五行(sA, s1) + vbNL : Next : sMM = s.TrimEnd : txt月柱1.Text = sMM
		s = "" : For Each s1 As Char In sDgg : s += s1 + uf_印綬_干vs干_五行(sA, s1) + vbNL : Next : sDD = s.TrimEnd : txt日柱1.Text = sDD
		s = "" : For Each s1 As Char In sHgg : s += s1 + uf_印綬_干vs干_五行(sA, s1) + vbNL : Next : sHH = s.TrimEnd : txt時柱1.Text = sHH

		s = "" : For Each s1 As Char In s胎元gg : s += s1 + uf_印綬_干vs干_五行(sA, s1) + vbNL : Next : s胎元 = s.TrimEnd
		s = "" : For Each s1 As Char In s命宮gg : s += s1 + uf_印綬_干vs干_五行(sA, s1) + vbNL : Next : s命宮 = s.TrimEnd
		s = "" : For Each s1 As Char In s身宮gg : s += s1 + uf_印綬_干vs干_五行(sA, s1) + vbNL : Next : s身宮 = s.TrimEnd
		'---------------------------------------------
		Dim sY八卦 = "(" + uf八卦_古法(_sbYgz) + ")", sY長生 = uf_12長生_干支_五行(_sbYg, _sbYz) + " ", sY弔字 = uf單柱引字_五行(_sbYgz)
		Dim sM八卦 = "(" + uf八卦_古法(_sbMgz) + ")", sM長生 = uf_12長生_干支_五行(_sbMg, _sbMz) + " ", sM弔字 = uf單柱引字_五行(_sbMgz)
		Dim sD八卦 = "(" + uf八卦_古法(_sbDgz) + ")", sD長生 = uf_12長生_干支_五行(_sbDg, _sbDz) + " ", sD弔字 = uf單柱引字_五行(_sbDgz)
		Dim sH八卦 = "(" + uf八卦_古法(_sbHgz) + ")", sH長生 = uf_12長生_干支_五行(_sbHg, _sbHz) + " ", sH弔字 = uf單柱引字_五行(_sbHgz)
		Dim s胎元八卦 = "(" + uf八卦_古法(_sb胎元gz) + ")", s胎元長生 = uf_12長生_干支_五行(_sb胎元g, _sb胎元z) + " ", s胎元弔字 = uf單柱引字_五行(_sb胎元gz)
		Dim s命宮八卦 = "(" + uf八卦_古法(_sb命宮gz) + ")", s命宮長生 = uf_12長生_干支_五行(_sb命宮g, _sb命宮z) + " ", s命宮弔字 = uf單柱引字_五行(_sb命宮gz)
		Dim s身宮八卦 = "(" + uf八卦_古法(_sb身宮gz) + ")", s身宮長生 = uf_12長生_干支_五行(_sb身宮g, _sb身宮z) + " ", s身宮弔字 = uf單柱引字_五行(_sb身宮gz)

		Dim sY串宮 = _dic串宮_12神(_dic串宮_干2支_古法(_sbYg)) + "+" + _dic串宮_12神(_sbYz) + "|"
		Dim sM串宮 = _dic串宮_12神(_dic串宮_干2支_古法(_sbMg)) + "+" + _dic串宮_12神(_sbMz) + "|"
		Dim sD串宮 = _dic串宮_12神(_dic串宮_干2支_古法(_sbDg)) + "+" + _dic串宮_12神(_sbDz) + "|"
		Dim sH串宮 = _dic串宮_12神(_dic串宮_干2支_古法(_sbHg)) + "+" + _dic串宮_12神(_sbHz) + "|"
		Dim s胎元串宮 = _dic串宮_12神(_dic串宮_干2支_古法(_sb胎元g)) + "+" + _dic串宮_12神(_sb胎元z) + "|"
		Dim s命宮串宮 = _dic串宮_12神(_dic串宮_干2支_古法(_sb命宮g)) + "+" + _dic串宮_12神(_sb命宮z) + "|"
		Dim s身宮串宮 = _dic串宮_12神(_dic串宮_干2支_古法(_sb身宮g)) + "+" + _dic串宮_12神(_sb身宮z) + "|"

		sYY = sY八卦 + uf單柱橫排(sYY, _sbYz) + sY長生 + sY弔字 + _sYo
		sMM = sM八卦 + uf單柱橫排(sMM, _sbMz) + sM長生 + sM弔字 + _sMo
		sDD = sD八卦 + uf單柱橫排(sDD, _sbDz) + sD長生 + sD弔字 + _sDo
		sHH = sH八卦 + uf單柱橫排(sHH, _sbHz) + sH長生 + sH弔字 + _sHo

		s胎元 = s胎元八卦 + s胎元串宮 + uf單柱橫排(s胎元, _sb胎元z) + s胎元長生 + s胎元弔字 + _s胎元o
		s命宮 = s命宮八卦 + s命宮串宮 + uf單柱橫排(s命宮, _sb命宮z) + s命宮長生 + s命宮弔字 + _s命宮o
		s身宮 = s身宮八卦 + s身宮串宮 + uf單柱橫排(s身宮, _sb身宮z) + s身宮長生 + s身宮弔字 + _s身宮o
		txt八字_傳統法.Text = sYY + vbNL + sMM + vbNL + sDD + vbNL + sHH + vbNL + vbNL + s胎元 + vbNL + s命宮 + vbNL + s身宮
		'---------------------------------------------
		s = ""
		s += uf命局缺的字_干支_五行(sA, sYY + sMM + sDD + sHH)
		s += "Y" + _sbYgz + "->" + uf干支巔倒_五行(_sbYgz) + "，" + "M" + _sbMgz + "->" + uf干支巔倒_五行(_sbMgz) + "，" +
				 "D" + _sbDgz + "->" + uf干支巔倒_五行(_sbDgz) + "，" + "H" + _sbHgz + "->" + uf干支巔倒_五行(_sbHgz) + vbNL
		s += Uf八字組合_五行(sYMDHgz)
		s += uf天地沖合_五行(sYMDHgz)
		s += uf_MzHz組和_五行(_sbMz, _sbHz) + vbNL
		s += ufDD組合_五行("Y", _sbYgz)
		s += ufDD組合_五行("M", _sbMgz)
		s += ufDD組合_五行("D", _sbDgz)
		s += ufDD組合_五行("H", _sbHgz)
		s += ufHH組合_五行("H", _sbHgz)
		s += ufDH組合_五行(_sbDgz, _sbHgz)
		s += "D" + _sbDg + "@M" + _sbMz + "=" + uf_12長生_干支_五行(_sbDg, _sbMz) + vbTab
		s += "D" + _sbDg + "@M" + _sbMz + "=" + uf天乙貴人(_sbDg, _sbMz)
		txt組合_傳統法.Text = s.TrimEnd  ' 去掉最后的vbNL

		s = ""
		s = uf地支組合_五行(_sbYz + _sbMz + _sbDz + _sbHz)
		's += "胎元：" + uf胎元(_sbMgz) + vbNL
		's += "命宮：" + uf命宮(_sbYg, _sbMz, _sbHz) + vbNL
		's += "身宮：" + uf身宮(_sbYg, _sbMz, _sbHz) + vbNL
		s += _sbYgz + "=" + __dic60納音(_sbYgz) + "," + uf_12長生_干支_五行(_sbYg, _sbYz) + vbNL
		s += _sbMgz + "=" + __dic60納音(_sbMgz) + "," + uf_12長生_干支_五行(_sbMg, _sbMz) + vbNL
		s += _sbDgz + "=" + __dic60納音(_sbDgz) + "," + uf_12長生_干支_五行(_sbDg, _sbDz) + vbNL
		s += _sbHgz + "=" + __dic60納音(_sbHgz) + "," + uf_12長生_干支_五行(_sbHg, _sbHz) + vbNL
		txt地支組合.Text = s.TrimEnd
		'---------------------------------------------
		us排大運_月柱()
		us算大運_傳統法()
		us算流年_傳統法()
		txt統計1.Text = us統計五行個數_天干(sYMDHgg, sYMDHgz)
	End Sub
	Private Sub us方法_傳統法_支()
		Dim sY = "", sM = "", sD = "", sH = "", s = "", sA = ""
		Dim dic支藏干 As Dictionary(Of String, String) = Nothing

		Select Case cbo干藏支.Text
			Case "甲=寅亥" : dic支藏干 = __dic干藏支_甲_寅亥
			Case "甲=寅卯" : dic支藏干 = __dic干藏支_甲_寅卯_乙_卯未
		End Select

		sY = dic支藏干(_sbYg) + _sbYz
		sM = dic支藏干(_sbMg) + _sbMz
		sD = dic支藏干(_sbDg) + _sbDz
		sH = dic支藏干(_sbHg) + _sbHz
		'---------------------------------------------
		Select Case cbo基準為何支2.Text
			Case "Yz" : sA = _sbYz
			Case "Mz" : sA = _sbMz
			Case "Dz" : sA = _sbDz
			Case "Hz" : sA = _sbHz
			Case "Dg-主氣" : sA = Mid(sD, 1, 1)
			Case "Dg-中氣" : sA = Mid(sD, 2, 1)
		End Select

		s = "" : For Each s1 As Char In sY : s += s1 + uf六親_支vs支_六行_辰未元(sA, s1) + vbNL : Next : txt年柱2.Text = s.TrimEnd 'us父2梟印(s)
		s = "" : For Each s1 As Char In sM : s += s1 + uf六親_支vs支_六行_辰未元(sA, s1) + vbNL : Next : txt月柱2.Text = s.TrimEnd 'us父2梟印(s)
		s = "" : For Each s1 As Char In sD : s += s1 + uf六親_支vs支_六行_辰未元(sA, s1) + vbNL : Next : txt日柱2.Text = s.TrimEnd 'us父2梟印(s)
		s = "" : For Each s1 As Char In sH : s += s1 + uf六親_支vs支_六行_辰未元(sA, s1) + vbNL : Next : txt時柱2.Text = s.TrimEnd 'us父2梟印(s)
		txt統計2.Text = us統計五行個目_地支_六行_辰未元(sY + sM + sD + sH)
	End Sub
	Private Sub us方法_壬癸己_干()
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		Dim sYgg = "", sYg, sYz, sYgz As String, s = "", sA = "", sYMDHgg = "", sYMDHgz = ""
		Dim sMgg = "", sMg, sMz, sMgz As String
		Dim sDgg = "", sDg, sDz, sDgz As String
		Dim sHgg = "", sHg, sHz, sHgz As String
		Dim s月亮gg = "", s月亮 = "", s月亮g, s月亮z, s月亮gz As String
		Dim s胎元gg = "", s胎元 = "", s胎元g, s胎元z, s胎元gz As String
		Dim s命宮gg = "", s命宮 = "", s命宮g, s命宮z, s命宮gz As String
		Dim s身宮gg = "", s身宮 = "", s身宮g, s身宮z, s身宮gz As String
		Dim dic支藏干 As Dictionary(Of String, String) = Nothing

		sYg = uf60甲子轉換_壬癸己(_sbYg) : sYgz = sYg + _sbYz
		sMg = uf60甲子轉換_壬癸己(_sbMg) : sMgz = sMg + _sbMz
		sDg = uf60甲子轉換_壬癸己(_sbDg) : sDgz = sDg + _sbDz
		sHg = uf60甲子轉換_壬癸己(_sbHg) : sHgz = sHg + _sbHz
		s月亮g = uf60甲子轉換_壬癸己(_sb月亮g) : s月亮gz = s月亮g + _sb月亮z
		s胎元g = uf60甲子轉換_壬癸己(_sb胎元g) : s胎元gz = s胎元g + _sb胎元z
		s命宮g = uf60甲子轉換_壬癸己(_sb命宮g) : s命宮gz = s命宮g + _sb命宮z
		s身宮g = uf60甲子轉換_壬癸己(_sb身宮g) : s身宮gz = s身宮g + _sb身宮z

		Select Case cbo支藏干.Text
			Case "寅[甲丙戊]" : dic支藏干 = __dic支藏干_六行_寅_甲丙戊
			Case "寅[甲丙己]" : dic支藏干 = __dic支藏干_六行_寅_甲丙己
			Case "寅[甲戊]" : dic支藏干 = __dic支藏干_六行_寅_甲戊
		End Select

		sYgg = sYg + dic支藏干(_sbYz)
		sMgg = sMg + dic支藏干(_sbMz)
		sDgg = sDg + dic支藏干(_sbDz)
		sHgg = sHg + dic支藏干(_sbHz)
		s月亮gg = s月亮g + dic支藏干(_sb月亮z)
		s胎元gg = s胎元g + dic支藏干(_sb胎元z)
		s命宮gg = s命宮g + dic支藏干(_sb命宮z)
		s身宮gg = s身宮g + dic支藏干(_sb身宮z)

		sYMDHgg = sYgg + sMgg + sDgg + sHgg
		sYMDHgz = sYgz + sMgz + sDgz + sHgz '+ s月gz
		txt轉換後干支1.Text = sYgz + " " + sMgz + " " + sDgz + " " + sHgz '+ " " + s月gz
		txt轉換後干支2.Text = sYgz + " " + sMgz + " " + sDgz + " " + sHgz '+ " " + s月gz
		'---------------------------------------------
		Select Case cbo基準為何支3.Text
			Case "Yg" : sA = sYg : Case "Yz" : sA = Mid(sYgg, 2, 1)
			Case "Mg" : sA = sMg : Case "Mz" : sA = Mid(sMgg, 2, 1)
			Case "Dg" : sA = sDg : Case "Dz" : sA = Mid(sDgg, 2, 1)
			Case "Hg" : sA = sHg : Case "Hz" : sA = Mid(sHgg, 2, 1)
			Case "月g" : sA = s月亮g : Case "月z" : sA = Mid(s月亮gg, 2, 1)
		End Select
		_s基準_壬癸己 = sA

		s = "" : For Each c As Char In sYgg : s += c + uf_印綬_干vs干_六行_奇隅(sA, c) + vbNL : Next : txt年柱3.Text = s.TrimEnd
		s = "" : For Each c As Char In sMgg : s += c + uf_印綬_干vs干_六行_奇隅(sA, c) + vbNL : Next : txt月柱3.Text = s.TrimEnd
		s = "" : For Each c As Char In sDgg : s += c + uf_印綬_干vs干_六行_奇隅(sA, c) + vbNL : Next : txt日柱3.Text = s.TrimEnd
		s = "" : For Each c As Char In sHgg : s += c + uf_印綬_干vs干_六行_奇隅(sA, c) + vbNL : Next : txt時柱3.Text = s.TrimEnd
		's = "" : For Each s1 As Char In s月gg : s += c + uf_印綬_干vs干_六行_奇隅(sA, c) + vbNL : Next : s月亮 = s.TrimEnd
		s = "" : For Each c As Char In s胎元gg : s += c + uf_印綬_干vs干_六行_奇隅(sA, c) + vbNL : Next : s胎元 = s.TrimEnd
		s = "" : For Each c As Char In s命宮gg : s += c + uf_印綬_干vs干_六行_奇隅(sA, c) + vbNL : Next : s命宮 = s.TrimEnd
		s = "" : For Each c As Char In s身宮gg : s += c + uf_印綬_干vs干_六行_奇隅(sA, c) + vbNL : Next : s身宮 = s.TrimEnd
		'---------------------------------------------
		Dim sYY = txt年柱3.Text
		Dim sMM = txt月柱3.Text
		Dim sDD = txt日柱3.Text
		Dim sHH = txt時柱3.Text

		sYY = uf單柱橫排(sYY, _sbYz) + uf_12長生_干支_六行(sYgz) + " " + uf單柱引字_六行(sYgz) + _sYo + vbNL
		sMM = uf單柱橫排(sMM, _sbMz) + uf_12長生_干支_六行(sMgz) + " " + uf單柱引字_六行(sMgz) + _sMo + vbNL
		sDD = uf單柱橫排(sDD, _sbDz) + uf_12長生_干支_六行(sDgz) + " " + uf單柱引字_六行(sDgz) + _sDo + vbNL
		sHH = uf單柱橫排(sHH, _sbHz) + uf_12長生_干支_六行(sHgz) + " " + uf單柱引字_六行(sHgz) + _sHo + vbNL
		's月亮 = uf單柱橫排(s月亮, _sb月z) + uf_12長生_干支_六行(s月亮gz) + " " + uf單柱引字_六行(s月亮gz)
		s胎元 = uf單柱橫排(s胎元, _sb胎元z) + uf_12長生_干支_六行(s胎元gz) + " " + uf單柱引字_六行(s胎元gz) + _s胎元o + vbNL
		s命宮 = uf單柱橫排(s命宮, _sb命宮z) + uf_12長生_干支_六行(s命宮gz) + " " + uf單柱引字_六行(s命宮gz) + _s命宮o + vbNL
		s身宮 = uf單柱橫排(s身宮, _sb身宮z) + uf_12長生_干支_六行(s身宮gz) + " " + uf單柱引字_六行(s身宮gz) + _s身宮o
		'---------------------------------------------
		Dim s1 = "", s2 = "", s3 = "", s4 = "", s5 = "", s6 = ""

		Select Case sDg ' 龙(甲乙)->雀(丙丁)->勾(戊)->虎(庚辛)->玄(壬癸)->蛇(己)
			Case "甲", "乙" : s1 = "龙" : s2 = "雀" : s3 = "勾" : s4 = "虎" : s5 = "玄" : s6 = "蛇"
			Case "丙", "丁" : s1 = "雀" : s2 = "勾" : s3 = "虎" : s4 = "玄" : s5 = "蛇" : s6 = "龙"
			Case "戊" : s1 = "勾" : s2 = "虎" : s3 = "玄" : s4 = "蛇" : s5 = "龙" : s6 = "雀"
			Case "庚", "辛" : s1 = "虎" : s2 = "玄" : s3 = "蛇" : s4 = "龙" : s5 = "雀" : s6 = "勾"
			Case "壬", "癸" : s1 = "玄" : s2 = "蛇" : s3 = "龙" : s4 = "雀" : s5 = "勾" : s6 = "虎"
			Case "己" : s1 = "蛇" : s2 = "龙" : s3 = "雀" : s4 = "勾" : s5 = "虎" : s6 = "玄"
		End Select
		s1 += " " : s2 += " " : s3 += " " : s4 += " " : s5 += " " : s6 += " "

		Select Case True
			Case _is男, chk大運_男順女順.Checked : s = s6 + vbNL + s5 + vbNL + s4 + sYY + s3 + sMM + s2 + sDD + s1 + sHH
			Case Else : s = s1 + vbNL + s2 + vbNL + s3 + sYY + s4 + sMM + s5 + sDD + s6 + sHH
		End Select

		txt八字_壬癸己.Text = s + vbNL + "胎_" + s胎元 + "命_" + s命宮 + "身_" + s身宮
		'---------------------------------------------
		s = ""
		s += uf命局缺的字_干支_六行(sA, sYY + sMM + sDD + sHH) ' + s月柱
		s += "Y" + sYgz + "->" + us干支巔倒(sYgz, sYMDHgz) + "，" +
				 "M" + sMgz + "->" + us干支巔倒(sMgz, sYMDHgz) + "，" +
				 "D" + sDgz + "->" + us干支巔倒(sDgz, sYMDHgz) + "，" +
				 "H" + sHgz + "->" + us干支巔倒(sHgz, sYMDHgz) + vbNL
		s += Uf八字組合_六行(sYMDHgz)
		s += uf天地沖合_六行(sYMDHgz) + vbNL
		s += "D" + sDg + "@M" + _sbMz + "=" + uf_12長生_干支_六行(sDg + _sbMz) + vbTab
		s += "D" + sDg + "@M" + _sbMz + "=" + uf天乙貴人(sDg, _sbMz)
		txt組合_壬癸己.Text = s
		'---------------------------------------------
		If cbo大運年月.Text.Equals("大運=月") Then _s大運_壬癸己 = _s大運_傳統法 Else us排大運_年柱()
		us算大運_壬癸己()
		us算流年_壬癸己()
		txt統計3.Text = us統計六行個數_天干(sYMDHgg, sYMDHgz)
	End Sub
	Private Function us干支巔倒(sGZ As String, sYMDHgz As String) As String
		Dim s0 = "", s1 = ""

		s0 = uf干支巔倒_六行(sGZ)
		If s0.Length = 2 Then
			If sYMDHgz.Contains(s0) Then s0 += "*" ' H甲辰->己寅  2個字
		Else
			s1 = Mid(s0, 4, 2) ' Y庚亥->壬申?壬酉，5個字 =>壬酉
			If sYMDHgz.Contains(s1) Then s0 += "*"
		End If

		Return s0
	End Function

	Private Sub us方法_壬癸己_支()
		Dim sYzz = "", sYg, sYgz As String, s = "", sA = "", sYMDHzz = ""
		Dim sMzz = "", sMg, sMgz As String
		Dim sDzz = "", sDg, sDgz As String
		Dim sHzz = "", sHg, sHgz As String
		Dim s月亮zz = "", s月亮gg = "", s月亮g, s月亮gz As String
		Dim dic干藏支 As Dictionary(Of String, String) = Nothing

		sYg = uf60甲子轉換_壬癸己(_sbYg) : sYgz = sYg + _sbYz
		sMg = uf60甲子轉換_壬癸己(_sbMg) : sMgz = sMg + _sbMz
		sDg = uf60甲子轉換_壬癸己(_sbDg) : sDgz = sDg + _sbDz
		sHg = uf60甲子轉換_壬癸己(_sbHg) : sHgz = sHg + _sbHz
		s月亮g = uf60甲子轉換_壬癸己(_sb月亮g) : s月亮gz = s月亮g + _sb月亮z : s月亮gg = s月亮g + __dic支藏干_六行_寅_甲丙己(_sb月亮z)
		'---------------------------------------------
		Select Case cbo干藏支.Text
			Case "甲=寅亥" : dic干藏支 = __dic干藏支_甲_寅亥
			Case "甲=寅卯" : dic干藏支 = __dic干藏支_甲_寅卯_乙_卯未
		End Select

		sYzz = dic干藏支(sYg) + _sbYz
		sMzz = dic干藏支(sMg) + _sbMz
		sDzz = dic干藏支(sDg) + _sbDz
		sHzz = dic干藏支(sHg) + _sbHz
		s月亮zz = dic干藏支(s月亮g) + _sb月亮z
		sYMDHzz = sYzz + sMzz + sDzz + sHzz
		'---------------------------------------------
		Select Case cbo基準為何支4.Text
			Case "Yz" : sA = _sbYz
			Case "Mz" : sA = _sbMz
			Case "Dz" : sA = _sbDz
			Case "Hz" : sA = _sbHz
			Case "月z" : sA = _sb月亮z
			Case "Dg-主氣" : sA = Mid(sDzz, 1, 1)
			Case "Dg-中氣" : sA = Mid(sDzz, 2, 1)
		End Select

		s = "" : For Each s1 As Char In sYzz : s += s1 + uf六親_支vs支_六行_丑辰元(sA, s1) + vbNL : Next : txt年柱4.Text = s.TrimEnd 'us父2梟印(s)
		s = "" : For Each s1 As Char In sMzz : s += s1 + uf六親_支vs支_六行_丑辰元(sA, s1) + vbNL : Next : txt月柱4.Text = s.TrimEnd 'us父2梟印(s)
		s = "" : For Each s1 As Char In sDzz : s += s1 + uf六親_支vs支_六行_丑辰元(sA, s1) + vbNL : Next : txt日柱4.Text = s.TrimEnd 'us父2梟印(s)
		s = "" : For Each s1 As Char In sHzz : s += s1 + uf六親_支vs支_六行_丑辰元(sA, s1) + vbNL : Next : txt時柱4.Text = s.TrimEnd 'us父2梟印(s)
		s = "" : For Each s1 As Char In sHzz : s += s1 + uf六親_支vs支_六行_丑辰元(sA, s1) + vbNL : Next : s月亮zz = s.TrimEnd 'us父2梟印(s)

		'us排大運()
		txt統計4.Text = us統計五行個目_地支_六行_丑辰元(sYMDHzz) ' + s月zz
	End Sub
	Private Sub us方法_壬癸己_六爻()
		Dim sYzz = "", sYg, sYgz, sYz1 As String, s = "", sA = "", sYMDHzz = ""
		Dim sMzz = "", sMg, sMgz, sMz1 As String
		Dim sDzz = "", sDg, sDgz, sDz1 As String
		Dim sHzz = "", sHg, sHgz, sHz1 As String
		Dim dic干藏支 As Dictionary(Of String, String) = Nothing

		sYg = uf60甲子轉換_壬癸己(_sbYg) : sYgz = sYg + _sbYz
		sMg = uf60甲子轉換_壬癸己(_sbMg) : sMgz = sMg + _sbMz
		sDg = uf60甲子轉換_壬癸己(_sbDg) : sDgz = sDg + _sbDz
		sHg = uf60甲子轉換_壬癸己(_sbHg) : sHgz = sHg + _sbHz
		'---------------------------------------------
		Select Case cbo干藏支.Text
			Case "甲=寅亥" : dic干藏支 = __dic干藏支_甲_寅亥
			Case "甲=寅卯" : dic干藏支 = __dic干藏支_甲_寅卯_乙_卯未
		End Select

		sYz1 = Mid(dic干藏支(sYg), 1, 1) : sYzz = sYz1 + _sbYz
		sMz1 = Mid(dic干藏支(sMg), 1, 1) : sMzz = sMz1 + _sbMz
		sDz1 = Mid(dic干藏支(sDg), 1, 1) : sDzz = sDz1 + _sbDz
		sHz1 = Mid(dic干藏支(sHg), 1, 1) : sHzz = sHz1 + _sbHz
		sYMDHzz = sYzz + sMzz + sDzz + sHzz
		'---------------------------------------------
		Select Case cbo基準為何支4.Text
			Case "Yz" : sA = _sbYz
			Case "Mz" : sA = _sbMz
			Case "Dz" : sA = _sbDz
			Case "Hz" : sA = _sbHz
			Case "月z" : sA = _sb月亮z
			Case "Dg-主氣" : sA = Mid(sDzz, 1, 1)
			Case "Dg-中氣" : sA = Mid(sDzz, 2, 1)
		End Select
		'---------------------------------------------
		Dim sYY, sMM, sDD, sHH As String
		s = "" : For Each s1 As Char In sYzz : s += s1 + uf六親_支vs支_六行_丑辰元(sA, s1) + vbNL : Next : sYY = s.TrimEnd 'us父2梟印(s)
		s = "" : For Each s1 As Char In sMzz : s += s1 + uf六親_支vs支_六行_丑辰元(sA, s1) + vbNL : Next : sMM = s.TrimEnd 'us父2梟印(s)
		s = "" : For Each s1 As Char In sDzz : s += s1 + uf六親_支vs支_六行_丑辰元(sA, s1) + vbNL : Next : sDD = s.TrimEnd 'us父2梟印(s)
		s = "" : For Each s1 As Char In sHzz : s += s1 + uf六親_支vs支_六行_丑辰元(sA, s1) + vbNL : Next : sHH = s.TrimEnd 'us父2梟印(s)
		'---------------------------------------------
		sYY = uf單柱橫排(sYY, _sbYz) + uf_12長生_支支_六行_禄刃(sYz1, _sbYz) + " " + uf單柱引字_六行(sYgz) + _sYo
		sMM = uf單柱橫排(sMM, _sbMz) + uf_12長生_支支_六行_禄刃(sMz1, _sbMz) + " " + uf單柱引字_六行(sMgz) + _sMo
		sDD = uf單柱橫排(sDD, _sbDz) + uf_12長生_支支_六行_禄刃(sDz1, _sbDz) + " " + uf單柱引字_六行(sDgz) + _sDo
		sHH = uf單柱橫排(sHH, _sbHz) + uf_12長生_支支_六行_禄刃(sHz1, _sbHz) + " " + uf單柱引字_六行(sHgz) + _sHo

		txt八字_六爻法.Text = sYY + vbNL + sMM + vbNL + sDD + vbNL + sHH  '+ vbNL + s月亮
		'---------------------------------------------
		s = ""
		s += uf命局缺的字_支支_六行(sA, sYY + sMM + sDD + sHH) ' + s月柱
		's += "Y" + sYgz + "->" + uf干支巔倒_六行(sYgz) + "，" + "M" + sMgz + "->" + uf干支巔倒_六行(sMgz) + "，" +
		'		 "D" + sDgz + "->" + uf干支巔倒_六行(sDgz) + "，" + "H" + sHgz + "->" + uf干支巔倒_六行(sHgz) + vbNL
		's += Uf八字組合_六行(sYMDHgz) ' 用不到
		's += uf天地沖合_六行(sYMDHgz) ' 用不到
		s += "D" + sDg + "@M" + _sbMz + "=" + uf_12長生_支支_六行_禄刃(sDz1, _sbMz) + vbTab
		txt組合_六爻法.Text = s
		'---------------------------------------------
		'us排大運()
		'txt統計4.Text = us統計五行個目_地支_六行_丑辰元(sYMDHzz) ' + s月zz
	End Sub

	Private Sub us方法_八行法_支()
		Dim sYzz = "", sMzz = "", sDzz = "", sHzz = "", s = "", sA = ""
		Dim dic支藏干 As Dictionary(Of String, String) = Nothing

		Select Case cbo干藏支.Text
			Case "甲=寅亥" : dic支藏干 = __dic干藏支_甲_寅亥
			Case "甲=寅卯" : dic支藏干 = __dic干藏支_甲_寅卯_乙_卯未
		End Select
		sYzz = dic支藏干(_sbYg) + _sbYz
		sMzz = dic支藏干(_sbMg) + _sbMz
		sDzz = dic支藏干(_sbDg) + _sbDz
		sHzz = dic支藏干(_sbHg) + _sbHz
		'---------------------------------------------
		Select Case cbo基準為何支5.Text
			Case "Yz" : sA = _sbYz
			Case "Mz" : sA = _sbMz
			Case "Dz" : sA = _sbDz
			Case "Hz" : sA = _sbHz
			Case "Dg-主氣" : sA = Mid(sDzz, 1, 1)
			Case "Dg-中氣" : sA = Mid(sDzz, 2, 1)
		End Select

		s = "" : For Each s1 As Char In sYzz : s += s1 + uf八親_支vs支(sA, s1) + vbNL : Next : txt年柱5.Text = s.TrimEnd
		s = "" : For Each s1 As Char In sMzz : s += s1 + uf八親_支vs支(sA, s1) + vbNL : Next : txt月柱5.Text = s.TrimEnd
		s = "" : For Each s1 As Char In sDzz : s += s1 + uf八親_支vs支(sA, s1) + vbNL : Next : txt日柱5.Text = s.TrimEnd
		s = "" : For Each s1 As Char In sHzz : s += s1 + uf八親_支vs支(sA, s1) + vbNL : Next : txt時柱5.Text = s.TrimEnd

		txt統計5.Text = us統計五行個目_地支_八行(sYzz + sMzz + sDzz + sHzz)
	End Sub

	Private Function uf60甲子轉換_壬癸己(sG As String) As String
		Select Case sG
			Case "己" : sG = "庚"
			Case "庚" : sG = "辛"
			Case "辛" : sG = "壬"
			Case "壬" : sG = "癸"
			Case "癸" : sG = "己"
		End Select
		Return sG
	End Function
	Private Function uf60甲子轉換_戊甲乙(sG As String) As String
		Select Case sG
			Case "甲" : sG = "戊"
			Case "乙" : sG = "甲"
			Case "丙" : sG = "乙"
			Case "丁" : sG = "丙"
			Case "戊" : sG = "丁"
		End Select
		Return sG
	End Function
	'================================================================
	Private Sub us排大運_月柱()
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		'大運以月柱為準
		'(++得+，--得+) 男命+阳年生 順排，女命+阴年生 順排
		'(+-得-，-+得-) 男命+阴年生 逆排，女命+阳年生 逆排

		Dim ir, i在何干支 As Int32, sGZ = "", s大運 = ""
		Dim is男命 = chk男命.Checked
		Dim is阳年生 = uf阴阳年生(_sbYg)

		For ir = 1 To 120
			If _s60甲子擴充_(ir).Equals(_sbMgz) Then
				If ir < 12 Then Continue For ' 若太小，如ir=2，大運逆排時，index 會不夠用
				i在何干支 = ir : Exit For
			End If
		Next
		'---------------------------------------------
		Select Case True
			Case is男命 AndAlso is阳年生, Not is男命 AndAlso Not is阳年生
				For ir = i在何干支 + 1 To i在何干支 + 10
					sGZ = _s60甲子擴充_(ir) : s大運 += sGZ + vbNL
				Next

			Case is男命 AndAlso Not is阳年生, Not is男命 AndAlso is阳年生
				For ir = i在何干支 - 1 To i在何干支 - 10 Step -1
					sGZ = _s60甲子擴充_(ir) : s大運 += sGZ + vbNL
				Next
		End Select

		_s大運_傳統法 = s大運.TrimEnd
	End Sub
	Private Sub us排大運_年柱()
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		'大運以月柱為準
		'(++得+，--得+) 男命+阳年生 順排，女命+阴年生 順排
		'(+-得-，-+得-) 男命+阴年生 逆排，女命+阳年生 逆排

		Dim ir, i在何干支 As Int32, sGZ = "", s大運 = ""
		Dim is男命 = chk男命.Checked
		'Dim is阳年生 = uf阴阳年生(_sbYg)
		If chk大運_男順女順.Checked Then is男命 = True

		For ir = 1 To 120
			If _s60甲子擴充_(ir).Equals(_sbYgz) Then
				If ir < 12 Then Continue For ' 若太小，如ir=2，大運逆排時，index 會不夠用
				i在何干支 = ir : Exit For
			End If
		Next
		'---------------------------------------------
		Select Case True
			Case is男命
				'Case is男命 AndAlso is阳年生, Not is男命 AndAlso Not is阳年生
				For ir = i在何干支 To i在何干支 + 11
					sGZ = _s60甲子擴充_(ir) : s大運 += sGZ + vbNL
				Next

			Case Else
				'Case is男命 AndAlso Not is阳年生, Not is男命 AndAlso is阳年生
				For ir = i在何干支 To i在何干支 - 11 Step -1
					sGZ = _s60甲子擴充_(ir) : s大運 += sGZ + vbNL
				Next
		End Select

		_s大運_壬癸己 = s大運.TrimEnd
	End Sub
	Public Function uf阴阳年生(sG As String) As Boolean
		Select Case sG
			Case "甲", "丙", "戊", "庚", "壬" : Return True
			Case Else : Return False
		End Select
		Return False
	End Function

	Dim _dic大運_傳統法, _dic大運_壬癸己 As New Dictionary(Of Int32, String)
	Private Sub us算大運_傳統法()
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		Dim s全部 = "", sGG = "", s, sG, sZ, sGZ, s岁年, s禄刃, s串宮, s_() As String, i大运年, i大运岁, i增年 As Int32

		If txt大運增年.Text.Trim.Length = 0 Then txt大運增年.Text = "12"
		i增年 = 10 : i大运年 = _i大运起年 - i增年 : i大运岁 = _i大运起岁 - i增年
		_dic大運_傳統法.Clear()
		s_ = Split(_s大運_傳統法, vbNL)
		For Each sGZ In s_
			sG = Mid(sGZ, 1, 1) : sZ = Mid(sGZ, 2, 1)
			sGG = sG + __dic支藏干_古法(sZ) : s = ""  ' 甲戌=甲[戊辛丁]
			For Each s1 As Char In sGG
				s += s1 + uf_印綬_干vs干_五行(_s基準_傳統法, s1) + vbNL
			Next

			i大运年 += i增年 : i大运岁 += i增年
			s禄刃 = uf_12長生_干支_五行_禄刃(_s基準_傳統法, sZ)
			s串宮 = "|" + _dic串宮_12神(_dic串宮_干2支_古法(sG)) + "+" + _dic串宮_12神(sZ) + "|"
			s岁年 = i大运岁.ToString("00") + "岁 " + Mid(i大运年.ToString, 3, 2) + "b"
			s = s岁年 + uf單柱橫排(s.TrimEnd, sZ) + " " + uf單柱引字_五行(sGZ) + s禄刃
			s全部 += s + vbNL : _dic大運_傳統法.Add(i大运年, s)
		Next
		txt大運_傳統法.Text = s全部.TrimEnd
	End Sub
	Private Sub us算大運_壬癸己()
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		Dim s全部 = "", sGG = "", s, sG, sZ, sGZ, sGZ2, s岁年, s禄刃, s串宮, s_() As String, i大运年, i大运岁, i增年 As Int32

		If txt大運增年.Text.Trim.Length = 0 Then txt大運增年.Text = "10"
		If cbo大運年月.Text.Equals("大運=月") Then
			i增年 = 10 : i大运年 = _i大运起年 - i增年 : i大运岁 = _i大运起岁 - i增年
		Else
			i增年 = CInt(txt大運增年.Text) : i大运年 = _i出生年 - i增年 : i大运岁 = -i增年 + 1
		End If

		_dic大運_壬癸己.Clear()
		s_ = Split(_s大運_壬癸己, vbNL)
		For Each sGZ In s_
			sG = Mid(sGZ, 1, 1) : sZ = Mid(sGZ, 2, 1)
			sG = uf60甲子轉換_壬癸己(sG) : sGZ2 = sG + sZ : s = ""

			Select Case cbo支藏干.Text
				Case "寅[甲丙戊]" : sGG = sG + __dic支藏干_六行_寅_甲丙戊(sZ)
				Case "寅[甲丙己]" : sGG = sG + __dic支藏干_六行_寅_甲丙己(sZ)
				Case "寅[甲戊]" : sGG = sG + __dic支藏干_六行_寅_甲戊(sZ)
			End Select

			For Each s1 As Char In sGG
				s += s1 + uf_印綬_干vs干_六行_奇隅(_s基準_壬癸己, s1) + vbNL
			Next

			i大运年 += i增年 : i大运岁 += i增年
			s禄刃 = uf_12長生_干支_六行_禄刃(_s基準_壬癸己, sZ)
			s串宮 = "|" + _dic串宮_12神(_dic串宮_干2支_六行(sG)) + "+" + _dic串宮_12神(sZ) + "|"
			s岁年 = i大运岁.ToString("00") + "岁 " + Mid(i大运年.ToString, 3, 2) + "b"
			s = s岁年 + uf單柱橫排(s.TrimEnd, sZ) + " " + uf單柱引字_六行(sGZ2) + s禄刃
			s全部 += s + vbNL : _dic大運_壬癸己.Add(i大运年, s)
		Next
		txt大運_壬癸己.Text = s全部.TrimEnd
	End Sub

	Private Sub us算流年_傳統法()
		If _i大运起年 < 0 Then MsgBox("只能算西元后出生的人") : Exit Sub
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		Dim s全部 = "", sGG = "", s, s2, sG, sZ, sGZ, s岁年, s禄刃, s串宮, s多月干 As String, iY, i起年, i生年 As Int32

		If cbo大運年月.Text.Equals("大運=月") Then
			i起年 = _i大运起年 : i生年 = i起年 - _i大运起岁 + 1 ' 因都四拾五入，+1回來
		Else
			i起年 = _i出生年 : i生年 = _i出生年
		End If

		For iY = i起年 To i起年 + 100
			sGZ = __dic西元年_60甲子(iY) : sG = Mid(sGZ, 1, 1) : sZ = Mid(sGZ, 2, 1)
			sGG = sG + __dic支藏干_古法(sZ) : s = ""  ' 甲戌=甲戊辛丁
			For Each s1 As Char In sGG
				s += s1 + uf_印綬_干vs干_五行(_s基準_傳統法, s1) + vbNL
			Next

			s多月干 = m易理.uf流年多出的月干_五行(_s基準_傳統法, sG)
			s禄刃 = uf_12長生_干支_五行_禄刃(_s基準_傳統法, sZ)
			s串宮 = "|" + _dic串宮_12神(_dic串宮_干2支_古法(sG)) + "+" + _dic串宮_12神(sZ) + "|"
			s岁年 = (iY - i生年 + 1).ToString("00") + "岁 " + Mid(iY.ToString, 3, 2) + "y"

			If _dic大運_傳統法.ContainsKey(iY) Then ' >9岁，让第1筆不要+"----"
				s2 = "" : If iY - i生年 > 1 Then s2 = "---------------------------------------------------" + vbNL
				s全部 += s2 + _dic大運_傳統法(iY).Replace("岁", "●") + vbNL
			End If
			s全部 += s岁年 + uf單柱橫排(s.TrimEnd, sZ) + " " + uf單柱引字_五行(sGZ) + s禄刃 + s多月干 + vbNL
		Next

		Dim idx = txt流年_傳統法.GetFirstCharIndexOfCurrentLine
		Dim ir = txt流年_傳統法.GetLineFromCharIndex(idx) + 1 ' 列號
		Dim ic = txt流年_傳統法.SelectionStart - idx + 1      ' 行號
		txt流年_傳統法.Text = s全部.TrimEnd
		txt流年_傳統法.Select(idx, 0)
		txt流年_傳統法.ScrollToCaret() ' 滖动到该位置
	End Sub
	Private Sub us算流年_壬癸己()
		If _i大运起年 < 0 Then MsgBox("只能算西元后出生的人") : Exit Sub
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		Dim s全部 = "", sGG = "", s, s2, sG, sZ, sGZ, sGZ2, s岁年, s禄刃, s串宮, s多月干, s年神煞 As String, iY, i起年, i生年 As Int32

		If cbo大運年月.Text.Equals("大運=月") Then
			i起年 = _i大运起年 : i生年 = i起年 - _i大运起岁 + 1 ' 因都四拾五入，+1回來
		Else
			i起年 = _i出生年 : i生年 = _i出生年
		End If

		For iY = i起年 To i起年 + 100
			sGZ = __dic西元年_60甲子(iY) : sG = Mid(sGZ, 1, 1) : sZ = Mid(sGZ, 2, 1)
			sG = uf60甲子轉換_壬癸己(sG) : sGZ2 = sG + sZ : s = ""

			Select Case cbo支藏干.Text
				Case "寅[甲丙戊]" : sGG = sG + __dic支藏干_六行_寅_甲丙戊(sZ)
				Case "寅[甲丙己]" : sGG = sG + __dic支藏干_六行_寅_甲丙己(sZ)
				Case "寅[甲戊]" : sGG = sG + __dic支藏干_六行_寅_甲戊(sZ)
			End Select

			For Each s1 As Char In sGG
				s += s1 + uf_印綬_干vs干_六行_奇隅(_s基準_壬癸己, s1) + vbNL
			Next

			s多月干 = m易理.uf流年多出的月干_六行(_s基準_壬癸己, sG)
			s禄刃 = uf_12長生_干支_六行_禄刃(_s基準_壬癸己, sZ)
			s串宮 = "|" + _dic串宮_12神(_dic串宮_干2支_六行(sG)) + "+" + _dic串宮_12神(sZ) + "|"
			s年神煞 = uf_神煞_年支_短(_sbYz, sZ)
			s岁年 = (iY - i生年 + 1).ToString("00") + "岁 " + Mid(iY.ToString, 3, 2) + "y"

			If _dic大運_壬癸己.ContainsKey(iY) Then ' >9岁，让第1筆不要+"----"
				s2 = "" : If iY - i生年 > 9 Then s2 = "---------------------------------------------------" + vbNL
				s全部 += s2 + _dic大運_壬癸己(iY).Replace("岁", "●") + vbNL
			End If

			s全部 += s岁年 + uf單柱橫排(s.TrimEnd, sZ) + " " + uf單柱引字_六行(sGZ2) + s禄刃 + s多月干 + s年神煞 + vbNL
		Next


		Dim idx = txt流年_壬癸己.GetFirstCharIndexOfCurrentLine
		Dim ir = txt流年_壬癸己.GetLineFromCharIndex(idx) + 1 ' 列號
		Dim ic = txt流年_壬癸己.SelectionStart - idx + 1      ' 行號
		txt流年_壬癸己.Text = s全部.TrimEnd
		txt流年_壬癸己.Select(idx, 0)
		txt流年_壬癸己.ScrollToCaret() ' 滖动到该位置
	End Sub
	'================================================================
	Private Function us統計五行個數_天干(sYMDHgg As String, sYMDHgz As String) As String
		'sYMDH=甲子乙丑庚午丙戌
		Dim i木, i火, i土, i金, i水 As Int32
		Dim d甲, d乙, d丙, d丁, d戊, d己, d庚, d辛, d壬, d癸 As Double

		For Each s As Char In sYMDHgg ' 統計個數
			Select Case s
				Case "甲"c, "乙"c : i木 += 1
				Case "丙"c, "丁"c : i火 += 1
				Case "戊"c, "己"c : i土 += 1
				Case "庚"c, "辛"c : i金 += 1
				Case "壬"c, "癸"c : i水 += 1
			End Select
		Next
		'---------------------------------------------
		For Each s As Char In sYMDHgz ' 統計分數
			Select Case s
				Case "甲"c : d甲 += 10 : Case "乙"c : d乙 += 10 : Case "丙"c : d丙 += 10 : Case "丁"c : d丁 += 10 : Case "戊"c : d戊 += 10
				Case "己"c : d己 += 10 : Case "庚"c : d庚 += 10 : Case "辛"c : d辛 += 10 : Case "壬"c : d壬 += 10 : Case "癸"c : d癸 += 10
					' ref: 李洪成，地支藏干
				Case "子"c : d癸 += 10
				Case "卯"c : d乙 += 10
				Case "午"c : d丁 += 7 : d己 += 3
				Case "酉"c : d辛 += 10

				Case "寅"c : d甲 += 6 : d丙 += 3 : d戊 += 1
				Case "巳"c : d丙 += 6 : d戊 += 3 : d庚 += 1
				Case "申"c : d庚 += 6 : d壬 += 3 : d戊 += 1
				Case "亥"c : d壬 += 7 : d甲 += 3

				Case "丑"c : d己 += 6 : d辛 += 3 : d癸 += 1
				Case "辰"c : d戊 += 6 : d乙 += 3 : d癸 += 1
				Case "未"c : d己 += 6 : d丁 += 3 : d乙 += 1
				Case "戌"c : d戊 += 6 : d辛 += 3 : d丁 += 1
			End Select
		Next
		'---------------------------------------------
		Dim s木 = "", s火 = "", s土 = "", s金 = "", s水 = ""

		Select Case _s基準_傳統法
			Case "甲", "乙" : s木 = "●木" : s火 = "仔火" : s土 = "妻土" : s金 = "官金" : s水 = "父水"
			Case "丙", "丁" : s木 = "父木" : s火 = "●火" : s土 = "仔土" : s金 = "妻金" : s水 = "官水"
			Case "戊", "己" : s木 = "官木" : s火 = "父火" : s土 = "●土" : s金 = "仔金" : s水 = "妻水"
			Case "庚", "辛" : s木 = "妻木" : s火 = "官火" : s土 = "父土" : s金 = "●金" : s水 = "仔水"
			Case "壬", "癸" : s木 = "仔木" : s火 = "妻火" : s土 = "官土" : s金 = "父金" : s水 = "●水"
		End Select
		'---------------------------------------------
		If True Then
			Return s木 + i木.ToString + "，" + (d甲 + d乙).ToString + vbNL +
						 s火 + i火.ToString + "，" + (d丙 + d丁).ToString + vbNL +
						 s土 + i土.ToString + "，" + (d戊 + d己).ToString + vbNL +
						 s金 + i金.ToString + "，" + (d庚 + d辛).ToString + vbNL +
						 s水 + i水.ToString + "，" + (d壬 + d癸).ToString
		Else
			Return "木火土金水" + vbNL +
					 StrConv(i木.ToString, VbStrConv.Wide) +
					 StrConv(i火.ToString, VbStrConv.Wide) +
					 StrConv(i土.ToString, VbStrConv.Wide) +
					 StrConv(i金.ToString, VbStrConv.Wide) +
					 StrConv(i水.ToString, VbStrConv.Wide) + vbNL +
					 "木 火 土 金 水" + vbNL +
				 (d甲 + d乙).ToString + " " + (d丙 + d丁).ToString + " " + (d戊 + d己).ToString + " " +
				 (d庚 + d辛).ToString + " " + (d壬 + d癸).ToString
		End If
	End Function
	Private Function us統計六行個數_天干(sYMDHgg As String, sYMDHgz As String) As String
		'sYMDH=甲子乙丑庚午丙戌
		Dim i木, i火, i戊, i金, i水, i己 As Int32
		Dim d甲, d乙, d丙, d丁, d戊, d己, d庚, d辛, d壬, d癸 As Double

		For Each s As Char In sYMDHgg
			Select Case s
				Case "甲"c, "乙"c : i木 += 1
				Case "丙"c, "丁"c : i火 += 1
				Case "戊"c : i戊 += 1
				Case "己"c : i己 += 1
				Case "庚"c, "辛"c : i金 += 1
				Case "壬"c, "癸"c : i水 += 1
			End Select
		Next
		'---------------------------------------------
		For Each s As Char In sYMDHgz
			Select Case s
				Case "甲"c : d甲 += 10 : Case "乙"c : d乙 += 10 : Case "丙"c : d丙 += 10 : Case "丁"c : d丁 += 10 : Case "戊"c : d戊 += 10
				Case "己"c : d己 += 10 : Case "庚"c : d庚 += 10 : Case "辛"c : d辛 += 10 : Case "壬"c : d壬 += 10 : Case "癸"c : d癸 += 10
			End Select
		Next

		If False Then ' 試3
			For Each s As Char In sYMDHgz
				Select Case s
					Case "子"c : d癸 += 6 : d壬 += 4
					Case "卯"c : d乙 += 6 : d甲 += 4
					Case "午"c : d丁 += 6 : d丙 += 4
					Case "酉"c : d辛 += 6 : d庚 += 4

					Case "寅"c : d甲 += 6 : d丙 += 2 : d戊 += 2 ' 变
					Case "巳"c : d丙 += 6 : d庚 += 2 : d戊 += 2
					Case "申"c : d庚 += 6 : d壬 += 2 : d己 += 2 ' 变
					Case "亥"c : d壬 += 6 : d甲 += 2 : d己 += 2

					Case "丑"c : d己 += 4 : d癸 += 2 : d辛 += 4
					Case "辰"c : d己 += 4 : d乙 += 2 : d癸 += 4
					Case "未"c : d戊 += 4 : d丁 += 2 : d乙 += 4
					Case "戌"c : d戊 += 4 : d辛 += 2 : d丁 += 4
				End Select
			Next
		Else ' 試5-3
			For Each s As Char In sYMDHgz
				Select Case s
					Case "子"c : d癸 += 7 : d壬 += 3
					Case "卯"c : d乙 += 7 : d甲 += 3
					Case "午"c : d丁 += 7 : d丙 += 3
					Case "酉"c : d辛 += 7 : d庚 += 3

					Case "寅"c : d甲 += 7 : d丙 += 2 : d戊 += 1 ' 变
					Case "巳"c : d丙 += 7 : d庚 += 2 : d戊 += 1
					Case "申"c : d庚 += 7 : d壬 += 2 : d己 += 1 ' 变
					Case "亥"c : d壬 += 7 : d甲 += 2 : d己 += 1

					Case "丑"c : d己 += 5 : d癸 += 2 : d辛 += 3
					Case "辰"c : d己 += 5 : d乙 += 2 : d癸 += 3
					Case "未"c : d戊 += 5 : d丁 += 2 : d乙 += 3
					Case "戌"c : d戊 += 5 : d辛 += 2 : d丁 += 3
				End Select
			Next
		End If
		'---------------------------------------------
		Dim s木 = "", s火 = "", s戊 = "", s金 = "", s水 = "", s己 = ""

		Select Case _s基準_壬癸己
			Case "甲", "乙" : s木 = "●木" : s火 = "仔火" : s戊 = "妻戊" : s金 = "奴金" : s水 = "官水" : s己 = "父己"
			Case "丙", "丁" : s木 = "父木" : s火 = "●火" : s戊 = "仔戊" : s金 = "妻金" : s水 = "奴水" : s己 = "官己"
			Case "庚", "辛" : s木 = "奴木" : s火 = "官火" : s戊 = "父戊" : s金 = "●金" : s水 = "仔水" : s己 = "妻己"
			Case "壬", "癸" : s木 = "妻木" : s火 = "奴火" : s戊 = "官戊" : s金 = "父金" : s水 = "●水" : s己 = "仔己"
			Case "戊" : s木 = "官木" : s火 = "父火" : s戊 = "●戊" : s金 = "仔金" : s水 = "妻水" : s己 = "奴己"
			Case "己" : s木 = "仔木" : s火 = "妻火" : s戊 = "奴戊" : s金 = "官金" : s水 = "父水" : s己 = "●己"
		End Select
		'---------------------------------------------
		If True Then ' 直排
			Return s木 + i木.ToString + "，" + (d甲 + d乙).ToString + vbNL +
						 s火 + i火.ToString + "，" + (d丙 + d丁).ToString + vbNL +
						 s戊 + i戊.ToString + "，" + d戊.ToString + vbNL +
						 s金 + i金.ToString + "，" + (d庚 + d辛).ToString + vbNL +
						 s水 + i水.ToString + "，" + (d壬 + d癸).ToString + vbNL +
						 s己 + i己.ToString + "，" + d己.ToString
		Else
			Return "木火戊金水己" + vbNL +
					 StrConv(i木.ToString, VbStrConv.Wide) + ' 全形
					 StrConv(i火.ToString, VbStrConv.Wide) +
					 StrConv(i戊.ToString, VbStrConv.Wide) +
					 StrConv(i金.ToString, VbStrConv.Wide) +
					 StrConv(i水.ToString, VbStrConv.Wide) +
					 StrConv(i己.ToString, VbStrConv.Wide) +
				"木 火 戊 金 水 己" + vbNL +
				(d甲 + d乙).ToString + " " + (d丙 + d丁).ToString + " " + (d戊).ToString + " " +
				(d庚 + d辛).ToString + " " + (d壬 + d癸).ToString + " " + (d己).ToString
		End If
	End Function
	Private Function us統計五行個目_地支_六行_辰未元(sYMDH As String) As String
		Dim i木, i火, i戊, i金, i水, i己 As Int32

		For Each s As Char In sYMDH
			Select Case s
				Case "寅"c, "卯"c : i木 += 1
				Case "巳"c, "午"c : i火 += 1
				Case "申"c, "酉"c : i金 += 1
				Case "亥"c, "子"c : i水 += 1

				Case "辰"c, "未"c : i己 += 1
				Case "丑"c, "戌"c : i戊 += 1
			End Select
		Next

		If True Then
			Return "木=" + i木.ToString + vbNL +
						 "火=" + i火.ToString + vbNL +
						 "戊=" + i戊.ToString + vbNL +
						 "金=" + i金.ToString + vbNL +
						 "水=" + i水.ToString + vbNL +
						 "己=" + i己.ToString
		Else
			Return "木火戊金水己" + vbNL +
						 StrConv(i木.ToString, VbStrConv.Wide) +
						 StrConv(i火.ToString, VbStrConv.Wide) +
						 StrConv(i戊.ToString, VbStrConv.Wide) +
						 StrConv(i金.ToString, VbStrConv.Wide) +
						 StrConv(i水.ToString, VbStrConv.Wide) +
						 StrConv(i己.ToString, VbStrConv.Wide)
		End If
	End Function
	Private Function us統計五行個目_地支_六行_丑辰元(sYMDHzz As String) As String
		Dim i木, i火, i戊, i金, i水, i己 As Int32

		For Each s As Char In sYMDHzz
			Select Case s
				Case "寅"c, "卯"c : i木 += 1
				Case "巳"c, "午"c : i火 += 1
				Case "申"c, "酉"c : i金 += 1
				Case "亥"c, "子"c : i水 += 1

				Case "辰"c, "丑"c : i己 += 1
				Case "未"c, "戌"c : i戊 += 1
			End Select
		Next

		If True Then
			Return "木=" + i木.ToString + vbNL +
			 "火=" + i火.ToString + vbNL +
			 "戊=" + i戊.ToString + vbNL +
			 "金=" + i金.ToString + vbNL +
			 "水=" + i水.ToString + vbNL +
			 "己=" + i己.ToString

		Else
			Return "木火土金水元" + vbNL +
			 StrConv(i木.ToString, VbStrConv.Wide) +
			 StrConv(i火.ToString, VbStrConv.Wide) +
			 StrConv(i戊.ToString, VbStrConv.Wide) +
			 StrConv(i金.ToString, VbStrConv.Wide) +
			 StrConv(i水.ToString, VbStrConv.Wide) +
			 StrConv(i己.ToString, VbStrConv.Wide)
		End If
	End Function
	Private Function us統計五行個目_地支_八行(sYMDH As String) As String
		Dim i木, i火, i金, i水, i丑, i辰, i未, i戌 As Int32

		For Each s As Char In sYMDH
			Select Case s
				Case "寅"c, "卯"c : i木 += 1
				Case "巳"c, "午"c : i火 += 1
				Case "申"c, "酉"c : i金 += 1
				Case "亥"c, "子"c : i水 += 1
				Case "丑"c : i丑 += 1
				Case "辰"c : i辰 += 1
				Case "未"c : i未 += 1
				Case "戌"c : i戌 += 1
			End Select
		Next

		If True Then
			Return "丑=" + i丑.ToString + vbNL +
						 "木=" + i木.ToString + vbNL +
						 "辰=" + i辰.ToString + vbNL +
						 "火=" + i火.ToString + vbNL +
						 "未=" + i未.ToString + vbNL +
						 "金=" + i金.ToString + vbNL +
						 "戌=" + i戌.ToString + vbNL +
						 "水=" + i水.ToString
		Else
			Return "木辰火未金戌水丑" + vbNL + ' 木杰火鈥金淦水沐
						 StrConv(i木.ToString, VbStrConv.Wide) +
						 StrConv(i辰.ToString, VbStrConv.Wide) +
						 StrConv(i火.ToString, VbStrConv.Wide) +
						 StrConv(i未.ToString, VbStrConv.Wide) +
						 StrConv(i金.ToString, VbStrConv.Wide) +
						 StrConv(i戌.ToString, VbStrConv.Wide) +
						 StrConv(i水.ToString, VbStrConv.Wide) +
						 StrConv(i丑.ToString, VbStrConv.Wide)
		End If
	End Function
	'================================================================
	Private Sub lsb流年干_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lsb流年干.SelectedIndexChanged
		If Not Me.Visible Then Exit Sub
		us算流月()
	End Sub
	Private Sub us算流月()
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		Dim s流年干1, s流年干2, s As String
		Dim s1_() As String = Nothing, s2_() As String = Nothing
		s流年干1 = lsb流年干.Text
		s流年干2 = uf60甲子轉換_壬癸己(s流年干1)
		Select Case s流年干1
			Case "甲", "己" : s1_ = {"丙寅", "丁卯", "戊辰", "己巳", "庚午", "辛未", "壬申", "癸酉", "甲戌", "乙亥", "丙子", "丁丑"}
			Case "乙", "庚" : s1_ = {"戊寅", "己卯", "庚辰", "辛巳", "壬午", "癸未", "甲申", "乙酉", "丙戌", "丁亥", "戊子", "己丑"}
			Case "丙", "辛" : s1_ = {"庚寅", "辛卯", "壬辰", "癸巳", "甲午", "乙未", "丙申", "丁酉", "戊戌", "己亥", "庚子", "辛丑"}
			Case "丁", "壬" : s1_ = {"壬寅", "癸卯", "甲辰", "乙巳", "丙午", "丁未", "戊申", "己酉", "庚戌", "辛亥", "壬子", "癸丑"}
			Case "戊", "癸" : s1_ = {"甲寅", "乙卯", "丙辰", "丁巳", "戊午", "己未", "庚申", "辛酉", "壬戌", "癸亥", "甲子", "乙丑"}
		End Select

		Select Case s流年干2
			Case "甲", "庚" : s2_ = {"丙寅", "丁卯", "戊辰", "庚巳", "辛午", "壬未", "癸申", "己酉", "甲戌", "乙亥", "丙子", "丁丑"}
			Case "乙", "辛" : s2_ = {"戊寅", "庚卯", "辛辰", "壬巳", "癸午", "己未", "甲申", "乙酉", "丙戌", "丁亥", "戊子", "庚丑"}
			Case "丙", "壬" : s2_ = {"辛寅", "壬卯", "癸辰", "己巳", "甲午", "乙未", "丙申", "丁酉", "戊戌", "庚亥", "辛子", "壬丑"}
			Case "丁", "癸" : s2_ = {"癸寅", "己卯", "甲辰", "乙巳", "丙午", "丁未", "戊申", "庚酉", "辛戌", "壬亥", "癸子", "己丑"}
			Case "戊", "己" : s2_ = {"甲寅", "乙卯", "丙辰", "丁巳", "戊午", "庚未", "辛申", "壬酉", "癸戌", "己亥", "甲子", "乙丑"}
		End Select
		'---------------------------------------------
		Dim sG, sZ, sZgg, s1, s2, s3 As String, iL As Int32
		Dim sZ1g = "", sZ2g = "", sZ3g = "", sOut1 = "", sOut2 = ""
		'---------------------------------------------
		For Each s In s1_
			sG = Mid(s, 1, 1) : sZ = Mid(s, 2, 1)
			sZgg = __dic支藏干_古法(sZ) : iL = sZgg.Length
			s1 = "" : If iL > 0 Then s1 = Mid(sZgg, 1, 1) : sZ1g = s1 + uf_印綬_干vs干_五行(_s基準_傳統法, s1)
			s2 = "" : If iL > 1 Then s2 = Mid(sZgg, 2, 1) : sZ2g = s2 + uf_印綬_干vs干_五行(_s基準_傳統法, s2)
			s3 = "" : If iL > 2 Then s3 = Mid(sZgg, 3, 1) : sZ3g = s3 + uf_印綬_干vs干_五行(_s基準_傳統法, s3)
			sG = uf_印綬_干vs干_五行(_s基準_傳統法, sG) + sG ' 天干
			Select Case iL
				Case 1 : s = sG + " " + sZ + "[" + sZ1g + "]"
				Case 2 : s = sG + " " + sZ + "[" + sZ1g + "+" + sZ2g + "]"
				Case 3 : s = sG + " " + sZ + "[" + sZ1g + "+" + sZ2g + "+" + sZ3g + "]"
			End Select
			sOut1 += s + vbNL
		Next
		'---------------------------------------------
		For Each s In s2_
			sG = Mid(s, 1, 1) : sZ = Mid(s, 2, 1)
			sZgg = __dic支藏干_六行_寅_甲丙戊(sZ) : iL = sZgg.Length
			s1 = "" : If iL > 0 Then s1 = Mid(sZgg, 1, 1) : sZ1g = s1 + uf_印綬_干vs干_六行_奇隅(_s基準_壬癸己, s1)
			s2 = "" : If iL > 1 Then s2 = Mid(sZgg, 2, 1) : sZ2g = s2 + uf_印綬_干vs干_六行_奇隅(_s基準_壬癸己, s2)
			s3 = "" : If iL > 2 Then s3 = Mid(sZgg, 3, 1) : sZ3g = s3 + uf_印綬_干vs干_六行_奇隅(_s基準_壬癸己, s3)
			sG = uf_印綬_干vs干_六行_奇隅(_s基準_壬癸己, sG) + sG ' 天干
			Select Case iL
				Case 1 : s = sG + " " + sZ + "[" + sZ1g + "]"
				Case 2 : s = sG + " " + sZ + "[" + sZ1g + "+" + sZ2g + "]"
				Case 3 : s = sG + " " + sZ + "[" + sZ1g + "+" + sZ2g + "+" + sZ3g + "]"
			End Select
			sOut2 += s + vbNL
		Next
		'---------------------------------------------
		txt流月_傳統法.Text = sOut1.TrimEnd
		txt流月_壬癸己.Text = sOut2.TrimEnd
	End Sub

	Private Sub cbo干藏支_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo干藏支.SelectedIndexChanged
		If Me.Visible AndAlso cbo干藏支.Focused Then us方法_傳統法_支() : us方法_壬癸己_支() : us方法_壬癸己_六爻() : us方法_八行法_支()
	End Sub
	Private Sub cbo支藏干_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo支藏干.SelectedIndexChanged
		If Me.Visible AndAlso cbo干藏支.Focused Then us方法_壬癸己_支()
	End Sub
	Private Sub cbo基準為何支1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo基準為何支1.SelectedIndexChanged
		If Me.Visible AndAlso cbo基準為何支1.Focused Then us方法_傳統法_干()
	End Sub
	Private Sub cbo基準為何支2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo基準為何支2.SelectedIndexChanged
		If Me.Visible AndAlso cbo基準為何支2.Focused Then us方法_傳統法_支()
	End Sub
	Private Sub cbo基準為何支3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo基準為何支3.SelectedIndexChanged
		If Me.Visible AndAlso cbo基準為何支3.Focused Then us方法_壬癸己_干()
	End Sub
	Private Sub cbo基準為何支4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo基準為何支4.SelectedIndexChanged
		If Me.Visible AndAlso cbo基準為何支4.Focused Then us方法_壬癸己_支()
	End Sub
	Private Sub cbo基準為何支5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo基準為何支5.SelectedIndexChanged
		If Me.Visible AndAlso cbo基準為何支5.Focused Then us方法_壬癸己_六爻()
	End Sub
	'================================================================
	Private Sub cbo大運年月_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbo大運年月.SelectedIndexChanged
		If Me.Visible AndAlso cbo大運年月.Focused Then

			'us排大運_月柱() : us算大運_傳統法() : us算流年_傳統法()
			'---------------------------------------------
			If cbo大運年月.Text.Equals("大運=月") Then _s大運_壬癸己 = _s大運_傳統法 Else us排大運_年柱()
			us算大運_壬癸己() : us算流年_壬癸己()
		End If
	End Sub
	Private Sub chk男命_CheckStateChanged(sender As Object, e As EventArgs) Handles chk男命.CheckStateChanged
		If Not Me.Visible AndAlso Not chk男命.Focused Then Exit Sub

		If chk男命.Checked Then
			chk男命.Text = "男" : chk男命.BackColor = Color.Cyan : txt地支組合.ForeColor = Color.Cyan
		Else
			chk男命.Text = "女" : chk男命.BackColor = Color.Magenta : txt地支組合.ForeColor = Color.Magenta
		End If
		'---------------------------------------------
		If False Then
			us排大運_月柱() : us算大運_傳統法() : us算流年_傳統法()
			'--------------------
			If cbo大運年月.Text.Equals("大運=月") Then _s大運_壬癸己 = _s大運_傳統法 Else us排大運_年柱()
			us算大運_壬癸己() : us算流年_壬癸己()
		Else
			btn排盤.PerformClick()
		End If
	End Sub
	Private Sub chk大運_男順女順_CheckStateChanged(sender As Object, e As EventArgs) Handles chk大運_男順女順.CheckStateChanged
		If Not Me.Visible AndAlso Not chk大運_男順女順.Focused Then Exit Sub
		If False Then
			us排大運_年柱() : us算大運_壬癸己() : us算流年_壬癸己()
		Else
			btn排盤.PerformClick()
		End If
	End Sub
	'==================================================================
	Private Sub txt旬空_MouseClick(sender As Object, e As MouseEventArgs) Handles txt旬空.MouseClick
		Clipboard.SetText(txt旬空.Text)
	End Sub
	'Private Sub txt統計_全部_MouseClick(sender As Object, e As MouseEventArgs) Handles txt統計1.MouseClick, txt統計2.MouseClick, txt統計3.MouseClick, txt統計4.MouseClick
	'  Clipboard.SetText(DirectCast(sender, TextBox).Text)
	'End Sub
	Private Sub txtYMDHgz_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtYgz.KeyDown, txtMgz.KeyDown, txtDgz.KeyDown, txtHgz.KeyDown, txt月gz.KeyDown
		If e.KeyCode = Keys.Escape Then dgv60甲子.Hide()
	End Sub
	Private Sub txtYYMDHgz_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txtYgz.MouseClick, txtMgz.MouseClick, txtDgz.MouseClick, txtHgz.MouseClick, txt月gz.MouseClick
		'讓60甲子可以跳出來，方便選
		Dim ir, ic As Int32
		Dim txt = DirectCast(sender, TextBox), sGZ = txt.Text
		txt.SelectAll()

		'跳到目前的cell
		dgv60甲子.ClearSelection()
		If sGZ.Length > 0 AndAlso Not Mid(sGZ, 1, 1).Equals("Ｘ") Then
			For ir = 0 To 5
				For ic = 0 To 9
					If __s60甲子_(ir, ic).Equals(sGZ) Then dgv60甲子(ic, ir).Selected = True : GoTo 10
				Next
			Next
		End If

10: dgv60甲子.Tag = txt.Name
		dgv60甲子.Location = New Point(0, 25)
		dgv60甲子.BringToFront()
		dgv60甲子.Show()
	End Sub
	Private Sub txt大运起年_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txt出生年.MouseClick, txt大运起年.MouseClick, txt大运起岁.MouseClick
		DirectCast(sender, TextBox).SelectAll()
	End Sub
	Private Sub txt八字_傳統法_KeyDown(sender As Object, e As KeyEventArgs) Handles txt八字_傳統法.KeyDown, txt八字_壬癸己.KeyDown, txt八字_六爻法.KeyDown
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
		Select Case e.KeyCode
			Case Keys.Escape
				Dim ss, s, sA_(), s1_(), s2_(), s1, s2 As String
				Dim s11, s12, s13, s14, s21, s22, s23, s24 As String

				ss = uf記憶體取代字（Clipboard.GetText.Trim)
				'癸 丁 丙 甲  總字數=7+7=14
				'亥 巳 午 午
				sA_ = Split(ss, vbNL)

				If sA_.Length = 1 Then
					btn記憶體2四格.PerformClick()
				Else
					's1_ = Split(sA_(0).Replace("  ", " "), " ")
					's2_ = Split(sA_(1).Replace("  ", " "), " ")
					's = s1_(0) + s2_(0) + " " + s1_(1) + s2_(1) + " " + s1_(2) + s2_(2) + " " + s1_(3) + s2_(3)

					s1 = sA_(0).Replace("  ", " ").Replace(" ", "")
					s2 = sA_(1).Replace("  ", " ").Replace(" ", "")
					s11 = Mid(s1, 1, 1) : s12 = Mid(s1, 2, 1) : s13 = Mid(s1, 3, 1) : s14 = Mid(s1, 4, 1)
					s21 = Mid(s2, 1, 1) : s22 = Mid(s2, 2, 1) : s23 = Mid(s2, 3, 1) : s24 = Mid(s2, 4, 1)
					s = s11 + s21 + " " + s12 + s22 + " " + s13 + s23 + " " + s14 + s24
					Clipboard.SetText(s)
					btn記憶體2四格.PerformClick()
				End If

			Case Keys.F1 : us位置變換()
			Case Keys.F2 ' 字體
				Dim txt = DirectCast(sender, TextBox)
				Dim fnt = txt.Font
				If fnt.Name.Equals("微軟正黑體") Then
					txt.Font = New Font("MS Gothic", fnt.Size) ' MS Gothic, Cascadia Code, 細明體
				Else
					txt.Font = New Font("微軟正黑體", fnt.Size)
				End If

			Case Keys.F3
				If Not txt大運_六爻法.Visible Then
					txt大運_傳統法.Hide() : txt八字_傳統法.Hide() : txt組合_傳統法.Hide() : txt流月_傳統法.Hide() : txt流年_傳統法.Hide()
					txt大運_六爻法.Show() : txt八字_六爻法.Show() : txt組合_六爻法.Show() : txt流月_六爻法.Show() : txt流年_六爻法.Show()
				Else
					txt大運_傳統法.Show() : txt八字_傳統法.Show() : txt組合_傳統法.Show() : txt流月_傳統法.Show() : txt流年_傳統法.Show()
					txt大運_六爻法.Hide() : txt八字_六爻法.Hide() : txt組合_六爻法.Hide() : txt流月_六爻法.Hide() : txt流年_六爻法.Hide()
				End If
		End Select
	End Sub

	Private Sub txt組合_傳統法_MouseDown(sender As Object, e As MouseEventArgs) Handles txt組合_傳統法.MouseDown
		If e.Button = MouseButtons.Middle Then txt組合_傳統法.SendToBack()
	End Sub
	Private Sub txt組合_壬癸己_MouseDown(sender As Object, e As MouseEventArgs) Handles txt組合_壬癸己.MouseDown
		If e.Button = MouseButtons.Middle Then txt組合_壬癸己.SendToBack()
	End Sub
	Private Sub txt組合_六爻法_MouseDown(sender As Object, e As MouseEventArgs) Handles txt組合_六爻法.MouseDown
		If e.Button = MouseButtons.Middle Then txt組合_六爻法.SendToBack()
	End Sub

	Private Sub txt流月_傳統法_MouseDown(sender As Object, e As MouseEventArgs) Handles txt流月_傳統法.MouseDown
		If e.Button = MouseButtons.Middle Then txt流月_傳統法.SendToBack()
	End Sub
	Private Sub txt流月_壬癸己_MouseDown(sender As Object, e As MouseEventArgs) Handles txt流月_壬癸己.MouseDown
		If e.Button = MouseButtons.Middle Then txt流月_壬癸己.SendToBack()
	End Sub
	Private Sub txt流月_六爻法_MouseDown(sender As Object, e As MouseEventArgs) Handles txt流年_六爻法.MouseDown
		If e.Button = MouseButtons.Middle Then txt流年_六爻法.SendToBack()
	End Sub
	'==================================================================
	' Copilot 寫出的同步捲動code, 有點問題，watting..
	' 宣告API函數
	<System.Runtime.InteropServices.DllImport("user32.dll")>
	Private Shared Function GetScrollPos(hWnd As IntPtr, nBar As Integer) As Integer
	End Function

	<System.Runtime.InteropServices.DllImport("user32.dll")>
	Private Shared Function SetScrollPos(hWnd As IntPtr, nBar As Integer, nPos As Integer, bRedraw As Boolean) As Integer
	End Function

	<System.Runtime.InteropServices.DllImport("user32.dll")>
	Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As Integer, lParam As Integer) As Integer
	End Function

	Private Const SBS_VERT As Integer = 1
	Private Const WM_VSCROLL As Integer = &H115
	'----------------------------------------------------
	Private Sub TextBox1_Scroll(sender As Object, e As EventArgs) 'Handles txt流年_傳統法.MouseWheel, txt流年_傳統法.KeyDown
		SyncScroll(txt流年_傳統法, txt流年_壬癸己)
	End Sub
	Private Sub TextBox2_Scroll(sender As Object, e As EventArgs) 'Handles txt流年_壬癸己.MouseWheel, txt流年_壬癸己.KeyDown
		SyncScroll(txt流年_壬癸己, txt流年_傳統法)
	End Sub

	Private Sub SyncScroll(source As TextBox, target As TextBox)
		' 獲取源TextBox的滾動信息
		Dim info As Integer = GetScrollPos(source.Handle, SBS_VERT)
		' 設置目標TextBox的滾動信息
		SetScrollPos(target.Handle, SBS_VERT, info, True)
		' 發送滾動消息以同步滾動
		SendMessage(target.Handle, WM_VSCROLL, info, 0)
		target.Invalidate()
	End Sub
	'================================================================
	Private Function uf天乙貴人(sG As String, sZ As String) As String
		Dim s = "ＯＯ"
		Select Case sG
			Case "甲"， "戊"， "庚" : If sZ.Equals("丑") OrElse sZ.Equals("未") Then s = "乙贵"
			Case "乙"， "己" : If sZ.Equals("子") OrElse sZ.Equals("申") Then s = "乙贵"
			Case "丙"， "丁" : If sZ.Equals("亥") OrElse sZ.Equals("酉") Then s = "乙贵"
			Case "壬"， "癸" : If sZ.Equals("卯") OrElse sZ.Equals("巳") Then s = "乙贵"
			Case "辛" : If sZ.Equals("寅") OrElse sZ.Equals("午") Then s = "乙贵"
		End Select
		Return s
	End Function

	Private Function uf胎元(sMgz As String) As String
		Dim sG = Mid(sMgz, 1, 1), sZ = Mid(sMgz, 2, 1), s1 = "", s2 = ""
		Select Case sG
			Case "甲" : s1 = "乙" : Case "乙" : s1 = "丙" : Case "丙" : s1 = "丁" : Case "丁" : s1 = "戊" : Case "戊" : s1 = "己"
			Case "己" : s1 = "庚" : Case "庚" : s1 = "辛" : Case "辛" : s1 = "壬" : Case "壬" : s1 = "癸" : Case "癸" : s1 = "甲"
		End Select

		Select Case sZ
			Case "子" : s2 = "卯" : Case "丑" : s2 = "辰"
			Case "寅" : s2 = "巳" : Case "卯" : s2 = "午"
			Case "辰" : s2 = "未" : Case "巳" : s2 = "申"
			Case "午" : s2 = "酉" : Case "未" : s2 = "戌"
			Case "申" : s2 = "亥" : Case "酉" : s2 = "子"
			Case "戌" : s2 = "丑" : Case "亥" : s2 = "寅"
		End Select

		Return s1 + s2
	End Function
	Private Function uf命宮(sYg As String, sMz As String, sHz As String) As String
		Dim s1 = "", s2 = ""
		'                        寅月  卯月  辰    巳    午    未    申    酉   戌     亥    子    丑
		Dim s子_() As String = {"卯", "寅", "丑", "子", "亥", "戌", "酉", "申", "未", "午", "巳", "辰"}
		Dim s丑_() As String = {"寅", "丑", "子", "亥", "戌", "酉", "申", "未", "午", "巳", "辰", "卯"}
		Dim s寅_() As String = {"丑", "子", "亥", "戌", "酉", "申", "未", "午", "巳", "辰", "卯", "寅"}
		Dim s卯_() As String = {"子", "亥", "戌", "酉", "申", "未", "午", "巳", "辰", "卯", "寅", "丑"}
		Dim s辰_() As String = {"亥", "戌", "酉", "申", "未", "午", "巳", "辰", "卯", "寅", "丑", "子"}
		Dim s巳_() As String = {"戌", "酉", "申", "未", "午", "巳", "辰", "卯", "寅", "丑", "子", "亥"}
		Dim s午_() As String = {"酉", "申", "未", "午", "巳", "辰", "卯", "寅", "丑", "子", "亥", "戌"}
		Dim s未_() As String = {"申", "未", "午", "巳", "辰", "卯", "寅", "丑", "子", "亥", "戌", "酉"}
		Dim s申_() As String = {"未", "午", "巳", "辰", "卯", "寅", "丑", "子", "亥", "戌", "酉", "申"}
		Dim s酉_() As String = {"午", "巳", "辰", "卯", "寅", "丑", "子", "亥", "戌", "酉", "申", "未"}
		Dim s戌_() As String = {"巳", "辰", "卯", "寅", "丑", "子", "亥", "戌", "酉", "申", "未", "午"}
		Dim s亥_() As String = {"辰", "卯", "寅", "丑", "子", "亥", "戌", "酉", "申", "未", "午", "巳"}
		Dim iM, iH As Int32

		Select Case sMz
			Case "寅" : iM = 0 : Case "卯" : iM = 1 : Case "辰" : iM = 2
			Case "巳" : iM = 3 : Case "午" : iM = 4 : Case "未" : iM = 5
			Case "申" : iM = 6 : Case "酉" : iM = 7 : Case "戌" : iM = 8
			Case "亥" : iM = 9 : Case "子" : iM = 10 : Case "丑" : iM = 11
		End Select

		Select Case sHz
			Case "亥" : s2 = s亥_(iM) : Case "子" : s2 = s子_(iM) : Case "丑" : s2 = s辰_(iM)
			Case "寅" : s2 = s寅_(iM) : Case "卯" : s2 = s卯_(iM) : Case "辰" : s2 = s辰_(iM)
			Case "巳" : s2 = s巳_(iM) : Case "午" : s2 = s午_(iM) : Case "未" : s2 = s未_(iM)
			Case "申" : s2 = s申_(iM) : Case "酉" : s2 = s酉_(iM) : Case "戌" : s2 = s戌_(iM)
		End Select

		'                         子    丑    寅    卯    辰    巳    午    未    申    酉   戌    亥
		Dim s甲己_() As String = {"丙", "丁", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸", "甲", "乙"}
		Dim s乙庚_() As String = {"戊", "己", "戊", "己", "庚", "辛", "壬", "癸", "甲", "乙", "丙", "丁"}
		Dim s丙辛_() As String = {"庚", "辛", "庚", "辛", "壬", "癸", "甲", "乙", "丙", "丁", "戊", "己"}
		Dim s丁壬_() As String = {"壬", "癸", "壬", "癸", "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛"}
		Dim s戊癸_() As String = {"甲", "乙", "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"}

		Select Case s2
			Case "寅" : iH = 2 : Case "卯" : iH = 3 : Case "辰" : iH = 4
			Case "巳" : iH = 5 : Case "午" : iH = 6 : Case "未" : iH = 7
			Case "申" : iH = 8 : Case "酉" : iH = 9 : Case "戌" : iH = 10
			Case "亥" : iH = 11 : Case "子" : iH = 0 : Case "丑" : iH = 1
		End Select

		Select Case sYg
			Case "甲", "己" : s1 = s甲己_(iH)
			Case "乙", "庚" : s1 = s乙庚_(iH)
			Case "丙", "辛" : s1 = s丙辛_(iH)
			Case "丁", "壬" : s1 = s丁壬_(iH)
			Case "戊", "癸" : s1 = s戊癸_(iH)
		End Select

		Return s1 + s2
	End Function
	Private Function uf身宮(sYg As String, sMz As String, sHz As String) As String
		Dim s1 = "", s2 = ""
		'                        寅月  卯月  辰    巳    午    未    申    酉   戌     亥    子    丑
		Dim s子_() As String = {"卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥", "子", "丑", "寅"}
		Dim s丑_() As String = {"辰", "巳", "午", "未", "申", "酉", "戌", "亥", "子", "丑", "寅", "卯"}
		Dim s寅_() As String = {"巳", "午", "未", "申", "酉", "戌", "亥", "子", "丑", "寅", "卯", "辰"}
		Dim s卯_() As String = {"午", "未", "申", "酉", "戌", "亥", "子", "丑", "寅", "卯", "辰", "巳"}
		Dim s辰_() As String = {"未", "申", "酉", "戌", "亥", "子", "丑", "寅", "卯", "辰", "巳", "午"}
		Dim s巳_() As String = {"申", "酉", "戌", "亥", "子", "丑", "寅", "卯", "辰", "巳", "午", "未"}
		Dim s午_() As String = {"酉", "戌", "亥", "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申"}
		Dim s未_() As String = {"戌", "亥", "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉"}
		Dim s申_() As String = {"亥", "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌"}
		Dim s酉_() As String = {"子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"}
		Dim s戌_() As String = {"丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥", "子"}
		Dim s亥_() As String = {"寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥", "子", "丑"}
		Dim iM, iH As Int32

		Select Case sMz
			Case "寅" : iM = 0 : Case "卯" : iM = 1 : Case "辰" : iM = 2
			Case "巳" : iM = 3 : Case "午" : iM = 4 : Case "未" : iM = 5
			Case "申" : iM = 6 : Case "酉" : iM = 7 : Case "戌" : iM = 8
			Case "亥" : iM = 9 : Case "子" : iM = 10 : Case "丑" : iM = 11
		End Select

		Select Case sHz
			Case "亥" : s2 = s亥_(iM) : Case "子" : s2 = s子_(iM) : Case "丑" : s2 = s辰_(iM)
			Case "寅" : s2 = s寅_(iM) : Case "卯" : s2 = s卯_(iM) : Case "辰" : s2 = s辰_(iM)
			Case "巳" : s2 = s巳_(iM) : Case "午" : s2 = s午_(iM) : Case "未" : s2 = s未_(iM)
			Case "申" : s2 = s申_(iM) : Case "酉" : s2 = s酉_(iM) : Case "戌" : s2 = s戌_(iM)
		End Select

		'                         子    丑    寅    卯                                      戌    亥
		Dim s甲己_() As String = {"丙", "丁", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸", "甲", "乙"}
		Dim s乙庚_() As String = {"戊", "己", "戊", "己", "庚", "辛", "壬", "癸", "甲", "乙", "丙", "丁"}
		Dim s丙辛_() As String = {"庚", "辛", "庚", "辛", "壬", "癸", "甲", "乙", "丙", "丁", "戊", "己"}
		Dim s丁壬_() As String = {"壬", "癸", "壬", "癸", "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛"}
		Dim s戊癸_() As String = {"甲", "乙", "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"}

		Select Case s2
			Case "寅" : iH = 2 : Case "卯" : iH = 3 : Case "辰" : iH = 4
			Case "巳" : iH = 5 : Case "午" : iH = 6 : Case "未" : iH = 7
			Case "申" : iH = 8 : Case "酉" : iH = 9 : Case "戌" : iH = 10
			Case "亥" : iH = 11 : Case "子" : iH = 0 : Case "丑" : iH = 1
		End Select

		Select Case sYg
			Case "甲", "己" : s1 = s甲己_(iH)
			Case "乙", "庚" : s1 = s乙庚_(iH)
			Case "丙", "辛" : s1 = s丙辛_(iH)
			Case "丁", "壬" : s1 = s丁壬_(iH)
			Case "戊", "癸" : s1 = s戊癸_(iH)
		End Select

		Return s1 + s2
	End Function
	'================================================================
	Dim _dic串宮_12神, _dic串宮_干2支_古法, _dic串宮_干2支_六行 As New Dictionary(Of String, String)
	Private Sub us串宮压运(sYz As String)
		'Dim s12神_ = {"太岁", "青龙", "丧门", "六合", "官符", "小耗", "大耗", "朱雀", "白虎", "贵神", "吊客", "病符"}
		'               1     2    3     4     5    6     7     8    9     10   11    12
		'Dim s12神_ = {"岁", "龙", "丧", "合", "官", "耒", "耗", "雀", "虎", "贵", "吊", "病"}
		' 徐調整       申     酉   戌    亥    子    丑    寅    卯    辰     巳    午    未
		Dim s12神_ = {"岁", "丧", "雀", "虎", "龙", "耒", "耗", "病", "贵", "合", "官", "吊"}
		Dim s子_() As String = {"子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"}
		Dim s丑_() As String = {"丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥", "子"}
		Dim s寅_() As String = {"寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥", "子", "丑"}
		Dim s卯_() As String = {"卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥", "子", "丑", "寅"}
		Dim s辰_() As String = {"辰", "巳", "午", "未", "申", "酉", "戌", "亥", "子", "丑", "寅", "卯"}
		Dim s巳_() As String = {"巳", "午", "未", "申", "酉", "戌", "亥", "子", "丑", "寅", "卯", "辰"}
		Dim s午_() As String = {"午", "未", "申", "酉", "戌", "亥", "子", "丑", "寅", "卯", "辰", "巳"}
		Dim s未_() As String = {"未", "申", "酉", "戌", "亥", "子", "丑", "寅", "卯", "辰", "巳", "午"}
		Dim s申_() As String = {"申", "酉", "戌", "亥", "子", "丑", "寅", "卯", "辰", "巳", "午", "未"}
		Dim s酉_() As String = {"酉", "戌", "亥", "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申"}
		Dim s戌_() As String = {"戌", "亥", "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉"}
		Dim s亥_() As String = {"亥", "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌"}
		Dim s_() As String
		'---------------------------------------------
		Select Case sYz
			Case "寅" : s_ = s寅_ : Case "卯" : s_ = s卯_ : Case "辰" : s_ = s辰_
			Case "巳" : s_ = s巳_ : Case "午" : s_ = s午_ : Case "未" : s_ = s未_
			Case "申" : s_ = s申_ : Case "酉" : s_ = s酉_ : Case "戌" : s_ = s戌_
			Case "亥" : s_ = s亥_ : Case "子" : s_ = s子_ : Case "丑" : s_ = s丑_
		End Select

		_dic串宮_12神.Clear()
		For i = 0 To 11 : _dic串宮_12神.Add(s_(i), s12神_(i)) : Next
		'---------------------------------------------

	End Sub
	'================================================================
	Private Function uf命局缺的字_干支_五行(sDg As String, sGZ As String) As String
		Dim s = "", s1 = "天干缺："
		s = "甲" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_五行(sDg, s) + "，"
		s = "乙" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_五行(sDg, s) + "，"
		s = "丙" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_五行(sDg, s) + "，"
		s = "丁" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_五行(sDg, s) + "，"
		s = "戊" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_五行(sDg, s) + "，"
		s = "己" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_五行(sDg, s) + "，"
		s = "庚" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_五行(sDg, s) + "，"
		s = "辛" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_五行(sDg, s) + "，"
		s = "壬" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_五行(sDg, s) + "，"
		s = "癸" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_五行(sDg, s) + "，"

		If False Then ' 無此情形
			s1 += vbNL + "十神缺："
			s = "枭" : If Not sGZ.Contains(s) Then s1 += s
			s = "印" : If Not sGZ.Contains(s) Then s1 += s
			s = "比" : If Not sGZ.Contains(s) Then s1 += s
			s = "劫" : If Not sGZ.Contains(s) Then s1 += s
			s = "食" : If Not sGZ.Contains(s) Then s1 += s
			s = "伤" : If Not sGZ.Contains(s) Then s1 += s
			s = "贝" : If Not sGZ.Contains(s) Then s1 += s
			s = "财" : If Not sGZ.Contains(s) Then s1 += s
			s = "杀" : If Not sGZ.Contains(s) Then s1 += s
			s = "官" : If Not sGZ.Contains(s) Then s1 += s
		End If

		Return Mid(s1, 1, s1.Length - 1) + vbNL
		'Return s1 + vbNL
	End Function
	Private Function uf命局缺的字_干支_六行(sDg As String, sGZ As String) As String
		Dim s = "", s1 = "天干缺："
		s = "甲" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_六行_奇隅(sDg, s) + "，"
		s = "乙" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_六行_奇隅(sDg, s) + "，"
		s = "丙" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_六行_奇隅(sDg, s) + "，"
		s = "丁" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_六行_奇隅(sDg, s) + "，"
		s = "戊" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_六行_奇隅(sDg, s) + "，"
		s = "己" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_六行_奇隅(sDg, s) + "，"
		s = "庚" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_六行_奇隅(sDg, s) + "，"
		s = "辛" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_六行_奇隅(sDg, s) + "，"
		s = "壬" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_六行_奇隅(sDg, s) + "，"
		s = "癸" : If Not sGZ.Contains(s) Then s1 += s + uf_印綬_干vs干_六行_奇隅(sDg, s) + "，"

		s1 = Mid(s1, 1, s1.Length - 1) + "+"
		s = "枭" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "印" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "比" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "劫" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "食" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "伤" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "贝" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "财" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "奴" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "迁" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "杀" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "官" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s

		'Return Mid(s1, 1, s1.Length - 1) + vbNL
		Return s1 + vbNL
	End Function
	Private Function uf命局缺的字_支支_六行(sDg As String, sGZ As String) As String
		Dim s = "", s1 = "天支缺："
		s = "子" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "丑" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "寅" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "卯" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "辰" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "巳" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "午" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "未" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "申" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "酉" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "戌" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"
		s = "亥" : If Not sGZ.Contains(s) Then s1 += s + uf六親_支vs支_六行_丑辰元(sDg, s) + "，"

		s1 = Mid(s1, 1, s1.Length - 1) + "+"
		s = "枭" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "印" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "比" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "劫" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "食" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "伤" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "贝" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "财" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "奴" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "迁" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "杀" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s
		s = "官" : If Not sGZ.Contains(s) AndAlso Not s1.Contains(s) Then s1 += s

		'Return Mid(s1, 1, s1.Length - 1) + vbNL
		Return s1 + vbNL
	End Function
	'================================================================
	Private Function ufDD組合_五行(sYMDH As String, sGZ As String) As String
		Dim s = "", s1 = "" ' 註：sYMDH=Y,M,D,H

		s1 = sYMDH + sGZ
		Select Case sGZ ' 自坐天乙贵人
			Case "丁酉", "丁亥", "癸卯", "癸巳" : s += s1 + "=ｕ天贵" + vbNL
		End Select
		Select Case sGZ
			Case "丙子", "丁丑", "戊寅", "辛卯", "壬辰", "癸巳" : s += s1 + "=(男)阴阳日，首婚90%离婚" + vbNL
			Case "丙午", "丁未", "戊申", "辛酉", "壬戌", "癸亥" : s += s1 + "=(女)阴阳日，首婚90%离婚" + vbNL
		End Select
		Select Case sGZ
			Case "甲辰", "甲戌" : s += s1 + "=克配偶日" + vbNL
		End Select
		Select Case sGZ
			Case "丙午", "丁未", "戊子", "戊午", "己丑", "己未" : s += s1 + "=六秀日" + vbNL
		End Select
		Select Case sGZ
			Case "甲寅", "乙卯", "丁未", "戊戌", "己未", "庚申", "辛酉", "癸丑" : s += s1 + "=(男)八专日，重情欲" + vbNL
		End Select
		Select Case sGZ
			Case "甲寅", "丁酉", "戊子", "戊午", "己卯", "己酉", "辛卯", "辛酉", "壬子", "壬午" : s += s1 + "=(女)九丑日，重情欲，名声差" + vbNL
		End Select
		Select Case sGZ
			Case "甲午", "丙寅", "丁未", "戊辰", "庚戌", "辛酉", "壬子" : s += s1 + "=红艳煞，重情欲" + vbNL
		End Select

		Select Case sGZ
			Case "丁卯" : s += s1 + "=丁u卯[乙枭]，男女婚姻皆不美满" + vbNL
			Case "癸酉" : s += s1 + "=癸u酉[辛枭]，男女婚姻皆不美满" + vbNL
		End Select

		Select Case sGZ
			' 甲u寅禄 NA
			' 庚u申禄 NA
			Case "丙午" : s += sYMDH + "丙u午[丁劫刃+己]，眼光高，克妻" + vbNL ' u四正支
			Case "戊午" : s += sYMDH + "戊u午[丁劫刃+己]，眼光高，克妻" + vbNL
			Case "壬子" : s += sYMDH + "壬u子[癸劫刃]，眼光高，克妻" + vbNL
			' 乙u卯禄 NA
			' 辛u酉禄 NA
			Case "丁巳" : s += sYMDH + "丁u巳[丙劫刃+庚+戊]，克妻" + vbNL ' 不u四正支
			Case "己巳" : s += sYMDH + "己u巳[丙+庚+戊劫刃]，克妻" + vbNL
			Case "癸亥" : s += sYMDH + "癸u亥[壬劫刃+甲]，克妻" + vbNL
		End Select

		Select Case sGZ
			Case "庚辰" : s += sYMDH + "庚u辰[戊枭+乙财+癸伤]=魁罡日，绝夫罡" + vbNL
			Case "庚戌" : s += sYMDH + "庚u戌[戊枭+辛劫+丁官]=魁罡日，绝夫罡" + vbNL
			Case "壬辰" : s += sYMDH + "壬u辰[戊杀+乙伤+癸劫]=魁罡日，绝妻罡" + vbNL
			Case "戊戌" : s += sYMDH + "戊u戌[戊比+辛伤+丁印]=魁罡日，绝妻罡" + vbNL
		End Select

		If Not _is男 Then
			Select Case sGZ
				Case "甲午" : s += sYMDH + "甲u午[丁伤+己]，克夫" + vbNL
				Case "乙巳" : s += sYMDH + "乙u巳[丙伤+庚+戊]，克夫" + vbNL
				Case "丙戊" : s += sYMDH + "丙u辰[戊食+乙+癸]，克夫" + vbNL
				Case "丙戌" : s += sYMDH + "丙u戌[戊食+辛+酉]，克夫" + vbNL
				Case "丁丑" : s += sYMDH + "丁u丑[己食+癸+辛]，克夫" + vbNL
				Case "丁未" : s += sYMDH + "丁u未[己食+丁比+乙]，克夫" + vbNL
				Case "戊申" : s += sYMDH + "戊u申[庚食+壬+戊比]，克夫" + vbNL
				Case "己酉" : s += sYMDH + "己u酉[辛食]，克夫" + vbNL    ' 庚u酉刃  六行法
				Case "庚子" : s += sYMDH + "庚u子[癸伤]，克夫最重" + vbNL ' 辛u子[癸食+壬伤]
				Case "辛亥" : s += sYMDH + "辛u亥[壬伤+甲]，克夫" + vbNL  ' 壬u亥禄
				Case "壬寅" : s += sYMDH + "壬u寅[甲食+丙+戊]，克夫" + vbNL' 癸u寅[甲财+丙迁+戊官]
				Case "癸卯" : s += sYMDH + "癸u卯[乙食]，克夫" + vbNL     ' 己u卯[乙食]
			End Select

			Select Case sGZ
				Case "甲子" : s += sYMDH + "甲u子[癸印]，孤孀禄马，克夫" + vbNL
				Case "庚午" : s += sYMDH + "庚u午[丁+己印]，孤孀禄马，克夫" + vbNL
				Case "丙寅" : s += sYMDH + "丙u寅[甲枭+丙+戊]，孤孀禄马，克夫" + vbNL
				Case "壬申" : s += sYMDH + "壬u申[庚枭+壬+戊]，孤孀禄马，克夫" + vbNL
			End Select
		End If

		Select Case sGZ
			Case "甲辰", "乙亥", "丙辰", "丁酉", "戊午", "庚戌", "庚寅", "辛亥", "壬寅", "癸未" : s += s1 + "=十灵日" + vbNL
		End Select
		Select Case sGZ
			Case "甲辰", "乙巳", "丙申", "丁亥", "戊戌", "己丑", "庚辰", "辛巳", "壬申", "癸亥" : s += s1 + "=十惡大敗日" + vbNL
		End Select
		Select Case sGZ
			Case "甲寅", "乙亥", "丙子", "丙戊", "丁酉", "戊申", "己未", "庚午", "辛巳", "壬辰", "癸卯" : s += s1 + "=福星贵" + vbNL
		End Select
		Select Case sGZ
			Case "甲寅", "丙辰", "戊辰", "庚戌", "壬戌" : s += s1 + "=日德日" + vbNL
		End Select

		Return s
	End Function
	Private Function ufHH組合_五行(sYMDH As String, sGZ As String) As String
		Dim s = "", s1 = ""
		Select Case sGZ
			Case "庚子", "戊子", "甲子", "丙子", "壬子" : s += sGZ + "=斩子剑" + vbNL
		End Select
		Return s
	End Function
	Private Function ufDH組合_五行(sDgz As String, sHgz As String) As String
		Dim s = "", s1 = ""

		Select Case sDgz
			Case "丙子" : If sHgz.Equals("辛卯") Then s += "D" + sDgz + "H" + sHgz + "=滚浪桃花" + vbNL
			Case "甲子" : If sHgz.Equals("庚午") Then s += "D" + sDgz + "H" + sHgz + "=裸体桃花" + vbNL
			Case "庚午" : If sHgz.Equals("甲子") Then s += "D" + sDgz + "H" + sHgz + "=裸体桃花" + vbNL
		End Select
		Return s
	End Function
	Private Function uf_MzHz組和_五行(sMz As String, sHz As String) As String
		Dim s = "", s1 = "", s2 = ""
		s1 = "命犯绝梁，伤残"
		Select Case True
			Case sMz.Equals("寅") AndAlso sHz.Equals("巳") : s = "M寅+H巳=" + s1
			Case sMz.Equals("卯") AndAlso sHz.Equals("酉") : s = "M卯+H酉=" + s1
			Case sMz.Equals("辰") AndAlso sHz.Equals("丑") : s = "M辰+H丑=" + s1
			Case sMz.Equals("巳") AndAlso sHz.Equals("辰") : s = "M巳+H辰=" + s1
			Case sMz.Equals("午") AndAlso sHz.Equals("申") : s = "M午+H申=" + s1
			Case sMz.Equals("未") AndAlso sHz.Equals("子") : s = "M未+H子=" + s1
			Case sMz.Equals("申") AndAlso sHz.Equals("卯") : s = "M申+H卯=" + s1
			Case sMz.Equals("酉") AndAlso sHz.Equals("未") : s = "M酉+H未=" + s1
			Case sMz.Equals("戌") AndAlso sHz.Equals("亥") : s = "M戌+H亥=" + s1
			Case sMz.Equals("亥") AndAlso sHz.Equals("寅") : s = "M亥+H寅=" + s1
			Case sMz.Equals("子") AndAlso sHz.Equals("午") : s = "M子+H午=" + s1
			Case sMz.Equals("丑") AndAlso sHz.Equals("戌") : s = "M丑+H戌=" + s1
		End Select
		Return s
	End Function
End Class