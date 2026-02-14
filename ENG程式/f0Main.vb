Option Explicit On : Option Infer On : Option Strict On
Imports System.IO, System.Text

Public Class f0Main
	Private Sub frm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
		' 因 dt 要寫回 db，要 Form.Close()
	End Sub
	Private Sub frm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
		Me.Dispose() : GC.Collect() '(1)將物件標記為可被回收 (2)釋放記憶體
	End Sub

	Private Sub frm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		m1.us_營幕位置()
		AddHandler __tmr井字.Tick, AddressOf m1.tmr井字_Tick
	End Sub

	Private Sub btn八字排盤1_Click(sender As Object, e As EventArgs) Handles btn八字排盤1.Click
		f31八字排盤.Show()
	End Sub

	Private Sub btnWebster字典mp3_List_Click(sender As Object, e As EventArgs)
		Dim sP = __sPath_db英文 + "發音\Webster\"
		Dim sw1 As New StreamWriter(sP + "0目前mp3_List.txt")

		Dim di As New DirectoryInfo(sP), fi_() As FileInfo
		Dim sShortName As String

		fi_ = di.GetFiles("*", SearchOption.TopDirectoryOnly)
		For Each f As FileInfo In fi_
			sShortName = f.Name
			If Not sShortName.Contains(".mp3") Then Continue For
			sw1.WriteLine(sShortName.Replace(".mp3", "")) ' word.mp3 => word
		Next
		sw1.Close()
	End Sub

	Private Sub btnCambridge字典mp3_List_Click(sender As Object, e As EventArgs)
		Dim sP = __sPath_db英文 + "發音\Cambridge\"
		Dim sw1 As New StreamWriter(sP + "0目前mp3_List.txt")

		Dim di As New DirectoryInfo(sP), fi_() As FileInfo
		Dim sShortName As String

		fi_ = di.GetFiles("*", SearchOption.TopDirectoryOnly)
		For Each f As FileInfo In fi_
			sShortName = f.Name
			If Not sShortName.Contains(".mp3") Then Continue For
			sw1.WriteLine(sShortName.Replace(".mp3", "")) ' word.mp3 => word
		Next
		sw1.Close()
	End Sub

	Private Sub btnAnki_短句_List_Click(sender As Object, e As EventArgs)
		Dim sP = __sPath_db英文 + "發音\Anki\"
		Dim sw1 As New StreamWriter(sP + "0目前mp3_List.txt")

		Dim di As New DirectoryInfo(sP), fi_() As FileInfo
		Dim sShortName As String

		fi_ = di.GetFiles("*", SearchOption.TopDirectoryOnly)
		For Each f As FileInfo In fi_
			sShortName = f.Name
			If Not sShortName.Contains(".mp3") Then Continue For
			sw1.WriteLine(sShortName.Replace(".mp3", "")) ' word.mp3 => word
		Next
		sw1.Close()
	End Sub

	Private Sub btn中文字_Click(sender As Object, e As EventArgs)
		Dim s, sALL, sF中字, sF符號, s網頁, sOut As String, n中字0, n中字1, n符號0, n符號1 As Int32
		Dim dic中字, dic符號 As New Dictionary(Of String, String)
		Dim sr1 As StreamReader

		sF中字 = "E:\谷谷碟\List_中字.txt"
		sr1 = New StreamReader(sF中字, Encoding.Default) : sALL = sr1.ReadToEnd.Trim
		For i = 1 To sALL.Length
			s = Mid(sALL, i, 1) : If Not dic中字.ContainsKey(s) Then dic中字.Add(s, Nothing)
		Next : sr1.Close()

		sF符號 = "E:\谷谷碟\List_符號.txt"
		sr1 = New StreamReader(sF符號, Encoding.Default) : sALL = sr1.ReadToEnd.Trim
		For i = 1 To sALL.Length
			s = Mid(sALL, i, 1) : If Not dic符號.ContainsKey(s) Then dic符號.Add(s, Nothing)
		Next : sr1.Close()
		'---------------------------------------------------------
		n中字1 = dic中字.Count : n符號1 = dic符號.Count

		s網頁 = Clipboard.GetText.Trim : If s網頁.Length = 0 Then Exit Sub
		For i = 1 To s網頁.Length
			s = Mid(s網頁, i, 1).Trim
			If s.Length = 0 Then Continue For
			If dic符號.ContainsKey(s) Then Continue For
			If dic中字.ContainsKey(s) Then Continue For Else dic中字.Add(s, Nothing)

			'If Not dic符號.ContainsKey(s) Then dic符號.Add(s, Nothing)
		Next
		n中字0 = dic中字.Count : n符號0 = dic符號.Count
		'---------------------------------------------------------
		If n中字1 < n中字0 Then
			Dim sw1 As New StreamWriter(sF中字, False, Encoding.Default)
			Dim lst1 As List(Of String)
			lst1 = dic中字.Keys.ToList : lst1.Sort()

			sOut = "" : For Each it In lst1 : sOut += it : Next
			sw1.WriteLine(sOut) : sw1.Close()
			Console.WriteLine("新增字數=" + (n中字1 - n中字0).ToString)
		End If

		If False Then
			Dim sw1 As New StreamWriter(sF符號, False, Encoding.Default)
			Dim lst1 As List(Of String)
			lst1 = dic符號.Keys.ToList : lst1.Sort()

			sOut = "" : For Each it In lst1 : sOut += it : Next
			sw1.WriteLine(sOut) : sw1.Close()
			Console.WriteLine("新增符號=" + (n符號1 - n符號0).ToString)
		End If
	End Sub
End Class