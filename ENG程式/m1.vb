Option Explicit On : Option Infer On : Option Strict On
Imports System.Data.OleDb
Imports System.Runtime.CompilerServices ' 擴充功能

Module m1
  'MessageBox(參數縮寫)
  Public Const mbb_O As MessageBoxButtons = MessageBoxButtons.OK,
               mbb_YN As MessageBoxButtons = MessageBoxButtons.YesNo,
               mbb_OC As MessageBoxButtons = MessageBoxButtons.OKCancel,
               mbb_YNC As MessageBoxButtons = MessageBoxButtons.YesNoCancel,
               mbi_I As MessageBoxIcon = MessageBoxIcon.Information,
               mbi_Q As MessageBoxIcon = MessageBoxIcon.Question,
               mbi_S As MessageBoxIcon = MessageBoxIcon.Stop,
               mbi_W As MessageBoxIcon = MessageBoxIcon.Warning

  Public Const __iRowTH = 25 ' dgv.RowTemplate.Height
  Public Const __iHH = 25
  Public __iVsbW As Int32 = SystemInformation.VerticalScrollBarWidth ' 垂直捲軸寬度
  Public __iHsbH As Int32 = SystemInformation.HorizontalScrollBarHeight ' 水平捲軸高度

  Public __cntAcc8 As OleDbConnection
  Public Const cm_Bin As CompareMethod = CompareMethod.Binary
  Public Const TT = True, FF = False
  Public Const __sPath_db命命 = "D:\工作碟\sky資料庫\db命命\"
  Public Const __sPath_db英文 = "D:\工作碟\sky資料庫\db英文\"
  Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Int32) As Int32
  '==================================================================
  Public Sub us_ErrMsg3(ex As Exception, Optional Title As Object = Nothing)
    ' 徐20266.2.4：copilot 改的，但一樣沒法顯示 Form+sub() 名稱
    Dim st As New StackTrace(ex, True)
    ' 嘗試找第一個「不是在錯誤處理工具類」的 frame（避免永遠顯示 us_ErrMsg2）
    Dim sf As StackFrame = Nothing
    For i = 0 To st.FrameCount - 1
      Dim frm = st.GetFrame(i)
      Dim mb = frm?.GetMethod()
      Dim t = mb?.DeclaringType
      If t IsNot Nothing AndAlso t IsNot GetType(m1) Then
        sf = frm
        Exit For
      End If
    Next
    If sf Is Nothing AndAlso st.FrameCount > 0 Then sf = st.GetFrame(0)

    Dim m = sf.GetMethod()
    Dim sfrmName = m?.ReflectedType?.Name
    Dim sSubName = m?.Name

    Dim msg = $"{ex.Message}(程式：@{sfrmName}﹒{sSubName})" ' Null 時採用 "錯誤"
    Dim caption = If(Title?.ToString(), "錯誤")
    MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Error)
  End Sub


  Public Sub us_ErrMsg2(ex As Exception, Optional Title As Object = Nothing)
    ' 徐20266.2.4：用C#的方法，改寫成VB
    Dim st = New StackTrace(ex, True) ' 用例外物件建立 StackTrace
    Dim sf = st.GetFrame(0)
    Dim mb = sf.GetMethod()
    Dim sfrmName = mb?.ReflectedType?.Name
    Dim sSubName = mb?.Name

    Dim msg = $"{ex.Message}(程式：@{sfrmName}﹒{sSubName})" ' Null 時採用 "錯誤"
    Dim caption = If(Title?.ToString(), "錯誤")
    MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Error)
  End Sub

  Public Sub us_ErrMsg(ByRef sf0 As StackFrame, sErr As String, Optional Title As Object = Nothing)
    Dim sfrmName = sf0.GetMethod.ReflectedType.Name
    Dim sSubName = sf0.GetMethod.Name
    sErr += "(程式：@" + sfrmName + "﹒" + sSubName + ")"
    MsgBox(sErr, MsgBoxStyle.Critical, Title)
  End Sub

  Public Sub us_效能(ByRef sf0 As StackFrame, ByRef sw As Stopwatch)
    Dim sfrmName = sf0.GetMethod.ReflectedType.Name
    Dim sSubName = sf0.GetMethod.Name
    sw.Stop()
    Console.WriteLine("效能：@" + sfrmName + "﹒" + sSubName + ", ms=" + sw.ElapsedMilliseconds.ToString)
  End Sub
  '==================================================================
  Public Sub us_webPageLoading(ByRef wb As WebBrowser) ' 等待網頁讀取完成
    'Static iCnt As Int32 = 0
    Try
      Do Until wb.ReadyState = WebBrowserReadyState.Complete
        'iCnt += 1
        Application.DoEvents()
        'Console.Write(wb.Name + "=" + iCnt.ToString + vbTab)
      Loop
    Catch ex As Exception
      ' 不能省try，在關閉程式時，本sub 會出現 error，不理它
    End Try
  End Sub
  '===================================================================
  Public Sub tabCtl_DrawItem(ByRef tc As TabControl, ByRef e As DrawItemEventArgs)
    Dim g = e.Graphics, eIdx = e.Index, eBnd = e.Bounds
    Dim s = tc.TabPages(eIdx).Text
    Dim sf As New StringFormat : sf.LineAlignment = StringAlignment.Center : sf.Alignment = StringAlignment.Center

    If eIdx = tc.SelectedIndex Then ' 選取頁籤
      g.FillRectangle(Brushes.Pink, eBnd)
      g.DrawString(s, New Font(__sfnt_DGV, 11.0F), New SolidBrush(Color.Black), eBnd, sf)
    Else
      'e.DrawBackground()
      g.FillRectangle(Brushes.DimGray, eBnd)
      g.DrawString(s, New Font(__sfnt_DGV, 11.0F), New SolidBrush(Color.White), eBnd, sf)
    End If
    g.Dispose()
    'Console.WriteLine("tabCtl_DrawItem")
  End Sub
  '===================================================================
  Public __tmr井字 As New Timer
  Public __hc_Selfrm, __hc_f1Eng As Int32

  Dim _nScr As Int32
  Dim _wa() As Rectangle
  Dim _iwaL(), _iwaT(), _iwaW(), _iwaH() As Int32
  Dim _iwaW_10等分(), _iwaW_5等分(), _iwaW_中間() As Int32
  Dim _iwaH_10等分(), _iwaH_5等分(), _iwaH_中間() As Int32
  Public Sub us_營幕位置()
    _nScr = Screen.AllScreens.Count - 1
    ReDim _wa(_nScr)
    ReDim _iwaL(_nScr), _iwaT(_nScr), _iwaW(_nScr), _iwaH(_nScr)
    ReDim _iwaW_10等分(_nScr), _iwaW_5等分(_nScr), _iwaW_中間(_nScr)
    ReDim _iwaH_10等分(_nScr), _iwaH_5等分(_nScr), _iwaH_中間(_nScr)

    For i As Int32 = 0 To _nScr
      _wa(i) = Screen.AllScreens(i).WorkingArea
      _iwaL(i) = _wa(i).Left : _iwaT(i) = _wa(i).Top : _iwaW(i) = _wa(i).Width : _iwaH(i) = _wa(i).Height
      _iwaW_10等分(i) = CInt(_iwaW(i) * 0.075) : _iwaW_5等分(i) = _iwaW_10等分(i) * 2 : _iwaW_中間(i) = CInt(_iwaW(i) * 0.5)
      _iwaH_10等分(i) = CInt(_iwaH(i) * 0.075) : _iwaH_5等分(i) = _iwaH_10等分(i) * 2 : _iwaH_中間(i) = CInt(_iwaH(i) * 0.5)
    Next
  End Sub
  Public Sub tmr井字_Tick(sender As Object, e As EventArgs)
    If Control.MouseButtons <> Windows.Forms.MouseButtons.None Then Exit Sub ' 有按滑鼠時，就離開
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
    __tmr井字.Stop()
    Dim frm As Form = Nothing

    Select Case __hc_Selfrm
      Case __hc_f1Eng : frm = f1ENG
      Case Else : Exit Sub ' debug
    End Select
    m1.us_frm井字配置(frm)
  End Sub
  Public Sub us_frm井字配置(ByRef frm As Form)
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------
    ' mouse 在螢幕上的位置
    Dim cp = Windows.Forms.Cursor.Position, icpX = cp.X, icpY = cp.Y
    '-------------------------------------
    Dim iL, iT, iW, iH As Int32
    Dim iW_中間, iW_5等分, iW_10等分 As Int32
    Dim iH_中間, iH_5等分, iH_10等分 As Int32

    ' 偵測 mouse 在那1個 LCD 中
    For i As Int32 = 0 To _nScr
      If icpX >= _iwaL(i) AndAlso icpX <= _iwaW(i) + _iwaL(i) Then ' LCD 1
        iL = _iwaL(i) : iT = _iwaT(i) : iW = _iwaW(i) : iH = _iwaH(i)
        iW_中間 = _iwaW_中間(i) : iW_5等分 = _iwaW_5等分(i) : iW_10等分 = _iwaW_10等分(i)
        iH_中間 = _iwaH_中間(i) : iH_5等分 = _iwaH_5等分(i) : iH_10等分 = _iwaH_10等分(i)
        icpX -= iL ' 扣除 該LCD 左邊的距離
        Exit For
      End If
    Next
    '-------------------------------------
    ' 井字-九宮格
    Dim pt = frm.Location, sz = frm.Size ' 不在上述區間內，就用原來的
    Dim iLn, iTn, iWn, iHn As Int32
    iLn = pt.X : iTn = pt.Y : iWn = sz.Width : iHn = sz.Height

    Select Case True
      Case icpX >= iW_中間 - iW_10等分 AndAlso icpX <= iW_中間 + iW_10等分   ' LCD 中間 (常用放前面)
        Select Case True
          Case icpY >= iT AndAlso icpY <= iH_5等分   ' 中上 1/2
            iLn = iL : iTn = iT : iWn = iW : iHn = iH_中間
          Case icpY >= iH - iH_5等分 AndAlso icpY <= iH   ' 中下 1/2
            iLn = iL : iTn = iT + iH_中間 : iWn = iW : iHn = iH_中間
        End Select

      Case icpX >= 0 AndAlso icpX <= iW_5等分   ' LCD 最左邊
        Select Case True
          Case icpY >= iT AndAlso icpY <= iH_5等分   ' 左上 1/4
            iLn = iL : iTn = iT : iWn = iW_中間 : iHn = iH_中間
          Case icpY >= iH - iH_5等分 AndAlso icpY <= iH   ' 左下 1/4
            iLn = iL : iTn = iH_中間 : iWn = iW_中間 : iHn = iH_中間
          Case icpY >= iH_中間 - iH_10等分 AndAlso icpY <= iH_中間 + iH_10等分   ' 左中 1/2
            iLn = iL : iTn = iT : iWn = iW_中間 : iHn = iH
        End Select

      Case icpX >= iW - iW_5等分 AndAlso icpX <= iW   ' LCD 最右邊
        Select Case True
          Case icpY >= iT AndAlso icpY <= iH_5等分   ' 右上 1/4
            iLn = iW_中間 + iL : iTn = iT : iWn = iW_中間 : iHn = iH_中間
          Case icpY >= iH - iH_5等分 AndAlso icpY <= iH   ' 右下 1/4
            iLn = iW_中間 + iL : iTn = iH_中間 : iWn = iW_中間 : iHn = iH_中間
          Case icpY >= iH_中間 - iH_10等分 AndAlso icpY <= iH_中間 + iH_10等分   ' 右中 1/2
            iLn = iW_中間 + iL : iTn = iT : iWn = iW_中間 : iHn = iH
        End Select
    End Select
    '-------------------------------------
    frm.Location = New Point(iLn, iTn) ' 再度引發 Me.LocationChanged(e)
    frm.Size = New Size(iWn, iHn)
  End Sub
  Public Function uf_SQL更新(ByRef dt As DataTable, ByRef oda As OleDbDataAdapter) As Int32
    If dt.Rows.Count = 0 Then Return 0 ' dt 內若沒有資料列，就離開
    Dim N As Int32 = 0

    Try
      For Each dr As DataRow In dt.Rows
        dr.EndEdit() ' 將 current row 關閉 edit 狀態
      Next

      N = oda.Update(dt)
      dt.AcceptChanges()
      If N > 0 Then MsgBox(Now.ToString("HH:mm") + "：更新：" + dt.TableName + " = " + N.ToString + " 筆")
      Return N
    Catch ex As Exception
      Dim sf0 As New StackFrame(0, True) : m1.us_ErrMsg(sf0, Err.Description + vbNL + dt.TableName + " 資料表") : Return -1 : Exit Function
    End Try

    Return N ' 沒有更新
  End Function
  '===================================================================
#Region "擴充功能_Obj"
  <Extension()>
  Public Function Obj_2_iYMD(obj As Object) As String
    Dim DD = CDate(obj)
    Dim sMM = DD.Month.ToString : If sMM.Length = 1 Then sMM = "0" + sMM
    Dim sDD = DD.Day.ToString : If sDD.Length = 1 Then sDD = "0" + sDD
    Return DD.Year.ToString + sMM + sDD
    'Return CInt(DD.Year.ToString + sMM + sDD)
  End Function
#End Region
#Region "擴充功能_Date"
  <Extension()>
  Public Function TosY_M_D(DD As DateTime) As String
    Dim sMM = DD.Month.ToString : If sMM.Length = 1 Then sMM = "-0" + sMM Else sMM = "-" + sMM
    Dim sDD = DD.Day.ToString : If sDD.Length = 1 Then sDD = "-0" + sDD Else sDD = "-" + sDD
    Return DD.Year.ToString + sMM + sDD

    If False Then ' test
      'Dim i As Int32 = 0
      'Dim oo As Object = #2/3/2008#

      'Dim i2 = oo.Obj_2_iYMD ' NG，obj 本身不能，奇怪？
      'sMM = i.Obj_2_iYMD ' OK
      sMM = sMM.Obj_2_iYMD() ' OK
      sMM = DD.Obj_2_iYMD ' OK
    End If
  End Function
  <Extension()>
  Public Function TosYMD(DD As DateTime) As String
    Dim sMM = DD.Month.ToString : If sMM.Length = 1 Then sMM = "0" + sMM
    Dim sDD = DD.Day.ToString : If sDD.Length = 1 Then sDD = "0" + sDD
    Return DD.Year.ToString + sMM + sDD
  End Function
  <Extension()>
  Public Function ToiYMD(DD As DateTime) As Int32
    Dim sMM = DD.Month.ToString : If sMM.Length = 1 Then sMM = "0" + sMM
    Dim sDD = DD.Day.ToString : If sDD.Length = 1 Then sDD = "0" + sDD
    Return CInt(DD.Year.ToString + sMM + sDD)
  End Function

  <Extension()>
  Public Function 月底sYMD(DD As DateTime) As String
    Dim sYM = DD.ToString("yyyyMM"), s月底 = ""
    Select Case DD.Month
      Case 2
        If DD.Year Mod 4 = 0 Then s月底 = sYM + "29" Else s月底 = sYM + "28"
      Case 1, 3, 5, 7, 8, 10, 12
        s月底 = sYM + "31"
      Case Else
        s月底 = sYM + "30"
    End Select

    Return s月底
  End Function
  <Extension()>
  Public Function 月底iYMD(DD As DateTime) As Int32
    Dim sYM = DD.ToString("yyyyMM"), s月底 = ""
    Select Case DD.Month
      Case 2
        If DD.Year Mod 4 = 0 Then s月底 = sYM + "29" Else s月底 = sYM + "28"
      Case 1, 3, 5, 7, 8, 10, 12
        s月底 = sYM + "31"
      Case Else
        s月底 = sYM + "30"
    End Select

    Return CInt(s月底)
  End Function
  <Extension()>
  Public Function 月底Date(DD As DateTime) As DateTime
    Dim sYM = DD.ToString("yyyy-MM"), s月底 = ""
    Select Case DD.Month
      Case 2
        If DD.Year Mod 4 = 0 Then s月底 = sYM + "-29" Else s月底 = sYM + "-28"
      Case 1, 3, 5, 7, 8, 10, 12
        s月底 = sYM + "-31"
      Case Else
        s月底 = sYM + "-30"
    End Select

    Return CDate(s月底)
  End Function
#End Region
#Region "擴充功能_Int"
  <Extension()>
  Public Function ToDate(iYMD As Int32) As DateTime
    ' 20080818 =>2008/8/18
    Dim s = iYMD.ToString
    Return CDate(Mid(s, 1, 4) + "/" + Mid(s, 5, 2) + "/" + Mid(s, 7, 2))
  End Function
#End Region
#Region "擴充功能_Str"
  <Extension()>
  Public Function ToDate(s As String) As DateTime
    ' 20080818 =>2008/8/18
    Return CDate(Mid(s, 1, 4) + "/" + Mid(s, 5, 2) + "/" + Mid(s, 7, 2))
  End Function

  <Extension()>
  Public Function 月底sY_M_D(sYM As String) As String ' 2015/01/31
    Dim s月底 = "", sY = Mid(sYM, 1, 4), sM = Mid(sYM, 5, 2)
    Select Case CInt(sM)
      Case 2
        If CInt(sY) Mod 4 = 0 Then
          s月底 = sY + "/" + sM + "/29"
        Else
          s月底 = sY + "/" + sM + "/28"
        End If
      Case 1, 3, 5, 7, 8, 10, 12
        s月底 = sY + "/" + sM + "/31"
      Case Else
        s月底 = sY + "/" + sM + "/30"
    End Select

    Return s月底
  End Function

  <Extension()>
  Public Function 月底sYMD(sYM As String) As String ' 20160131
    Dim s月底 = "", sY = Mid(sYM, 1, 4), sM = Mid(sYM, 5, 2)
    Select Case CInt(sM)
      Case 2
        If CInt(sY) Mod 4 = 0 Then
          s月底 = sY + sM + "29"
        Else
          s月底 = sY + sM + "28"
        End If
      Case 1, 3, 5, 7, 8, 10, 12
        s月底 = sY + sM + "31"
      Case Else
        s月底 = sY + sM + "30"
    End Select

    Return s月底
  End Function

  <Extension()>
  Public Function 月初sYtw_M_D(sYM As String) As String ' 105/01/01
    Dim sY = Mid(sYM, 1, 4), sM = Mid(sYM, 5, 2), sYtw = (CInt(sY) - 1911).ToString
    Return sYtw + "/" + sM + "/01" ' 105/02/01
  End Function
  <Extension()>
  Public Function 月底sYtw_M_D(sYM As String) As String ' 105/01/31
    Dim s月底 = "", sY = Mid(sYM, 1, 4), sM = Mid(sYM, 5, 2), sYtw = (CInt(sY) - 1911).ToString
    Select Case CInt(sM)
      Case 2
        If CInt(sY) Mod 4 = 0 Then
          s月底 = sYtw + "/" + sM + "/29" ' 105/02/29
        Else
          s月底 = sYtw + "/" + sM + "/28" ' 104/02/29
        End If
      Case 1, 3, 5, 7, 8, 10, 12
        s月底 = sYtw + "/" + sM + "/31"
      Case Else
        s月底 = sYtw + "/" + sM + "/30"
    End Select

    Return s月底
  End Function
#End Region
End Module
