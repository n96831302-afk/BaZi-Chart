Module mdgv
  Public Const iW日期2 = 67, iW日期4 = 85, iW日期i = 75, iW日時 = 120
  Public Const iW數值 = 60, iW數值p2 = 70, iW數值p3 = 77, iW數值p4 = 85, iW數值大 = 85
  Public Const iWiYM = 60, iWiYQ = 60, iWiY = 50
  Public Const __iRowTH = 30 ' dgv.RowTemplate.Height, 21太小

  ' GetType()
  Public Tstr As Type = GetType(String), TDat As Type = GetType(DateTime)
  Public Ti16 As Type = GetType(Int16), Ti32 As Type = GetType(Int32), Ti64 As Type = GetType(Int64)
  Public Tsng As Type = GetType(Single), Tdbl As Type = GetType(Double)
  Public Tbln As Type = GetType(Boolean), Tobj As Type = GetType(Object)

  Public clr輸入欄 As Color = Color.DarkViolet, clr臨時欄 As Color = Color.DimGray
  Public clr標題底 As Color = Color.FromArgb(32, 32, 32)
  '-----------------------------------------------------------------
  Public Const iA00 = 0, iB01 = 1, iC02 = 2, iD03 = 3, iE04 = 4, iF05 = 5, iG06 = 6, iH07 = 7, iI08 = 8, iJ09 = 9, iK10 = 10, iL11 = 11, iM12 = 12, iN13 = 13
  Public Const iO14 = 14, iP15 = 15, iQ16 = 16, iR17 = 17, iS18 = 18, iT19 = 19, iU20 = 20, iV21 = 21, iW22 = 22, iX23 = 23, iY24 = 24, iZ25 = 25
  Public Const iAA26 = 26, iAB27 = 27, iAC28 = 28, iAD29 = 29, iAE30 = 30, iAF31 = 31, iAG32 = 32, iAH33 = 33, iAI34 = 34, iAJ35 = 35, iAK36 = 36, iAL37 = 37, iAM38 = 38
  Public Const iAN39 = 39, iAO40 = 40, iAP41 = 41, iAQ42 = 42, iAR43 = 43, iAS44 = 44, iAT45 = 45, iAU46 = 46, iAV47 = 47, iAW48 = 48, iAX49 = 49, iAY50 = 50, iAZ51 = 51
  Public Const iBA52 = 52, iBB53 = 53, iBC54 = 54, iBD55 = 55, iBE56 = 56, iBF57 = 57, iBG58 = 58, iBH59 = 59, iBI60 = 60, iBJ61 = 61, iBK62 = 62, iBL63 = 63, iBM64 = 64
  '-----------------------------------------------------------------
  ' DataRow
  Public Const drs_M As DataRowState = DataRowState.Modified
  Public Const drs_D As DataRowState = DataRowState.Deleted
  Public Const drs_A As DataRowState = DataRowState.Added
  Public Const dvrs_CR As DataViewRowState = DataViewRowState.CurrentRows
  Public Const __sfnt_DGV = "微軟正黑體" '細明體, 微軟正黑體, Cascadia Code, MS PGothic

  Public Sub 建_dgv框架_不寫_不刪(ByRef dgv As DataGridView, nColHgt As Int32)
    With dgv
      .ReadOnly = True
      .ShowCellToolTips = True '不顯示提示文字，太花時間
      .SelectionMode = DataGridViewSelectionMode.CellSelect
      .Font = New Font(__sfnt_DGV, 11.0F)
      .VirtualMode = True ' T=不可用要程式填值入cell，F=可用程式填入值
      .ScrollBars = ScrollBars.Both
      .BackgroundColor = Color.Black
      .GridColor = Color.DimGray
      '---------------------------------------------
      .DefaultCellStyle.ForeColor = Color.White
      .DefaultCellStyle.BackColor = Color.Black
      '.DefaultCellStyle.SelectionForeColor = Color.Yellow
      '.DefaultCellStyle.SelectionBackColor = Color.Transparent 會有殘影，不好看。若 = Color.Empty 則回到藍底。
      '---------------------------------------------
      .AutoGenerateColumns = False ' 自己建立欄位
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True

      .AllowUserToDeleteRows = False
      .AllowUserToAddRows = False
      .AllowUserToResizeRows = False
      '---------------------------------------------
      .ColumnHeadersVisible = True
      .ColumnHeadersDefaultCellStyle.BackColor = clr標題底
      .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
      .ColumnHeadersDefaultCellStyle.Font = New Font(__sfnt_DGV, 12, FontStyle.Bold)
      .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
      .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
      If nColHgt > 1 Then
        .ColumnHeadersHeight = __iRowTH * nColHgt
      Else
        .ColumnHeadersHeight = 26 ' 25太小，無法顯示出 排序3角圖形
      End If

      .RowHeadersVisible = False
      .RowHeadersWidth = 25
      .RowTemplate.Height = __iRowTH
      .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
      .RowHeadersDefaultCellStyle.SelectionBackColor = Color.Blue
    End With
  End Sub
  Public Sub 建_dgv框架_不寫_可刪(ByRef dgv As DataGridView, nColHgt As Int32)
    With dgv
      .ReadOnly = True
      .ShowCellToolTips = True '不顯示提示文字，太花時間
      .SelectionMode = DataGridViewSelectionMode.FullRowSelect
      '.SelectionMode = DataGridViewSelectionMode.CellSelect
      .Font = New Font(__sfnt_DGV, 11.0F)
      .VirtualMode = True ' T=不可用要程式填值入cell，F=可用程式填入值
      .ScrollBars = ScrollBars.Both
      .BackgroundColor = Color.Black
      .GridColor = Color.DimGray
      '---------------------------------------------
      .DefaultCellStyle.ForeColor = Color.White
      .DefaultCellStyle.BackColor = Color.Black
      '.DefaultCellStyle.SelectionForeColor = Color.Yellow
      '.DefaultCellStyle.SelectionBackColor = Color.Transparent 會有殘影，不好看。若 = Color.Empty 則回到藍底。
      '---------------------------------------------
      .AutoGenerateColumns = False
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True

      .AllowUserToDeleteRows = True ' *
      .AllowUserToAddRows = False
      .AllowUserToResizeRows = False
      '---------------------------------------------
      .ColumnHeadersVisible = True
      .ColumnHeadersDefaultCellStyle.BackColor = clr標題底
      .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
      .ColumnHeadersDefaultCellStyle.Font = New Font(__sfnt_DGV, 12, FontStyle.Bold)
      .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
      .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
      If nColHgt > 1 Then
        .ColumnHeadersHeight = __iRowTH * nColHgt
      Else
        .ColumnHeadersHeight = 26 ' 25太小，無法顯示出 排序3角圖形
      End If

      .RowHeadersVisible = True  ' *
      .RowHeadersWidth = 25
      .RowTemplate.Height = __iRowTH
      .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
      .RowHeadersDefaultCellStyle.SelectionBackColor = Color.Blue
    End With
  End Sub

  Public Sub 建_dgv框架_可寫_不刪(ByRef dgv As DataGridView, nColHgt As Int32)
    With dgv ' *號 => 不一樣的地方
      .ReadOnly = False
      .ShowCellToolTips = True
      .SelectionMode = DataGridViewSelectionMode.CellSelect
      .EditMode = DataGridViewEditMode.EditOnEnter ' *點下去馬上可編輯
      .VirtualMode = False  ' *要程式填值入cell
      .ScrollBars = ScrollBars.Both
      .Font = New Font(__sfnt_DGV, 11.0F)
      .BackgroundColor = Color.Black
      .GridColor = Color.DimGray
      '---------------------------------------------
      .DefaultCellStyle.ForeColor = Color.White
      .DefaultCellStyle.BackColor = Color.Black
      .DefaultCellStyle.SelectionForeColor = Color.Yellow
      '---------------------------------------------
      .AutoGenerateColumns = False
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True

      .AllowUserToAddRows = False
      .AllowUserToDeleteRows = False
      .AllowUserToResizeRows = False
      '---------------------------------------------
      .ColumnHeadersVisible = True
      .ColumnHeadersDefaultCellStyle.BackColor = clr標題底
      .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
      .ColumnHeadersDefaultCellStyle.Font = New Font(__sfnt_DGV, 12, FontStyle.Bold)
      .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
      .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
      If nColHgt > 1 Then
        .ColumnHeadersHeight = __iRowTH * nColHgt
      Else
        .ColumnHeadersHeight = 26 ' 25太小，無法顯示出 排序3角圖形
      End If

      .RowHeadersVisible = False
      .RowHeadersWidth = __iRowTH
      .RowTemplate.Height = __iRowTH
      .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
      .RowHeadersDefaultCellStyle.SelectionBackColor = Color.Blue
    End With
  End Sub
  Public Sub 建_dgv框架_可寫_可刪(ByRef dgv As DataGridView, nColHgt As Int32)
    With dgv ' *號 => 不一樣的地方
      .ReadOnly = False
      .ShowCellToolTips = True
      .SelectionMode = DataGridViewSelectionMode.CellSelect
      .EditMode = DataGridViewEditMode.EditOnEnter
      .VirtualMode = False  ' *要程式填值入cell
      .ScrollBars = ScrollBars.Both
      .Font = New Font(__sfnt_DGV, 11.0F)

      '.BackgroundColor = Color.Black
      '.ForeColor = Color.Yellow
      .GridColor = Color.DimGray
      '---------------------------------------------
      .DefaultCellStyle.ForeColor = Color.White
      .DefaultCellStyle.BackColor = Color.Black
      .DefaultCellStyle.SelectionForeColor = Color.Yellow
      '---------------------------------------------
      .AutoGenerateColumns = False
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True

      .AllowUserToAddRows = True  ' *
      .AllowUserToDeleteRows = True ' *
      .AllowUserToResizeRows = False
      '---------------------------------------------
      .ColumnHeadersVisible = True
      .ColumnHeadersDefaultCellStyle.BackColor = clr標題底
      .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
      .ColumnHeadersDefaultCellStyle.Font = New Font(__sfnt_DGV, 12, FontStyle.Bold)
      .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
      .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
      If nColHgt > 1 Then
        .ColumnHeadersHeight = __iRowTH * nColHgt
      Else
        .ColumnHeadersHeight = 26 ' 25太小，無法顯示出 排序3角圖形
      End If

      .RowHeadersVisible = True  ' *
      .RowHeadersWidth = 25
      .RowTemplate.Height = __iRowTH
      .RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
      ' [T]
      .RowHeadersDefaultCellStyle.SelectionBackColor = Color.Blue
      '.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Red
      '.RowHeadersDefaultCellStyle.SelectionForeColor = Color.Yellow
    End With
  End Sub
  '==================================================================
  Public Function uf欄_一般(iW As Int32, sColName As String,
    Optional sHdrTxt As String = "",
    Optional sFmt As String = "",
    Optional cHdrF As Color = Nothing,
    Optional cHdrB As Color = Nothing,
    Optional cDefF As Color = Nothing,
    Optional cDefB As Color = Nothing,
    Optional isRead As Boolean = False,
    Optional isVsb As Boolean = True,
    Optional isRsz As DataGridViewTriState = DataGridViewTriState.True,
    Optional isFrz As Boolean = False,
    Optional iDivW As Int32 = 1) As DataGridViewTextBoxColumn
    '---------------------------------------------
    ' 註：Optional cHdrF As Color = Color.White 錯，必須是運算式
    '---------------------------------------------
    Dim c As New DataGridViewTextBoxColumn
    c.Name = sColName

    If sHdrTxt.Length > 0 Then
      c.HeaderText = sHdrTxt
    Else
      c.HeaderText = sColName
    End If

    c.DataPropertyName = sColName
    c.Width = iW
    c.DefaultCellStyle.Format = sFmt
    c.HeaderCell.Style.ForeColor = cHdrF
    c.HeaderCell.Style.BackColor = cHdrB
    c.DefaultCellStyle.ForeColor = cDefF
    c.DefaultCellStyle.BackColor = cDefB

    c.ReadOnly = isRead
    c.Resizable = isRsz
    c.Frozen = isFrz
    If Not isVsb Then c.Visible = False
    If iDivW > 1 Then c.DividerWidth = iDivW
    Return c
  End Function
  Public Function uf欄_d日時4(sColName As String,
  Optional ByVal sHdrTxt As String = "",
    Optional cHdrF As Color = Nothing,
    Optional cHdrB As Color = Nothing,
    Optional cDefF As Color = Nothing,
    Optional cDefB As Color = Nothing,
    Optional isVsb As Boolean = True,
    Optional iDivW As Int32 = 1) As DataGridViewTextBoxColumn
    '---------------------------------------------
    Dim c As New DataGridViewTextBoxColumn
    c.Name = sColName : c.DataPropertyName = sColName : c.Width = 150 : c.DefaultCellStyle.Format = "yyyy.M.d HH:mm"
    If sHdrTxt.Length > 0 Then c.HeaderText = sHdrTxt Else c.HeaderText = sColName
    c.HeaderCell.Style.ForeColor = cHdrF
    c.HeaderCell.Style.BackColor = cHdrB
    c.DefaultCellStyle.ForeColor = cDefF
    c.DefaultCellStyle.BackColor = cDefB

    If Not isVsb Then c.Visible = False
    If iDivW > 1 Then c.DividerWidth = iDivW
    Return c
  End Function
  Public Function uf欄_chk(sColName As String, isReadOnly As Boolean,
    Optional iW As Int32 = 25,
    Optional sHdrTxt As String = "") As DataGridViewCheckBoxColumn
    Dim c As New DataGridViewCheckBoxColumn
    c.ThreeState = False : c.TrueValue = 1 : c.FalseValue = 0
    c.SortMode = DataGridViewColumnSortMode.Automatic
    c.Name = sColName : c.DataPropertyName = sColName
    If sHdrTxt.Length > 0 Then
      c.HeaderText = sHdrTxt
    Else
      c.HeaderText = sColName
    End If

    c.Width = iW : c.ReadOnly = isReadOnly
    c.Resizable = DataGridViewTriState.False
    'c.FlatStyle = FlatStyle.Flat
    'c.HeaderCell.Style.ForeColor = Color.Red
    Return c
  End Function
  '==================================================================
  Public Function uf_dgv寬度(ByRef dgv As DataGridView) As Int32
    Dim iHdrW, iW As Int32

    If dgv.RowHeadersVisible Then iHdrW = dgv.RowHeadersWidth

    iW = __iVsbW + iHdrW + 5
    For Each c As DataGridViewColumn In dgv.Columns
      If c.Visible Then iW += c.Width
    Next
    Return iW
  End Function
  Public Function uf_dgv寬度_沒捲軸(ByRef dgv As DataGridView) As Int32
    Dim iHdrW, iW As Int32

    If dgv.RowHeadersVisible Then iHdrW = dgv.RowHeadersWidth

    iW = iHdrW
    For Each c As DataGridViewColumn In dgv.Columns
      If c.Visible Then iW += c.Width
    Next
    Return iW
  End Function

  Public Sub us_選取列底線2(ByRef dgv As DataGridView)
    '    On Error GoTo Er : GoTo 0
    'Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
    '0:    '---------------------------------------------
    Dim ir1, ir0 As Int32
    ir1 = CInt(dgv.Tag)
    ir0 = dgv.CurrentCell.RowIndex

    If dgv.RowCount > ir1 Then dgv.Rows(ir1).DividerHeight = 0 ' 防 新dt 的row筆數比前1次少
    dgv.Rows(ir0).DividerHeight = 2
    dgv.Tag = ir0
  End Sub
  Public Sub us_選取列底線_顯示區(ByRef dgv As DataGridView)
    '    On Error GoTo Er : GoTo 0
    'Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
    '0:    '---------------------------------------------
    Dim i顯示區第1列 = dgv.FirstDisplayedScrollingRowIndex
    Dim i顯示區總列數 = dgv.DisplayedRowCount(False)

    For ir = i顯示區第1列 To i顯示區第1列 + i顯示區總列數 - 1
      dgv.Rows(ir).DividerHeight = 0
    Next
    dgv.CurrentRow.DividerHeight = 2
  End Sub

End Module
