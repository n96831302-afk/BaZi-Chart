Option Explicit On : Option Infer On : Option Strict On
Imports System.Data, System.Data.OleDb
Imports System.Collections.Generic
Imports System.Diagnostics

Module m易理
  ' 日干神煞 李鐵筆《四柱八字命運學》 P35
  Public __s日干_() As String = {甲, 乙, 丙, 丁, 戊, 己, 庚, 辛, 壬, 癸}
  Public __s文昌_() As String = {巳, 午, 申, 酉, 申, 酉, 亥, 子, 寅, 卯}
  Public __s金車_() As String = {辰, 巳, 未, 申, 未, 申, 戌, 亥, 丑, 寅}
  Public __s羊刃_() As String = {卯, 辰, 午, 未, 午, 未, 酉, 戌, 子, 丑}
  Public __s禄神_() As String = {寅, 卯, 巳, 午, 巳, 午, 申, 酉, 亥, 子} '=12長生的 臨官(建祿)
  Public __s紅艷_() As String = {申午, 申午, 寅, 未, 辰, 辰, 戌, 酉, 子, 申}
  Public __s天乙_() As String = {丑未, 子申, 亥酉, 戌酉, 丑未, 子申, 丑未, 寅午, 巳卯, 巳卯}

  ' 李鵬《快速六爻斷法》新增
  Public __s国印_() As String = {戊, 亥, 丑, 寅, 丑, 寅, 辰, 巳, 未, 申} '李鵬
  Public __s词馆_() As String = {庚寅, 辛卯, 乙巳, 戊午, 丁巳, 庚午, 壬申, 癸酉, 癸亥, 壬戌} '李鵬
  Public __s太极_() As String = {子午, 子午, 卯酉, 卯酉, 辰戌丑未, 辰戌丑未, 寅亥, 寅亥, 巳申, 巳申} '李鵬
  Public __s福神_() As String = {子寅, 丑卯, 子寅, 亥, 申, 未, 午, 巳, 辰, 丑卯} '李鵬

  ' 年支神煞，李鐵筆《四柱八字命運學》 P36
  Public __s年支_() As String = {子, 丑, 寅, 卯, 辰, 巳, 午, 未, 申, 酉, 戌, 亥}
  Public __s龍德_() As String = {未, 申, 酉, 戌, 亥, 子, 丑, 寅, 卯, 辰, 巳, 午} '1
  Public __s金匱_() As String = {子, 酉, 午, 卯, 子, 酉, 午, 卯, 子, 酉, 午, 卯} '2 四正支
  Public __s紅鶯_() As String = {卯, 寅, 丑, 子, 亥, 戌, 酉, 申, 未, 午, 巳, 辰}
  Public __s天狗_() As String = {戌, 亥, 子, 丑, 寅, 卯, 辰, 巳, 午, 未, 申, 酉}
  Public __s勾紋_() As String = {卯, 辰, 巳, 午, 未, 申, 酉, 戌, 亥, 子, 丑, 寅}
  Public __s岁破_() As String = {午, 未, 申, 酉, 戌, 亥, 子, 丑, 寅, 卯, 辰, 巳} ' =相沖
  Public __s破碎_() As String = {午, 未, 申, 酉, 戌, 亥, 子, 丑, 寅, 卯, 辰, 巳} ' =相沖
  Public __s大耗_() As String = {午, 未, 申, 酉, 戌, 亥, 子, 丑, 寅, 卯, 辰, 巳} ' =相沖10
  Public __s五鬼_() As String = {辰, 巳, 午, 未, 申, 酉, 戌, 亥, 子, 丑, 寅, 卯} '11
  Public __s年桃_() As String = {酉, 午, 卯, 子, 酉, 午, 卯, 子, 酉, 午, 卯, 子} '  四正支=日桃
  Public __s年刃_() As String = {戌, 酉, 申, 未, 午, 巳, 辰, 卯, 寅, 丑, 子, 亥} '5      =日刃
  Public __s年马_() As String = {寅, 亥, 申, 巳, 寅, 亥, 申, 巳, 寅, 亥, 申, 巳} '2      =日马(SKY自加)

  ' 日支神煞，李鐵筆《四柱八字命運學》 P37
  Public __s日支_() As String = {"子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"}
  Public __s将星_() As String = {"子", "酉", "午", "卯", "子", "酉", "午", "卯", "子", "酉", "午", "卯"} '1
  Public __s日马_() As String = {"寅", "亥", "申", "巳", "寅", "亥", "申", "巳", "寅", "亥", "申", "巳"} '2
  Public __s日桃_() As String = {"酉", "午", "卯", "子", "酉", "午", "卯", "子", "酉", "午", "卯", "子"}
  Public __s华盖_() As String = {"辰", "丑", "戌", "未", "辰", "丑", "戌", "未", "辰", "丑", "戌", "未"}
  Public __s劫煞_() As String = {"巳", "寅", "亥", "申", "巳", "寅", "亥", "申", "巳", "寅", "亥", "申"}
  Public __s亡神_() As String = {"亥", "申", "巳", "寅", "亥", "申", "巳", "寅", "亥", "申", "巳", "寅"}
  Public __s寡宿_() As String = {"戌", "戌", "丑", "丑", "丑", "辰", "辰", "辰", "未", "未", "未", "戌"}
  Public __s孤辰_() As String = {"寅", "寅", "巳", "巳", "巳", "申", "申", "申", "亥", "亥", "亥", "寅"}
  Public __s隔角_() As String = {"寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥", "子", "丑"} '10
  Public __s喪門_() As String = {"寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥", "子", "丑"} '11
  Public __s日刃_() As String = {"戌", "酉", "申", "未", "午", "巳", "辰", "卯", "寅", "丑", "子", "亥"} '5

  ' 月支神煞，李鐵筆《四柱八字命運學》 P37
  Public __s月支_() As String = {"子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"}
  Public __s月刃_() As String = {"午", "子", "丑", "未", "寅", "申", "卯", "酉", "辰", "戌", "巳", "亥"} '5 同日刃
  Public __s月桃_() As String = {"酉", "午", "卯", "子", "酉", "午", "卯", "子", "酉", "午", "卯", "子"} '  同日桃(SKY自加)
  Public __s月马_() As String = {"寅", "亥", "申", "巳", "寅", "亥", "申", "巳", "寅", "亥", "申", "巳"} '2 同日马(SKY自加)

  ' 以下為李鵬《快速六爻斷法》多的
  Public __s灾星_() As String = {"午", "卯", "子", "酉", "午", "卯", "子", "酉", "午", "卯", "子", "酉"} '李鵬
  Public __s罗网_() As String = {"　", "　", "　", "　", "巳", "辰", "　", "　", "　", "　", "亥", "戌"} '李鵬

  ' 12長生vs爻干                 {"甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"}
  Public __s长干_() As String = {"亥", "午", "寅", "酉", "寅", "酉", "巳", "子", "申", "卯"}
  Public __s沐干_() As String = {"子", "巳", "卯", "申", "卯", "申", "午", "亥", "酉", "寅"}
  Public __s冠干_() As String = {"丑", "辰", "辰", "未", "辰", "未", "未", "戌", "戌", "丑"}
  Public __s臨干_() As String = {"寅", "卯", "巳", "午", "巳", "午", "申", "酉", "亥", "子"}
  Public __s帝干_() As String = {"卯", "寅", "午", "巳", "午", "巳", "酉", "申", "子", "亥"}
  Public __s衰干_() As String = {"辰", "丑", "未", "辰", "未", "辰", "戌", "未", "丑", "戌"}
  Public __s病干_() As String = {"巳", "子", "申", "卯", "申", "卯", "亥", "午", "寅", "酉"}
  Public __s死干_() As String = {"午", "亥", "酉", "寅", "酉", "寅", "子", "巳", "卯", "申"}
  Public __s墓干_() As String = {"未", "戌", "戌", "丑", "戌", "丑", "丑", "辰", "辰", "未"}
  Public __s绝干_() As String = {"申", "酉", "亥", "子", "亥", "子", "寅", "卯", "巳", "午"}
  Public __s胎干_() As String = {"酉", "申", "子", "亥", "子", "亥", "卯", "寅", "午", "巳"}
  Public __s养干_() As String = {"戌", "未", "丑", "戌", "丑", "戌", "辰", "丑", "未", "辰"}
  ' 12長生vs爻支 Ref：李鵬快速六爻斷法(不用)
  Public __s爻支_() As String = {"子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"}
  Public __s长支_() As String = {"申", "申", "亥", "亥", "申", "寅", "寅", "申", "巳", "巳", "申", "申"} ' 長生
  Public __s沐支_() As String = {"酉", "酉", "子", "子", "酉", "卯", "卯", "酉", "午", "午", "酉", "酉"} ' 沐浴
  Public __s冠支_() As String = {"戌", "戌", "丑", "丑", "戌", "辰", "辰", "戌", "未", "未", "戌", "戌"} ' 冠帶
  Public __s臨支_() As String = {"亥", "亥", "寅", "寅", "亥", "巳", "巳", "亥", "申", "申", "亥", "亥"}
  Public __s帝支_() As String = {"子", "子", "卯", "卯", "子", "午", "午", "子", "酉", "酉", "子", "子"} ' 帝旺
  Public __s衰支_() As String = {"丑", "丑", "辰", "辰", "丑", "未", "未", "丑", "戌", "戌", "丑", "丑"}
  Public __s病支_() As String = {"寅", "寅", "巳", "巳", "寅", "申", "申", "寅", "亥", "亥", "寅", "寅"}
  Public __s死支_() As String = {"卯", "卯", "午", "午", "卯", "酉", "酉", "卯", "子", "子", "卯", "卯"}

  Public __s墓支_() As String = {"辰", "　", "未", "未", "　", "戌", "戌", "　", "丑", "丑", "　", "辰"}
  Public __s绝支_() As String = {"巳", "巳", "申", "申", "巳", "亥", "亥", "巳", "寅", "寅", "巳", "巳"}
  Public __s胎支_() As String = {"午", "午", "酉", "酉", "午", "子", "子", "午", "卯", "卯", "午", "午"}
  Public __s养支_() As String = {"未", "未", "戌", "戌", "未", "丑", "丑", "未", "辰", "辰", "未", "未"}

  '基本資料
  Public __sYgz, __sYg, __sYz, __sY空亡, __sY出空 As String
  Public __sMgz, __sMg, __sMz, __sM空亡, __sM出空 As String
  Public __sDgz, __sDg, __sDz, __sD空亡, __sD出空, __sDgz_新法, __s閏月 As String
  Public __sHgz, __sHg, __sHz, __sH空亡, __sH出空 As String
  Public __s月將, __s节气 As String
  Public __s六壬課式號碼_(12, 12), __s12天將_(12) As String
  Public __s60甲子_(5, 9), __s60甲子Idx_(60), __s12岁君_(12, 12), __s144zz_(11, 11) As String
  Public __s五行_(5), __sA8全_(8) As String
  Public __cnt农历 As OleDbConnection
  Public __dt农历, __dt节气 As New DataTable, __D农历 As DateTime
  Public __dic60納音, __dic支藏干_古法, __dic支藏干_六行_寅_甲丙己, __dic支藏干_六行_寅_甲丙戊, __dic支藏干_六行_寅_甲戊 As New Dictionary(Of String, String)
  Public __dic干藏支_甲_寅亥, __dic干藏支_甲_寅卯, __dic干藏支_甲_寅卯_乙_卯未 As New Dictionary(Of String, String)
  Public __dic西元年_60甲子 As New Dictionary(Of Int32, String)
  Public __dic干, __dic支 As New Dictionary(Of String, Int32)
  Public Const vbNL = vbCrLf, vbNL2 = vbCrLf + vbCrLf
  '================================================================
  Public Sub us_基本資料_干支()
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------------
    If __dic干.Count > 0 Then Exit Sub ' 在別的form 有call過了
    With __dic干
      .Add("甲", 1) : .Add("乙", 2) : .Add("丙", 3) : .Add("丁", 4)
      .Add("戊", 5) : .Add("己", 6)
      .Add("庚", 7) : .Add("辛", 8) : .Add("壬", 9) : .Add("癸", 10)
    End With
    With __dic支
      .Add("子", 1) : .Add("丑", 2) : .Add("寅", 3) : .Add("卯", 4)
      .Add("辰", 5) : .Add("巳", 6) : .Add("午", 7) : .Add("未", 8)
      .Add("申", 9) : .Add("酉", 10) : .Add("戌", 11) : .Add("亥", 12)
    End With

    Dim i, j, i1 As Int32
    __s五行_(1) = "金" : __s五行_(2) = "水" : __s五行_(3) = "木" : __s五行_(4) = "火" : __s五行_(5) = "土"

    Dim g60甲子() As String = {"甲子", "甲戌", "甲申", "甲午", "甲辰", "甲寅",
                           "乙丑", "乙亥", "乙酉", "乙未", "乙巳", "乙卯",
                           "丙寅", "丙子", "丙戌", "丙申", "丙午", "丙辰",
                           "丁卯", "丁丑", "丁亥", "丁酉", "丁未", "丁巳",
                           "戊辰", "戊寅", "戊子", "戊戌", "戊申", "戊午",
                           "己巳", "己卯", "己丑", "己亥", "己酉", "己未",
                           "庚午", "庚辰", "庚寅", "庚子", "庚戌", "庚申",
                           "辛未", "辛巳", "辛卯", "辛丑", "辛亥", "辛酉",
                           "壬申", "壬午", "壬辰", "壬寅", "壬子", "壬戌",
                           "癸酉", "癸未", "癸巳", "癸卯", "癸丑", "癸亥"}
    Dim s1_(60), s2_(5, 9) As String
    s1_(1) = "甲子" : s1_(11) = "甲戌" : s1_(21) = "甲申" : s1_(31) = "甲午" : s1_(41) = "甲辰" : s1_(51) = "甲寅"
    s1_(2) = "乙丑" : s1_(12) = "乙亥" : s1_(22) = "乙酉" : s1_(32) = "乙未" : s1_(42) = "乙巳" : s1_(52) = "乙卯"
    s1_(3) = "丙寅" : s1_(13) = "丙子" : s1_(23) = "丙戌" : s1_(33) = "丙申" : s1_(43) = "丙午" : s1_(53) = "丙辰"
    s1_(4) = "丁卯" : s1_(14) = "丁丑" : s1_(24) = "丁亥" : s1_(34) = "丁酉" : s1_(44) = "丁未" : s1_(54) = "丁巳"
    s1_(5) = "戊辰" : s1_(15) = "戊寅" : s1_(25) = "戊子" : s1_(35) = "戊戌" : s1_(45) = "戊申" : s1_(55) = "戊午"

    s1_(6) = "己巳" : s1_(16) = "己卯" : s1_(26) = "己丑" : s1_(36) = "己亥" : s1_(46) = "己酉" : s1_(56) = "己未"
    s1_(7) = "庚午" : s1_(17) = "庚辰" : s1_(27) = "庚寅" : s1_(37) = "庚子" : s1_(47) = "庚戌" : s1_(57) = "庚申"
    s1_(8) = "辛未" : s1_(18) = "辛巳" : s1_(28) = "辛卯" : s1_(38) = "辛丑" : s1_(48) = "辛亥" : s1_(58) = "辛酉"
    s1_(9) = "壬申" : s1_(19) = "壬午" : s1_(29) = "壬辰" : s1_(39) = "壬寅" : s1_(49) = "壬子" : s1_(59) = "壬戌"
    s1_(10) = "癸酉" : s1_(20) = "癸未" : s1_(30) = "癸巳" : s1_(40) = "癸卯" : s1_(50) = "癸丑" : s1_(60) = "癸亥"
    __s60甲子Idx_ = s1_

    s2_(0, 0) = "甲子" : s2_(1, 0) = "甲戌" : s2_(2, 0) = "甲申" : s2_(3, 0) = "甲午" : s2_(4, 0) = "甲辰" : s2_(5, 0) = "甲寅"
    s2_(0, 1) = "乙丑" : s2_(1, 1) = "乙亥" : s2_(2, 1) = "乙酉" : s2_(3, 1) = "乙未" : s2_(4, 1) = "乙巳" : s2_(5, 1) = "乙卯"
    s2_(0, 2) = "丙寅" : s2_(1, 2) = "丙子" : s2_(2, 2) = "丙戌" : s2_(3, 2) = "丙申" : s2_(4, 2) = "丙午" : s2_(5, 2) = "丙辰"
    s2_(0, 3) = "丁卯" : s2_(1, 3) = "丁丑" : s2_(2, 3) = "丁亥" : s2_(3, 3) = "丁酉" : s2_(4, 3) = "丁未" : s2_(5, 3) = "丁巳"
    s2_(0, 4) = "戊辰" : s2_(1, 4) = "戊寅" : s2_(2, 4) = "戊子" : s2_(3, 4) = "戊戌" : s2_(4, 4) = "戊申" : s2_(5, 4) = "戊午"

    s2_(0, 5) = "己巳" : s2_(1, 5) = "己卯" : s2_(2, 5) = "己丑" : s2_(3, 5) = "己亥" : s2_(4, 5) = "己酉" : s2_(5, 5) = "己未"
    s2_(0, 6) = "庚午" : s2_(1, 6) = "庚辰" : s2_(2, 6) = "庚寅" : s2_(3, 6) = "庚子" : s2_(4, 6) = "庚戌" : s2_(5, 6) = "庚申"
    s2_(0, 7) = "辛未" : s2_(1, 7) = "辛巳" : s2_(2, 7) = "辛卯" : s2_(3, 7) = "辛丑" : s2_(4, 7) = "辛亥" : s2_(5, 7) = "辛酉"
    s2_(0, 8) = "壬申" : s2_(1, 8) = "壬午" : s2_(2, 8) = "壬辰" : s2_(3, 8) = "壬寅" : s2_(4, 8) = "壬子" : s2_(5, 8) = "壬戌"
    s2_(0, 9) = "癸酉" : s2_(1, 9) = "癸未" : s2_(2, 9) = "癸巳" : s2_(3, 9) = "癸卯" : s2_(4, 9) = "癸丑" : s2_(5, 9) = "癸亥"
    __s60甲子_ = s2_

    With __dic支藏干_古法
      .Add("子", "癸") : .Add("丑", "己癸辛") : .Add("寅", "甲丙戊")
      .Add("卯", "乙") : .Add("辰", "戊乙癸") : .Add("巳", "丙庚戊")
      .Add("午", "丁己") : .Add("未", "己丁乙") : .Add("申", "庚壬戊")
      .Add("酉", "辛") : .Add("戌", "戊辛丁") : .Add("亥", "壬甲")
    End With

    If TT Then
      With __dic支藏干_六行_寅_甲丙戊
        .Add("子", "癸壬") : .Add("丑", "己癸辛") : .Add("寅", "甲丙戊") ' 寅变
        .Add("卯", "乙甲") : .Add("辰", "己乙癸") : .Add("巳", "丙戊庚")
        .Add("午", "丁丙") : .Add("未", "戊丁乙") : .Add("申", "庚壬己") ' 申变
        .Add("酉", "辛庚") : .Add("戌", "戊辛丁") : .Add("亥", "壬己甲")
      End With
    Else ' 調生的順序
      With __dic支藏干_六行_寅_甲丙戊
        .Add("子", "癸壬") : .Add("丑", "辛癸己") : .Add("寅", "甲丙戊")
        .Add("卯", "乙甲") : .Add("辰", "癸己乙") : .Add("巳", "丙戊庚")
        .Add("午", "丁丙") : .Add("未", "乙丁戊") : .Add("申", "庚壬己")
        .Add("酉", "辛庚") : .Add("戌", "丁戊辛") : .Add("亥", "壬己甲")
      End With
    End If

    With __dic支藏干_六行_寅_甲丙己
      .Add("子", "癸壬") : .Add("丑", "己癸辛") : .Add("寅", "甲丙己")
      .Add("卯", "乙甲") : .Add("辰", "己乙癸") : .Add("巳", "丙庚戊")
      .Add("午", "丁丙") : .Add("未", "戊丁乙") : .Add("申", "庚壬戊")
      .Add("酉", "辛庚") : .Add("戌", "戊辛丁") : .Add("亥", "壬甲己")
    End With
    With __dic支藏干_六行_寅_甲戊
      .Add("子", "癸壬") : .Add("丑", "己癸辛") : .Add("寅", "甲戊") ' 寅变
      .Add("卯", "乙甲") : .Add("辰", "己乙癸") : .Add("巳", "丙庚戊")
      .Add("午", "丁丙") : .Add("未", "戊丁乙") : .Add("申", "庚己") ' 申变
      .Add("酉", "辛庚") : .Add("戌", "戊辛丁") : .Add("亥", "壬甲己")
    End With

    With __dic干藏支_甲_寅亥
      .Add("甲", "寅亥") : .Add("丙", "巳寅") : .Add("庚", "申巳") : .Add("壬", "亥申")
      .Add("乙", "卯辰") : .Add("丁", "午未") : .Add("辛", "酉戌") : .Add("癸", "子丑")
      .Add("己", "丑辰") : .Add("戊", "未戌") '.Add("戊", "辰戌"): .Add("己", "丑未")
    End With
    With __dic干藏支_甲_寅卯
      .Add("甲", "寅卯") : .Add("丙", "巳午") : .Add("庚", "申酉") : .Add("壬", "亥子")
      .Add("乙", "卯辰") : .Add("丁", "午未") : .Add("辛", "酉戌") : .Add("癸", "子丑")
      .Add("己", "丑辰") : .Add("戊", "未戌")
    End With
    With __dic干藏支_甲_寅卯_乙_卯未
      .Add("甲", "寅卯") : .Add("丙", "巳午") : .Add("庚", "申酉") : .Add("壬", "亥子")
      .Add("乙", "卯未") : .Add("丁", "午戌") : .Add("辛", "酉丑") : .Add("癸", "子辰") ' 四正支变
      .Add("己", "丑辰") : .Add("戊", "未戌")
    End With

    ' 因為干支配在八字中有地支藏干，每本書寫的都不一樣，很奇怪，故SKY自創了支支配
    Dim s(11, 11) As String
    s(0, 0) = "子子" : s(0, 1) = "子丑" : s(0, 2) = "子寅" : s(0, 3) = "子卯" : s(0, 4) = "子辰" : s(0, 5) = "子巳" : s(0, 6) = "子午" : s(0, 7) = "子未" : s(0, 8) = "子申" : s(0, 9) = "子酉" : s(0, 10) = "子戌" : s(0, 11) = "子亥"
    s(1, 0) = "丑子" : s(1, 1) = "丑丑" : s(1, 2) = "丑寅" : s(1, 3) = "丑卯" : s(1, 4) = "丑辰" : s(1, 5) = "丑巳" : s(1, 6) = "丑午" : s(1, 7) = "丑未" : s(1, 8) = "丑申" : s(1, 9) = "丑酉" : s(1, 10) = "丑戌" : s(1, 11) = "丑亥"
    s(2, 0) = "寅子" : s(2, 1) = "寅丑" : s(2, 2) = "寅寅" : s(2, 3) = "寅卯" : s(2, 4) = "寅辰" : s(2, 5) = "寅巳" : s(2, 6) = "寅午" : s(2, 7) = "寅未" : s(2, 8) = "寅申" : s(2, 9) = "寅酉" : s(2, 10) = "寅戌" : s(2, 11) = "寅亥"
    s(3, 0) = "卯子" : s(3, 1) = "卯丑" : s(3, 2) = "卯寅" : s(3, 3) = "卯卯" : s(3, 4) = "卯辰" : s(3, 5) = "卯巳" : s(3, 6) = "卯午" : s(3, 7) = "卯未" : s(3, 8) = "卯申" : s(3, 9) = "卯酉" : s(3, 10) = "卯戌" : s(3, 11) = "卯亥"
    s(4, 0) = "辰子" : s(4, 1) = "辰丑" : s(4, 2) = "辰寅" : s(4, 3) = "辰卯" : s(4, 4) = "辰辰" : s(4, 5) = "辰巳" : s(4, 6) = "辰午" : s(4, 7) = "辰未" : s(4, 8) = "辰申" : s(4, 9) = "辰酉" : s(4, 10) = "辰戌" : s(4, 11) = "辰亥"
    s(5, 0) = "巳子" : s(5, 1) = "巳丑" : s(5, 2) = "巳寅" : s(5, 3) = "巳卯" : s(5, 4) = "巳辰" : s(5, 5) = "巳巳" : s(5, 6) = "巳午" : s(5, 7) = "巳未" : s(5, 8) = "巳申" : s(5, 9) = "巳酉" : s(5, 10) = "巳戌" : s(5, 11) = "巳亥"
    s(6, 0) = "午子" : s(6, 1) = "午丑" : s(6, 2) = "午寅" : s(6, 3) = "午卯" : s(6, 4) = "午辰" : s(6, 5) = "午巳" : s(6, 6) = "午午" : s(6, 7) = "午未" : s(6, 8) = "午申" : s(6, 9) = "午酉" : s(6, 10) = "午戌" : s(6, 11) = "午亥"
    s(7, 0) = "未子" : s(7, 1) = "未丑" : s(7, 2) = "未寅" : s(7, 3) = "未卯" : s(7, 4) = "未辰" : s(7, 5) = "未巳" : s(7, 6) = "未午" : s(7, 7) = "未未" : s(7, 8) = "未申" : s(7, 9) = "未酉" : s(7, 10) = "未戌" : s(7, 11) = "未亥"
    s(8, 0) = "申子" : s(8, 1) = "申丑" : s(8, 2) = "申寅" : s(8, 3) = "申卯" : s(8, 4) = "申辰" : s(8, 5) = "申巳" : s(8, 6) = "申午" : s(8, 7) = "申未" : s(8, 8) = "申申" : s(8, 9) = "申酉" : s(8, 10) = "申戌" : s(8, 11) = "申亥"
    s(9, 0) = "酉子" : s(9, 1) = "酉丑" : s(9, 2) = "酉寅" : s(9, 3) = "酉卯" : s(9, 4) = "酉辰" : s(9, 5) = "酉巳" : s(9, 6) = "酉午" : s(9, 7) = "酉未" : s(9, 8) = "酉申" : s(9, 9) = "酉酉" : s(9, 10) = "酉戌" : s(9, 11) = "酉亥"
    s(10, 0) = "戌子" : s(10, 1) = "戌丑" : s(10, 2) = "戌寅" : s(10, 3) = "戌卯" : s(10, 4) = "戌辰" : s(10, 5) = "戌巳" : s(10, 6) = "戌午" : s(10, 7) = "戌未" : s(10, 8) = "戌申" : s(10, 9) = "戌酉" : s(10, 10) = "戌戌" : s(10, 11) = "戌亥"
    s(11, 0) = "亥子" : s(11, 1) = "亥丑" : s(11, 2) = "亥寅" : s(11, 3) = "亥卯" : s(11, 4) = "亥辰" : s(11, 5) = "亥巳" : s(11, 6) = "亥午" : s(11, 7) = "亥未" : s(11, 8) = "亥申" : s(11, 9) = "亥酉" : s(11, 10) = "亥戌" : s(11, 11) = "亥亥"

    For i = 0 To 11 : For j = 0 To 11 : __s144zz_(i, j) = s(i, j) : Next : Next
    __s12天將_(1) = "貴" : __s12天將_(2) = "蛇" : __s12天將_(3) = "雀"
    __s12天將_(4) = "合" : __s12天將_(5) = "陈" : __s12天將_(6) = "龙"
    __s12天將_(7) = "空" : __s12天將_(8) = "虎" : __s12天將_(9) = "常"
    __s12天將_(10) = "武" : __s12天將_(11) = "阴" : __s12天將_(12) = "后"
    '-----------------------------------------------
    '四柱的 12岁君（資料忘了從那來的）--------------------------------
    'Dim s_() As String = {"太岁", "太阳", "喪门", "太阴", "五鬼", "死符", "岁破", "龙德", "白虎", "福德", "天狗", "病符"}
    Dim s_() As String = {"太岁", "病符", "天狗", "福德", "白虎", "龙德", "岁破", "死符", "五鬼", "太阴", "喪门", "太阳"}
    Dim i生, i流 As Int32
    Dim s1 As String

    '左上方
    For i生 = 1 To 12
      i1 = 0
      For i流 = i生 To 12
        __s12岁君_(i生, i流) = s_(i1)
        i1 = i1 + 1
      Next
      'i生=3, i流=1,2 => g12岁君(10), g12岁君(11)
      For i流 = 1 To i生 - 1
        __s12岁君_(i生, i流) = s_(i1)
        i1 = i1 + 1
      Next
    Next
    'check
    's1 = "子   丑   寅   卯   辰   巳   午   未   申   酉   戌   亥" + vbNL
    'For i = 1 To 12
    ' For j = 1 To 12
    ' s1 = s1 + __s12岁君_(i, j) + " "
    ' Next
    ' s1 = s1 + vbNL
    'Next
    'f六爻.txt範例.Text = s1

    '六壬課式號碼表-------------------------------------
    For j = 1 To 12
      i1 = 0
      For i = j To 12
        i1 = i1 + 1
        __s六壬課式號碼_(i, j) = Format(i1, "00")
      Next
    Next
    '14 -下半部的值 = 上半部的值
    For j = 2 To 12
      For i = 1 To j - 1
        i1 = 14 - CInt(__s六壬課式號碼_(j, i))
        __s六壬課式號碼_(i, j) = Format(i1, "00")
      Next
    Next
    'check
    's1 = "子 丑 寅 卯 辰 巳 午 未 申 酉 戌 亥" + vbNL
    'For i = 1 To 12
    'For j = 1 To 12
    '  s1 = s1 + __六壬課式號碼_(i, j) + " "
    'Next
    's1 = s1 + vbNL
    '  Next
    'tDes1.Text = s1

    With __dic60納音
      .Add("甲子", "海中金") : .Add("乙丑", "海中金") : .Add("丙寅", "爐中火") : .Add("丁卯", "爐中火") : .Add("戊辰", "大林木") : .Add("己巳", "大林木") : .Add("庚午", "路旁土") : .Add("辛未", "路旁土") : .Add("壬申", "劍峰金") : .Add("癸酉", "劍峰金")
      .Add("甲戌", "山頭火") : .Add("乙亥", "山頭火") : .Add("丙子", "澗下水") : .Add("丁丑", "澗下水") : .Add("戊寅", "城牆土") : .Add("己卯", "城牆土") : .Add("庚辰", "白蠟金") : .Add("辛巳", "白蠟金") : .Add("壬午", "楊柳木") : .Add("癸未", "楊柳木")
      .Add("甲申", "泉中水") : .Add("乙酉", "泉中水") : .Add("丙戌", "屋上土") : .Add("丁亥", "屋上土") : .Add("戊子", "霹靂火") : .Add("己丑", "霹靂火") : .Add("庚寅", "松柏木") : .Add("辛卯", "松柏木") : .Add("壬辰", "長流水") : .Add("癸巳", "長流水")
      .Add("甲午", "沙中金") : .Add("乙未", "沙中金") : .Add("丙申", "山下火") : .Add("丁酉", "山下火") : .Add("戊戌", "平地木") : .Add("己亥", "平地木") : .Add("庚子", "壁上土") : .Add("辛丑", "壁上土") : .Add("壬寅", "金箔金") : .Add("癸卯", "金箔金")
      .Add("甲辰", "佛燈火") : .Add("乙巳", "佛燈火") : .Add("丙午", "天河水") : .Add("丁未", "天河水") : .Add("戊申", "大驛土") : .Add("己酉", "大驛土") : .Add("庚戌", "釵釧金") : .Add("辛亥", "釵釧金") : .Add("壬子", "桑松木") : .Add("癸丑", "桑松木")
      .Add("甲寅", "大溪水") : .Add("乙卯", "大溪水") : .Add("丙辰", "沙中土") : .Add("丁巳", "沙中土") : .Add("戊午", "天上火") : .Add("己未", "天上火") : .Add("庚申", "石榴木") : .Add("辛酉", "石榴木") : .Add("壬戌", "大海水") : .Add("癸亥", "大海水")
    End With

    '---------------------------------------------
    ' 公式，易學VB公式.docx
    Dim i餘數, iG, iZ As Int32, sG = "", sZ = ""

    For iY = 0 To Today.Year + 100 ' 康熙皇帝 1654年
      i餘數 = (iY - 3) Mod 60
      iG = i餘數 Mod 10
      iZ = i餘數 Mod 12

      Select Case iG
        Case 1 : sG = "甲" : Case 2 : sG = "乙" : Case 3 : sG = "丙" : Case 4 : sG = "丁" : Case 5 : sG = "戊"
        Case 6 : sG = "己" : Case 7 : sG = "庚" : Case 8 : sG = "辛" : Case 9 : sG = "壬" : Case 0 : sG = "癸"
      End Select
      Select Case iZ
        Case 1 : sZ = "子" : Case 2 : sZ = "丑" : Case 3 : sZ = "寅" : Case 4 : sZ = "卯"
        Case 5 : sZ = "辰" : Case 6 : sZ = "巳" : Case 7 : sZ = "午" : Case 8 : sZ = "未"
        Case 9 : sZ = "申" : Case 10 : sZ = "酉" : Case 11 : sZ = "戌" : Case 0 : sZ = "亥"
      End Select

      __dic西元年_60甲子.Add(iY, sG + sZ) ': Console.WriteLine(iY.ToString + vbTab + sG + sZ) ' test
    Next
  End Sub

  Public __dic乾巽_姤 As New Dictionary(Of String, String)
  Public Sub us_基本資料_八卦()
    If __dic乾巽_姤.Count > 0 Then Exit Sub
    With __dic乾巽_姤
      .Add("乾乾", "乾") : .Add("乾巽", "姤") : .Add("乾艮", "遁") : .Add("乾坤", "否")
      .Add("巽坤", "观") : .Add("艮坤", "剥") : .Add("离坤", "晋") : .Add("离乾", "有") ' 大有

      .Add("兑兑", "兑") : .Add("兑坎", "困") : .Add("兑坤", "萃") : .Add("兑艮", "咸")
      .Add("坎艮", "蹇") : .Add("坤艮", "谦") : .Add("震艮", "小") : .Add("震兑", "妹") ' 小过，归妹

      .Add("离离", "离") : .Add("离艮", "旅") : .Add("离巽", "鼎") : .Add("离坎", "未") ' 未济
      .Add("艮坎", "蒙") : .Add("巽坎", "涣") : .Add("乾坎", "讼") : .Add("乾离", "同") ' 同人

      .Add("震震", "震") : .Add("震坤", "豫") : .Add("震坎", "解") : .Add("震巽", "恒")
      .Add("坤巽", "升") : .Add("坎巽", "井") : .Add("兑巽", "过") : .Add("兑震", "随") ' 大过

      .Add("巽巽", "巽") : .Add("巽乾", "竺") : .Add("巽离", "家") : .Add("巽震", "益") ' 小畜=竺，家人
      .Add("乾震", "无") : .Add("离震", "噬") : .Add("艮震", "颐") : .Add("艮巽", "蛊") ' 无妄，噬嗑

      .Add("坎坎", "坎") : .Add("坎兑", "节") : .Add("坎震", "屯") : .Add("坎离", "既") ' 既济
      .Add("兑离", "革") : .Add("震离", "丰") : .Add("坤离", "明") : .Add("坤坎", "师") ' 明夷

      .Add("艮艮", "艮") : .Add("艮离", "贲") : .Add("艮乾", "畜") : .Add("艮兑", "损") ' 大畜
      .Add("离兑", "睽") : .Add("乾兑", "履") : .Add("巽兑", "孚") : .Add("巽艮", "渐") ' 中孚

      .Add("坤坤", "坤") : .Add("坤震", "复") : .Add("坤兑", "临") : .Add("坤乾", "泰")
      .Add("震乾", "壮") : .Add("兑乾", "夬") : .Add("坎乾", "需") : .Add("坎坤", "比") ' 大壮
    End With
  End Sub
  Public Function uf八卦_古法(sGZ As String) As String
    Dim sG = Mid(sGZ, 1, 1), sZ = Mid(sGZ, 2, 1)
    Dim sT = "", sB = ""

    Select Case sG
      Case "甲" : sT = "震" : Case "乙" : sT = "巽"
      Case "庚" : sT = "乾" : Case "辛" : sT = "兑"
      Case "丙" : sT = "离" : Case "丁" : sT = "离"
      Case "壬" : sT = "坎" : Case "癸" : sT = "坎"
      Case "戊" : sT = "艮"
      Case "己" : sT = "坤"
    End Select

    Select Case sZ
      Case "寅" : sB = "震" : Case "卯" : sB = "巽"
      Case "申" : sB = "乾" : Case "酉" : sB = "兑"
      Case "巳" : sB = "离" : Case "午" : sB = "离"
      Case "亥" : sB = "坎" : Case "子" : sB = "坎"
      Case "辰", "戌" : sB = "艮"
      Case "丑", "未" : sB = "坤"
    End Select

    Return __dic乾巽_姤(sT + sB) ' 乾+巽=姤
  End Function

  Public Function uf流年多出的月干_五行(sDg As String, sYg As String) As String
    Dim s1 = "", s2 = "", s = ""
    Select Case sYg
      Case "甲", "己" : s1 = "丙" : s2 = "丁" : s = s1 + uf_印綬_干vs干_五行(sDg, s1) + "+" + s2 + uf_印綬_干vs干_五行(sDg, s2)
      Case "乙", "庚" : s1 = "戊" : s2 = "己" : s = s1 + uf_印綬_干vs干_五行(sDg, s1) + "+" + s2 + uf_印綬_干vs干_五行(sDg, s2)
      Case "丙", "辛" : s1 = "庚" : s2 = "辛" : s = s1 + uf_印綬_干vs干_五行(sDg, s1) + "+" + s2 + uf_印綬_干vs干_五行(sDg, s2)
      Case "丁", "壬" : s1 = "壬" : s2 = "癸" : s = s1 + uf_印綬_干vs干_五行(sDg, s1) + "+" + s2 + uf_印綬_干vs干_五行(sDg, s2)
      Case "戊", "癸" : s1 = "甲" : s2 = "乙" : s = s1 + uf_印綬_干vs干_五行(sDg, s1) + "+" + s2 + uf_印綬_干vs干_五行(sDg, s2)
    End Select
    Return "【" + s + "】"
  End Function
  Public Function uf流年多出的月干_六行(sDg As String, sYg As String) As String
    Dim s1 = "", s2 = "", s = ""
    Select Case sYg
      Case "甲", "庚" : s1 = "丙" : s2 = "丁" : s = s1 + uf_印綬_干vs干_六行_奇隅(sDg, s1) + "+" + s2 + uf_印綬_干vs干_六行_奇隅(sDg, s2)
      Case "乙", "辛" : s1 = "戊" : s2 = "庚" : s = s1 + uf_印綬_干vs干_六行_奇隅(sDg, s1) + "+" + s2 + uf_印綬_干vs干_六行_奇隅(sDg, s2)
      Case "丙", "壬" : s1 = "辛" : s2 = "壬" : s = s1 + uf_印綬_干vs干_六行_奇隅(sDg, s1) + "+" + s2 + uf_印綬_干vs干_六行_奇隅(sDg, s2)
      Case "丁", "癸" : s1 = "癸" : s2 = "己" : s = s1 + uf_印綬_干vs干_六行_奇隅(sDg, s1) + "+" + s2 + uf_印綬_干vs干_六行_奇隅(sDg, s2)
      Case "戊", "己" : s1 = "甲" : s2 = "乙" : s = s1 + uf_印綬_干vs干_六行_奇隅(sDg, s1) + "+" + s2 + uf_印綬_干vs干_六行_奇隅(sDg, s2)
    End Select
    Return "【" + s + "】"
  End Function
  '================================================================
  Public Sub us_dgv支支(ByRef dgv As DataGridView)
    Dim ic, ir, iW, iH As Int32
    Dim cc As DataGridViewColumnCollection

    With dgv
      .ReadOnly = True
      .AutoGenerateColumns = False
      .AllowUserToResizeColumns = False
      .AllowUserToResizeRows = False
      .AllowUserToAddRows = False
      .AllowUserToOrderColumns = False
      .ShowCellToolTips = False '不顯示提示文字
      .ColumnHeadersVisible = False
      .RowHeadersVisible = False
      .RowTemplate.Height = 22
      .Font = New Font(__sfnt_DGV, 9.0!)
      .ForeColor = Color.Maroon
      .ScrollBars = ScrollBars.None

      .BringToFront()
      .Hide()
    End With

    cc = dgv.Columns
    For ic = 0 To 11
      cc.Add("", "") : cc(ic).Width = 44
      cc(ic).ValueType = GetType(System.String)
    Next
    dgv.Rows.Add(12)

    For ir = 0 To 11
      For ic = 0 To 11
        dgv(ic, ir).Value = __s144zz_(ir, ic)
      Next
    Next

    iW = 3
    For ic = 0 To cc.Count - 1 : iW += cc(ic).Width : Next
    iH = dgv.RowTemplate.Height * dgv.RowCount
    dgv.Size = New Size(iW, iH)

  End Sub
  '================================================================
  Public Sub us_dgv60甲子(ByRef dgv As DataGridView)
    Dim ic, ir As Int32, iW = 45, iH = 25
    With dgv
      .AutoGenerateColumns = False
      .AllowUserToOrderColumns = False
      .AllowUserToResizeColumns = False
      .AllowUserToResizeRows = False
      .AllowUserToAddRows = False
      .AllowUserToDeleteRows = False

      .ReadOnly = True
      .ShowCellToolTips = False '不顯示提示文字
      .Font = New Font(__sfnt_DGV, 11.0!)
      .ForeColor = Color.Maroon
      .ScrollBars = ScrollBars.None

      .ColumnHeadersVisible = False
      .RowHeadersVisible = False
      .RowTemplate.Height = iH
      .Hide()

      For ic = 0 To 9 : .Columns.Add("", "") : .Columns(ic).Width = iW : Next
      .Rows.Add(6) ' 增加 6 列

      For ir = 0 To 5 : For ic = 0 To 9 : dgv(ic, ir).Value = __s60甲子_(ir, ic) : Next : Next
      .Size = New Size(iW * 10, iH * 6)
    End With
  End Sub
  '================================================================
  Public Function uf天干五行(sG As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Select Case sG
      Case "甲", "乙" : Return "木"
      Case "丙", "丁" : Return "火"
      Case "戊", "己" : Return "土"
      Case "庚", "辛" : Return "金"
      Case "壬", "癸" : Return "水"
    End Select
    Return "木"
  End Function
  Public Function uf地支五行(sZ As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Select Case sZ
      Case "寅", "卯" : Return "木"
      Case "巳", "午" : Return "火"
      Case "丑", "辰", "未", "戌" : Return "土"
      Case "申", "酉" : Return "金"
      Case "亥", "子" : Return "水"
    End Select
    Return "木"
  End Function
  Public Function uf地支六行(sZ As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Select Case sZ
      Case "卯", "辰" : Return "木"
      Case "巳", "午" : Return "火"
      Case "未", "申" : Return "土"
      Case "酉", "戌" : Return "金"
      Case "亥", "子" : Return "水"
      Case "丑", "寅" : Return "元"
    End Select
    Return "木"
  End Function
  '================================================================
  Public Function uf父兄仔妻官_世干vs爻干(s世干 As String, s爻干 As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    'Dim i世G = __dic干(s世干), i爻G = __dic干(s爻干)
    Dim s = ""

    Select Case s世干
      Case "子", "丑"
        Select Case s爻干
          Case "子", "丑" : s = "兄" ' 比兄,劫兄
          Case "寅", "卯" : s = "仔" ' 傷仔,食仔
          Case "辰", "巳" : s = "妻" ' 才妻,财妻
          Case "午", "未" : s = "官" ' 杀官,官官
          Case "申", "酉" : s = "父" ' 綬父,印父
        End Select

      Case "寅", "卯"
        Select Case s爻干
          Case "子", "丑" : s = "父"
          Case "寅", "卯" : s = "兄"
          Case "辰", "巳" : s = "仔"
          Case "午", "未" : s = "妻"
          Case "申", "酉" : s = "官"
        End Select

      Case "辰", "巳"
        Select Case s爻干
          Case "子", "丑" : s = "官"
          Case "寅", "卯" : s = "父"
          Case "辰", "巳" : s = "兄"
          Case "午", "未" : s = "仔"
          Case "申", "酉" : s = "妻"
        End Select

      Case "午", "未"
        Select Case s爻干
          Case "子", "丑" : s = "妻"
          Case "寅", "卯" : s = "官"
          Case "辰", "巳" : s = "父"
          Case "午", "未" : s = "兄"
          Case "申", "酉" : s = "仔"
        End Select

      Case "申", "酉"
        Select Case s爻干
          Case "子", "丑" : s = "仔"
          Case "寅", "卯" : s = "妻"
          Case "辰", "巳" : s = "官"
          Case "午", "未" : s = "父"
          Case "申", "酉" : s = "兄"
        End Select
    End Select
    Return s
  End Function
  Public Function uf父兄仔妻官_五行vs爻干(s五行 As String, s爻干 As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Dim s = ""

    Select Case s五行
      Case "木"
        Select Case s爻干
          Case "子", "丑" : s = "兄" '比兄,劫兄
          Case "寅", "卯" : s = "仔" '傷仔,食仔
          Case "辰", "巳" : s = "妻" '才妻,财妻
          Case "午", "未" : s = "官" '杀官,官官
          Case "申", "酉" : s = "父" '綬父,印父
        End Select

      Case "火"
        Select Case s爻干
          Case "子", "丑" : s = "父"
          Case "寅", "卯" : s = "兄"
          Case "辰", "巳" : s = "仔"
          Case "午", "未" : s = "妻"
          Case "申", "酉" : s = "官"
        End Select

      Case "土"
        Select Case s爻干
          Case "子", "丑" : s = "官"
          Case "寅", "卯" : s = "父"
          Case "辰", "巳" : s = "兄"
          Case "午", "未" : s = "仔"
          Case "申", "酉" : s = "妻"
        End Select

      Case "金"
        Select Case s爻干
          Case "子", "丑" : s = "妻"
          Case "寅", "卯" : s = "官"
          Case "辰", "巳" : s = "父"
          Case "午", "未" : s = "兄"
          Case "申", "酉" : s = "仔"
        End Select
      Case "水"
        Select Case s爻干
          Case "子", "丑" : s = "仔"
          Case "寅", "卯" : s = "妻"
          Case "辰", "巳" : s = "官"
          Case "午", "未" : s = "父"
          Case "申", "酉" : s = "兄"
        End Select
    End Select
    Return s
  End Function
  Public Function uf父兄仔妻官_五行vs爻支(s五行 As String, s爻支 As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    'A=卦宮支＝世支, B=爻支=我
    '辰.土  已.火  午.火  未.土
    '卯.木               申.金
    '寅.木               酉.金
    '丑.土  子.水  亥.水  戌.土
    '------------------------------
    ' 我=爻支：我生[卦宮](父)，我同[卦宮](兄)，我克[卦宮](官)
    '         [卦宮]生我(仔)，[卦宮]克我(妻)
    Dim s = ""

    Select Case s五行
      Case "木"
        Select Case s爻支
          Case "寅", "卯" : s = "兄"
          Case "巳", "午" : s = "仔"
          Case "丑", "辰", "未", "戌" : s = "妻"
          Case "申", "酉" : s = "官"
          Case "亥", "子" : s = "父"
        End Select

      Case "火"
        Select Case s爻支
          Case "寅", "卯" : s = "父"
          Case "巳", "午" : s = "兄"
          Case "丑", "辰", "未", "戌" : s = "仔"
          Case "申", "酉" : s = "妻"
          Case "亥", "子" : s = "官"
        End Select

      Case "土"
        Select Case s爻支
          Case "寅", "卯" : s = "官"
          Case "巳", "午" : s = "父"
          Case "丑", "辰", "未", "戌" : s = "兄"
          Case "申", "酉" : s = "仔"
          Case "亥", "子" : s = "妻"
        End Select

      Case "金"
        Select Case s爻支
          Case "寅", "卯" : s = "妻"
          Case "巳", "午" : s = "官"
          Case "丑", "辰", "未", "戌" : s = "父"
          Case "申", "酉" : s = "兄"
          Case "亥", "子" : s = "仔"
        End Select

      Case "水"
        Select Case s爻支
          Case "寅", "卯" : s = "仔"
          Case "巳", "午" : s = "妻"
          Case "丑", "辰", "未", "戌" : s = "官"
          Case "申", "酉" : s = "父"
          Case "亥", "子" : s = "兄"
        End Select
    End Select
    Return s
  End Function
  Public Function uf父兄仔妻官_卦支vs爻支_五行(sAz As String, sBz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    'A=卦宮支＝世支=我, B=爻支
    '辰.土  已.火  午.火  未.土
    '卯.木               申.金
    '寅.木               酉.金
    '丑.土  子.水  亥.水  戌.土
    '------------------------------
    ' 我=爻支：我生[卦宮](父)，我同[卦宮](兄)，我克[卦宮](官)
    '         [卦宮]生我(仔)，[卦宮]克我(妻)
    If sAz.Length = 0 Then Return "　　"
    Dim s = ""

    Select Case sAz
      Case "亥"  '卦宮，我=爻
        Select Case sBz
          Case "子" : s = "+兄" : Case "亥" : s = "-兄"
          Case "寅" : s = "+仔" : Case "卯" : s = "-仔"
          Case "午" : s = "+妻" : Case "巳" : s = "-妻"
          Case "申" : s = "+父" : Case "酉" : s = "-父"
          Case "辰" : s = "+官" : Case "丑" : s = "-官"
          Case "戌" : s = "+官" : Case "未" : s = "-官"
        End Select
      Case "子"
        Select Case sBz
          Case "子" : s = "-兄" : Case "亥" : s = "+兄"
          Case "寅" : s = "-仔" : Case "卯" : s = "+仔"
          Case "午" : s = "-妻" : Case "巳" : s = "+妻"
          Case "申" : s = "-父" : Case "酉" : s = "+父"
          Case "辰" : s = "-官" : Case "丑" : s = "+官"
          Case "戌" : s = "-官" : Case "未" : s = "+官"
        End Select
      Case "寅"
        Select Case sBz
          Case "子" : s = "-父" : Case "亥" : s = "+父"
          Case "寅" : s = "-兄" : Case "卯" : s = "+兄"
          Case "午" : s = "-仔" : Case "巳" : s = "+仔"
          Case "申" : s = "-官" : Case "酉" : s = "+官"
          Case "辰" : s = "-妻" : Case "丑" : s = "+妻"
          Case "戌" : s = "-妻" : Case "未" : s = "+妻"
        End Select
      Case "卯"
        Select Case sBz
          Case "子" : s = "+父" : Case "亥" : s = "-父"
          Case "寅" : s = "+兄" : Case "卯" : s = "-兄"
          Case "午" : s = "+仔" : Case "巳" : s = "-仔"
          Case "申" : s = "+官" : Case "酉" : s = "-官"
          Case "辰" : s = "+妻" : Case "丑" : s = "-妻"
          Case "戌" : s = "+妻" : Case "未" : s = "-妻"
        End Select
      Case "巳"
        Select Case sBz
          Case "子" : s = "+官" : Case "亥" : s = "-官"
          Case "寅" : s = "+父" : Case "卯" : s = "-父"
          Case "午" : s = "+兄" : Case "巳" : s = "-兄"
          Case "申" : s = "+妻" : Case "酉" : s = "-妻"
          Case "辰" : s = "+仔" : Case "丑" : s = "-仔"
          Case "戌" : s = "+仔" : Case "未" : s = "-仔"
        End Select
      Case "午"
        Select Case sBz
          Case "子" : s = "-官" : Case "亥" : s = "+官"
          Case "寅" : s = "-父" : Case "卯" : s = "+父"
          Case "午" : s = "-兄" : Case "巳" : s = "+兄"
          Case "申" : s = "-妻" : Case "酉" : s = "+妻"
          Case "辰" : s = "-仔" : Case "丑" : s = "+仔"
          Case "戌" : s = "-仔" : Case "未" : s = "+仔"
        End Select
      Case "申"
        Select Case sBz
          Case "子" : s = "-仔" : Case "亥" : s = "+仔"
          Case "寅" : s = "-妻" : Case "卯" : s = "+妻"
          Case "午" : s = "-官" : Case "巳" : s = "+官"
          Case "申" : s = "-兄" : Case "酉" : s = "+兄"
          Case "辰" : s = "-父" : Case "丑" : s = "+父"
          Case "戌" : s = "-父" : Case "未" : s = "+父"
        End Select
      Case "酉"
        Select Case sBz
          Case "子" : s = "+仔" : Case "亥" : s = "-仔"
          Case "寅" : s = "+妻" : Case "卯" : s = "-妻"
          Case "午" : s = "+官" : Case "巳" : s = "-官"
          Case "申" : s = "+兄" : Case "酉" : s = "-兄"
          Case "辰" : s = "+父" : Case "丑" : s = "-父"
          Case "戌" : s = "+父" : Case "未" : s = "-父"
        End Select

      Case "丑"
        Select Case sBz
          Case "子" : s = "+妻" : Case "亥" : s = "-妻"
          Case "寅" : s = "+官" : Case "卯" : s = "-官"
          Case "午" : s = "+父" : Case "巳" : s = "-父"
          Case "申" : s = "+仔" : Case "酉" : s = "-仔"
          Case "辰" : s = "+兄" : Case "丑" : s = "-兄"
          Case "戌" : s = "+兄" : Case "未" : s = "-兄"
        End Select
      Case "辰"
        Select Case sBz
          Case "子" : s = "-妻" : Case "亥" : s = "+妻"
          Case "寅" : s = "-官" : Case "卯" : s = "+官"
          Case "午" : s = "-父" : Case "巳" : s = "+父"
          Case "申" : s = "-仔" : Case "酉" : s = "+仔"
          Case "辰" : s = "-兄" : Case "丑" : s = "+兄"
          Case "戌" : s = "-兄" : Case "未" : s = "+兄"
        End Select
      Case "未"
        Select Case sBz
          Case "子" : s = "+妻" : Case "亥" : s = "-妻"
          Case "寅" : s = "+官" : Case "卯" : s = "-官"
          Case "午" : s = "+父" : Case "巳" : s = "-父"
          Case "申" : s = "+仔" : Case "酉" : s = "-仔"
          Case "辰" : s = "+兄" : Case "丑" : s = "-兄"
          Case "戌" : s = "+兄" : Case "未" : s = "-兄"
        End Select
      Case "戌" '我=子
        Select Case sBz
          Case "子" : s = "-妻" : Case "亥" : s = "+妻"
          Case "寅" : s = "-官" : Case "卯" : s = "+官"
          Case "午" : s = "-父" : Case "巳" : s = "+父"
          Case "巳" : s = "-仔" : Case "酉" : s = "+仔"
          Case "辰" : s = "-兄" : Case "丑" : s = "+兄"
          Case "戌" : s = "-兄" : Case "未" : s = "+兄"
        End Select
    End Select
    Return s
  End Function
  '================================================================
  Public Function uf六親_支vs支_六行_丑寅元(sAz As String, sBz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    ' 辰.木  已.火  午.火  未.土
    ' 卯.木               申.土  
    ' 寅.元               酉.金
    ' 丑.元  子.水  亥.水  戌.金
    '------------------------------
    ' 父兄妻
    ' 仔奴官
    Dim s = ""

    Select Case sAz ' 卦宮(生年)
      Case "亥" ' 水爻(流年)
        Select Case sBz
          Case "亥" : s = "兄" : Case "子" : s = "兄" ' 水
          Case "丑" : s = "妻" : Case "寅" : s = "妻" ' 元
          Case "卯" : s = "仔" : Case "辰" : s = "仔" ' 木
          Case "巳" : s = "奴" : Case "午" : s = "奴" ' 火
          Case "未" : s = "官" : Case "申" : s = "官" ' 土
          Case "酉" : s = "父" : Case "戌" : s = "父" ' 金
        End Select
      Case "子"
        Select Case sBz
          Case "亥" : s = "兄" : Case "子" : s = "兄"
          Case "丑" : s = "妻" : Case "寅" : s = "妻"
          Case "卯" : s = "仔" : Case "辰" : s = "仔"
          Case "巳" : s = "奴" : Case "午" : s = "奴"
          Case "未" : s = "官" : Case "申" : s = "官"
          Case "酉" : s = "父" : Case "戌" : s = "父"
        End Select

      Case "丑"
        ' 爻生[宮](父)，爻同[宮](兄)，爻克[宮](官)，爻沖[宮](奴)
        ' [宮]生爻(妻)，[宮]克爻(仔)
        ' 辰.木  已.火  午.火  未.土
        ' 卯.木               申.土  
        ' 寅.元               酉.金
        ' 丑.元  子.水  亥.水  戌.金
        Select Case sBz
          Case "亥" : s = "父" : Case "子" : s = "父"
          Case "丑" : s = "兄" : Case "寅" : s = "兄"
          Case "卯" : s = "妻" : Case "辰" : s = "妻"
          Case "巳" : s = "仔" : Case "午" : s = "仔"
          Case "未" : s = "奴" : Case "申" : s = "奴"
          Case "酉" : s = "官" : Case "戌" : s = "官"
        End Select
      Case "寅"
        Select Case sBz
          Case "亥" : s = "父" : Case "子" : s = "父"
          Case "丑" : s = "兄" : Case "寅" : s = "兄"
          Case "卯" : s = "妻" : Case "辰" : s = "妻"
          Case "巳" : s = "仔" : Case "午" : s = "仔"
          Case "未" : s = "奴" : Case "申" : s = "奴"
          Case "酉" : s = "官" : Case "戌" : s = "官"
        End Select

      Case "卯"
        Select Case sBz
          Case "亥" : s = "官" : Case "子" : s = "官"
          Case "丑" : s = "父" : Case "寅" : s = "父"
          Case "卯" : s = "兄" : Case "辰" : s = "兄"
          Case "巳" : s = "妻" : Case "午" : s = "妻"
          Case "未" : s = "仔" : Case "申" : s = "仔"
          Case "酉" : s = "奴" : Case "戌" : s = "奴"
        End Select
      Case "辰"
        Select Case sBz
          Case "亥" : s = "官" : Case "子" : s = "官"
          Case "丑" : s = "父" : Case "寅" : s = "父"
          Case "卯" : s = "兄" : Case "辰" : s = "兄"
          Case "巳" : s = "妻" : Case "午" : s = "妻"
          Case "未" : s = "仔" : Case "申" : s = "仔"
          Case "酉" : s = "奴" : Case "戌" : s = "奴"
        End Select

      Case "巳"
        Select Case sBz
          Case "亥" : s = "奴" : Case "子" : s = "奴"
          Case "丑" : s = "官" : Case "寅" : s = "官"
          Case "卯" : s = "父" : Case "辰" : s = "父"
          Case "巳" : s = "兄" : Case "午" : s = "兄"
          Case "未" : s = "妻" : Case "申" : s = "妻"
          Case "酉" : s = "仔" : Case "戌" : s = "仔"
        End Select
      Case "午"
        ' 爻生[宮](父)，爻同[宮](兄)，爻克[宮](官)，爻沖[宮](奴)
        ' [宮]生爻(妻)，[宮]克爻(仔)
        ' 辰.木  已.火  午.火  未.土
        ' 卯.木               申.土  
        ' 寅.元               酉.金
        ' 丑.元  子.水  亥.水  戌.金
        Select Case sBz
          Case "亥" : s = "奴" : Case "子" : s = "奴"
          Case "丑" : s = "官" : Case "寅" : s = "官"
          Case "卯" : s = "父" : Case "辰" : s = "父"
          Case "巳" : s = "兄" : Case "午" : s = "兄"
          Case "未" : s = "妻" : Case "申" : s = "妻"
          Case "酉" : s = "仔" : Case "戌" : s = "仔"
            ' 男類排法
            'Case "辰" : s = "父" : Case "巳" : s = "兄" : Case "午" : s = "兄" : Case "未" : s = "妻"
            'Case "卯" : s = "父" : Case "" : s = "　" : Case "" : s = "　" : Case "申" : s = "妻"
            'Case "寅" : s = "官" : Case "" : s = "　" : Case "" : s = "　" : Case "酉" : s = "仔"
            'Case "丑" : s = "官" : Case "子" : s = "奴" : Case "亥" : s = "奴" : Case "戌" : s = "仔"
        End Select
      Case "未"
        Select Case sBz
          Case "亥" : s = "仔" : Case "子" : s = "仔"
          Case "丑" : s = "奴" : Case "寅" : s = "奴"
          Case "卯" : s = "官" : Case "辰" : s = "官"
          Case "巳" : s = "父" : Case "午" : s = "父"
          Case "未" : s = "兄" : Case "申" : s = "兄"
          Case "酉" : s = "妻" : Case "戌" : s = "妻"
        End Select
      Case "申"
        Select Case sBz
          Case "亥" : s = "仔" : Case "子" : s = "仔"
          Case "丑" : s = "奴" : Case "寅" : s = "奴"
          Case "卯" : s = "官" : Case "辰" : s = "官"
          Case "巳" : s = "父" : Case "午" : s = "父"
          Case "未" : s = "兄" : Case "申" : s = "兄"
          Case "酉" : s = "妻" : Case "戌" : s = "妻"
        End Select
      Case "酉"
        Select Case sBz
          Case "亥" : s = "妻" : Case "子" : s = "妻"
          Case "丑" : s = "仔" : Case "寅" : s = "仔"
          Case "卯" : s = "奴" : Case "辰" : s = "奴"
          Case "巳" : s = "官" : Case "午" : s = "官"
          Case "未" : s = "父" : Case "申" : s = "父"
          Case "酉" : s = "兄" : Case "戌" : s = "兄"
        End Select
      Case "戌"
        Select Case sBz
          Case "亥" : s = "妻" : Case "子" : s = "妻"
          Case "丑" : s = "仔" : Case "寅" : s = "仔"
          Case "卯" : s = "奴" : Case "辰" : s = "奴"
          Case "巳" : s = "官" : Case "午" : s = "官"
          Case "未" : s = "父" : Case "申" : s = "父"
          Case "酉" : s = "兄" : Case "戌" : s = "兄"
        End Select
    End Select
    Return s
  End Function
  Public Function uf六親_支vs支_六行_丑戌元(sAz As String, sBz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    ' 辰.土  已.火  午.火  未.土
    ' 卯.木               申.金
    ' 寅.木               酉.金
    ' 丑.元  子.水  亥.水  戌.元
    '------------------------------
    ' 父兄妻　　木火土
    ' 仔奴官　　金水元
    Dim s = ""

    Select Case sAz ' 卦宮(生年)
      Case "亥" ' 水爻(流年)
        Select Case sBz
          Case "亥" : s = "-兄" : Case "子" : s = "+兄" ' 水
          Case "丑" : s = "-妻" : Case "戌" : s = "+妻" ' 元
          Case "寅" : s = "+仔" : Case "卯" : s = "-仔" ' 木
          Case "巳" : s = "-奴" : Case "午" : s = "+奴" ' 火
          Case "辰" : s = "+官" : Case "未" : s = "-官" ' 土
          Case "申" : s = "+父" : Case "酉" : s = "-父" ' 金
        End Select
      Case "子"
        Select Case sBz
          Case "亥" : s = "+兄" : Case "子" : s = "-兄"
          Case "丑" : s = "+妻" : Case "戌" : s = "-妻"
          Case "寅" : s = "-仔" : Case "卯" : s = "+仔"
          Case "巳" : s = "+奴" : Case "午" : s = "-奴"
          Case "辰" : s = "-官" : Case "未" : s = "+官"
          Case "申" : s = "-父" : Case "酉" : s = "+父"
        End Select

      Case "寅"
        Select Case sBz
          Case "亥" : s = "+官" : Case "子" : s = "-官" ' 水
          Case "丑" : s = "+父" : Case "戌" : s = "-父" ' 元
          Case "寅" : s = "-兄" : Case "卯" : s = "+兄" ' 木
          Case "巳" : s = "+妻" : Case "午" : s = "-妻" ' 火
          Case "辰" : s = "-仔" : Case "未" : s = "+仔" ' 土
          Case "申" : s = "-奴" : Case "酉" : s = "+奴" ' 金
        End Select
      Case "卯"
        Select Case sBz
          Case "亥" : s = "-官" : Case "子" : s = "+官" ' 水
          Case "丑" : s = "-父" : Case "戌" : s = "+父" ' 元
          Case "寅" : s = "+兄" : Case "卯" : s = "-兄" ' 木
          Case "巳" : s = "-妻" : Case "午" : s = "+妻" ' 火
          Case "辰" : s = "+仔" : Case "未" : s = "-仔" ' 土
          Case "申" : s = "+奴" : Case "酉" : s = "-奴" ' 金
        End Select

      Case "巳" ' 火
        ' 爻生[宮](父)，爻同[宮](兄)，爻克[宮](官)，爻沖[宮](奴)
        ' [宮]生爻(妻)，[宮]克爻(仔)
        Select Case sBz
          Case "亥" : s = "-奴" : Case "子" : s = "+奴" ' 水
          Case "丑" : s = "-官" : Case "戌" : s = "+官" ' 元
          Case "寅" : s = "+父" : Case "卯" : s = "-父" ' 木
          Case "巳" : s = "-兄" : Case "午" : s = "+兄" ' 火
          Case "辰" : s = "+妻" : Case "未" : s = "-妻" ' 土
          Case "申" : s = "+仔" : Case "酉" : s = "-仔" ' 金
        End Select
      Case "午" ' 火
        Select Case sBz
          Case "亥" : s = "+奴" : Case "子" : s = "-奴" ' 水
          Case "丑" : s = "+官" : Case "戌" : s = "-官" ' 元
          Case "寅" : s = "-父" : Case "卯" : s = "+父" ' 木
          Case "巳" : s = "+兄" : Case "午" : s = "-兄" ' 火
          Case "辰" : s = "-妻" : Case "未" : s = "+妻" ' 土
          Case "申" : s = "-仔" : Case "酉" : s = "+仔" ' 金
        End Select

      Case "申" ' 金
        Select Case sBz
          Case "亥" : s = "+妻" : Case "子" : s = "-妻" ' 水
          Case "丑" : s = "+仔" : Case "戌" : s = "-仔" ' 元
          Case "寅" : s = "-奴" : Case "卯" : s = "+奴" ' 木
          Case "巳" : s = "+官" : Case "午" : s = "-官" ' 火
          Case "辰" : s = "-父" : Case "未" : s = "+父" ' 土
          Case "申" : s = "-兄" : Case "酉" : s = "+兄" ' 金
        End Select
      Case "酉" ' 金
        Select Case sBz
          Case "亥" : s = "-妻" : Case "子" : s = "+妻" ' 水
          Case "丑" : s = "-仔" : Case "戌" : s = "+仔" ' 元
          Case "寅" : s = "+奴" : Case "卯" : s = "-奴" ' 木
          Case "巳" : s = "-官" : Case "午" : s = "+官" ' 火
          Case "辰" : s = "+父" : Case "未" : s = "-父" ' 土
          Case "申" : s = "+兄" : Case "酉" : s = "-兄" ' 金
        End Select
        '---------------------------------------------------
      Case "辰" ＇ 土
        Select Case sBz
          Case "亥" : s = "+仔" : Case "子" : s = "-仔" ' 水
          Case "丑" : s = "+奴" : Case "戌" : s = "-奴" ' 元
          Case "寅" : s = "-官" : Case "卯" : s = "+官" ' 木
          Case "巳" : s = "+父" : Case "午" : s = "-父" ' 火
          Case "辰" : s = "-兄" : Case "未" : s = "+兄" ' 土
          Case "申" : s = "-妻" : Case "酉" : s = "+妻" ' 金
        End Select
      Case "未" ＇ 土
        Select Case sBz
          Case "亥" : s = "-仔" : Case "子" : s = "+仔" ' 水
          Case "丑" : s = "-奴" : Case "戌" : s = "+奴" ' 元
          Case "寅" : s = "+官" : Case "卯" : s = "-官" ' 木
          Case "巳" : s = "-父" : Case "午" : s = "+父" ' 火
          Case "辰" : s = "+兄" : Case "未" : s = "-兄" ' 土
          Case "申" : s = "+妻" : Case "酉" : s = "-妻" ' 金
        End Select
        '---------------------------------------------------
      Case "丑" ' 元
        Select Case sBz
          Case "亥" : s = "-父" : Case "子" : s = "+父" ' 水
          Case "丑" : s = "-兄" : Case "戌" : s = "+兄" ' 元
          Case "寅" : s = "+妻" : Case "卯" : s = "-妻" ' 木
          Case "巳" : s = "-仔" : Case "午" : s = "+仔" ' 火
          Case "辰" : s = "+奴" : Case "未" : s = "-奴" ' 土
          Case "申" : s = "+官" : Case "酉" : s = "-官" ' 金
        End Select
      Case "戌" ' 元
        Select Case sBz
          Case "亥" : s = "+父" : Case "子" : s = "-父" ' 水
          Case "丑" : s = "+兄" : Case "戌" : s = "-兄" ' 元
          Case "寅" : s = "-妻" : Case "卯" : s = "+妻" ' 木
          Case "巳" : s = "+仔" : Case "午" : s = "-仔" ' 火
          Case "辰" : s = "-奴" : Case "未" : s = "+奴" ' 土
          Case "申" : s = "-官" : Case "酉" : s = "+官" ' 金
        End Select
    End Select
    Return s
  End Function
  Public Function uf六親_支vs支_六行_丑辰元(sAz As String, sBz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    ' 辰.元  已.火  午.火  未.土
    ' 卯.木               申.金
    ' 寅.木               酉.金
    ' 丑.元  子.水  亥.水  戌.土
    '------------------------------
    ' 父兄妻　　木火土
    ' 仔奴官　　金水元
    Dim s = ""

    Select Case sAz ' 卦宮(生年)
      Case "亥" ' 水爻(流年)
        Select Case sBz
          Case "亥" : s = "-兄" : Case "子" : s = "+兄" ' 水
          Case "丑" : s = "-仔" : Case "辰" : s = "+仔" ' 元
          Case "寅" : s = "+妻" : Case "卯" : s = "-妻" ' 木
          Case "巳" : s = "-奴" : Case "午" : s = "+奴" ' 火
          Case "未" : s = "-官" : Case "戌" : s = "+官" ' 土
          Case "申" : s = "+父" : Case "酉" : s = "-父" ' 金
        End Select
      Case "子"
        Select Case sBz
          Case "亥" : s = "+兄" : Case "子" : s = "-兄"
          Case "丑" : s = "+仔" : Case "辰" : s = "-仔"
          Case "寅" : s = "-妻" : Case "卯" : s = "+妻"
          Case "巳" : s = "+奴" : Case "午" : s = "-奴"
          Case "未" : s = "+官" : Case "戌" : s = "-官"
          Case "申" : s = "-父" : Case "酉" : s = "+父"
        End Select

      Case "寅"
        Select Case sBz
          Case "亥" : s = "+官" : Case "子" : s = "-官" ' 水
          Case "丑" : s = "+父" : Case "辰" : s = "-父" ' 元
          Case "寅" : s = "-兄" : Case "卯" : s = "+兄" ' 木
          Case "巳" : s = "+仔" : Case "午" : s = "-仔" ' 火
          Case "未" : s = "+妻" : Case "戌" : s = "-妻" ' 土
          Case "申" : s = "-奴" : Case "酉" : s = "+奴" ' 金
        End Select
      Case "卯"
        Select Case sBz
          Case "亥" : s = "-官" : Case "子" : s = "+官" ' 水
          Case "丑" : s = "-父" : Case "辰" : s = "+父" ' 元
          Case "寅" : s = "+兄" : Case "卯" : s = "-兄" ' 木
          Case "巳" : s = "-仔" : Case "午" : s = "+仔" ' 火
          Case "未" : s = "-妻" : Case "戌" : s = "+妻" ' 土
          Case "申" : s = "+奴" : Case "酉" : s = "-奴" ' 金
        End Select

      Case "巳" ' 火
        ' 爻生[宮](父)，爻同[宮](兄)，爻克[宮](官)，爻沖[宮](奴)
        ' [宮]生爻(妻)，[宮]克爻(仔)
        Select Case sBz
          Case "亥" : s = "-奴" : Case "子" : s = "+奴" ' 水
          Case "丑" : s = "-官" : Case "辰" : s = "+官" ' 元
          Case "寅" : s = "+父" : Case "卯" : s = "-父" ' 木
          Case "巳" : s = "-兄" : Case "午" : s = "+兄" ' 火
          Case "未" : s = "-仔" : Case "戌" : s = "+仔" ' 土
          Case "申" : s = "+妻" : Case "酉" : s = "-妻" ' 金
        End Select
      Case "午" ' 火
        Select Case sBz
          Case "亥" : s = "+奴" : Case "子" : s = "-奴" ' 水
          Case "丑" : s = "+官" : Case "辰" : s = "-官" ' 元
          Case "寅" : s = "-父" : Case "卯" : s = "+父" ' 木
          Case "巳" : s = "+兄" : Case "午" : s = "-兄" ' 火
          Case "未" : s = "+仔" : Case "戌" : s = "-仔" ' 土
          Case "申" : s = "-妻" : Case "酉" : s = "+妻" ' 金
        End Select

      Case "申" ' 金
        Select Case sBz
          Case "亥" : s = "+仔" : Case "子" : s = "-仔" ' 水
          Case "丑" : s = "+妻" : Case "辰" : s = "-妻" ' 元
          Case "寅" : s = "-奴" : Case "卯" : s = "+奴" ' 木
          Case "巳" : s = "+官" : Case "午" : s = "-官" ' 火
          Case "未" : s = "+父" : Case "戌" : s = "-父" ' 土
          Case "申" : s = "-兄" : Case "酉" : s = "+兄" ' 金
        End Select
      Case "酉" ' 金
        Select Case sBz
          Case "亥" : s = "-仔" : Case "子" : s = "+仔" ' 水
          Case "丑" : s = "-妻" : Case "辰" : s = "+妻" ' 元
          Case "寅" : s = "+奴" : Case "卯" : s = "-奴" ' 木
          Case "巳" : s = "-官" : Case "午" : s = "+官" ' 火
          Case "未" : s = "-父" : Case "戌" : s = "+父" ' 土
          Case "申" : s = "+兄" : Case "酉" : s = "-兄" ' 金
        End Select
        '---------------------------------------------------
      Case "丑" ' 元
        Select Case sBz
          Case "亥" : s = "-父" : Case "子" : s = "+父" ' 水
          Case "丑" : s = "-兄" : Case "辰" : s = "+兄" ' 元
          Case "寅" : s = "+仔" : Case "卯" : s = "-仔" ' 木
          Case "巳" : s = "-妻" : Case "午" : s = "+妻" ' 火
          Case "未" : s = "-奴" : Case "戌" : s = "+奴" ' 土
          Case "申" : s = "+官" : Case "酉" : s = "-官" ' 金
        End Select
      Case "辰" ＇ 元
        Select Case sBz
          Case "亥" : s = "+父" : Case "子" : s = "-父" ' 水
          Case "丑" : s = "+兄" : Case "辰" : s = "-兄" ' 元
          Case "寅" : s = "-仔" : Case "卯" : s = "+仔" ' 木
          Case "巳" : s = "+妻" : Case "午" : s = "-妻" ' 火
          Case "未" : s = "+奴" : Case "戌" : s = "-奴" ' 土
          Case "申" : s = "-官" : Case "酉" : s = "+官" ' 金
        End Select
        '---------------------------------------------------
      Case "未" ＇ 燥土
        Select Case sBz
          Case "亥" : s = "-妻" : Case "子" : s = "+妻" ' 水
          Case "丑" : s = "-奴" : Case "辰" : s = "+奴" ' 元
          Case "寅" : s = "+官" : Case "卯" : s = "-官" ' 木
          Case "巳" : s = "-父" : Case "午" : s = "+父" ' 火
          Case "未" : s = "-兄" : Case "戌" : s = "+兄" ' 土
          Case "申" : s = "+仔" : Case "酉" : s = "-仔" ' 金
        End Select
      Case "戌" ' 燥土
        Select Case sBz
          Case "亥" : s = "+妻" : Case "子" : s = "-妻" ' 水
          Case "丑" : s = "+奴" : Case "辰" : s = "-奴" ' 元
          Case "寅" : s = "-官" : Case "卯" : s = "+官" ' 木
          Case "巳" : s = "+父" : Case "午" : s = "-父" ' 火
          Case "未" : s = "+兄" : Case "戌" : s = "-兄" ' 土
          Case "申" : s = "-仔" : Case "酉" : s = "+仔" ' 金
        End Select
    End Select

    Return us父2梟印(s)
    'Return s
  End Function

  Public Function uf六親_支vs支_六行_辰未元(sAz As String, sBz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    ' 辰.元  已.火  午.火  未.元
    ' 卯.木               申.金
    ' 寅.木               酉.金
    ' 丑.土  子.水  亥.水  戌.土
    '------------------------------
    ' 父兄妻　　木火土
    ' 仔奴官　　金水元
    Dim s = ""

    Select Case sAz ' 卦宮(生年)
      Case "亥" ' 水爻(流年)
        Select Case sBz
          Case "子" : s = "+兄" : Case "亥" : s = "-兄" ' 水
          Case "辰" : s = "+妻" : Case "未" : s = "-妻" ' 元
          Case "寅" : s = "+仔" : Case "卯" : s = "-仔" ' 木
          Case "午" : s = "+奴" : Case "巳" : s = "-奴" ' 火
          Case "戌" : s = "+官" : Case "丑" : s = "-官" ' 土
          Case "申" : s = "+父" : Case "酉" : s = "-父" ' 金
        End Select
      Case "子"
        Select Case sBz
          Case "子" : s = "-兄" : Case "亥" : s = "+兄"
          Case "辰" : s = "-妻" : Case "未" : s = "+妻"
          Case "寅" : s = "-仔" : Case "卯" : s = "+仔"
          Case "午" : s = "-奴" : Case "巳" : s = "+奴"
          Case "戌" : s = "-官" : Case "丑" : s = "+官"
          Case "申" : s = "-父" : Case "酉" : s = "+父"
        End Select

      Case "寅"
        Select Case sBz
          Case "子" : s = "-官" : Case "亥" : s = "+官"
          Case "辰" : s = "-父" : Case "未" : s = "+父"
          Case "寅" : s = "-兄" : Case "卯" : s = "+兄"
          Case "午" : s = "-妻" : Case "巳" : s = "+妻"
          Case "戌" : s = "-仔" : Case "丑" : s = "+仔"
          Case "申" : s = "-奴" : Case "酉" : s = "+奴"
        End Select
      Case "卯"
        Select Case sBz
          Case "子" : s = "+官" : Case "亥" : s = "-官"
          Case "辰" : s = "+父" : Case "未" : s = "-父"
          Case "寅" : s = "+兄" : Case "卯" : s = "-兄"
          Case "午" : s = "+妻" : Case "巳" : s = "-妻"
          Case "戌" : s = "+仔" : Case "丑" : s = "-仔"
          Case "申" : s = "+奴" : Case "酉" : s = "-奴"
        End Select

      Case "巳" ' 火
        ' 爻生[宮](父)，爻同[宮](兄)，爻克[宮](官)，爻沖[宮](奴)
        ' [宮]生爻(妻)，[宮]克爻(仔)
        Select Case sBz
          Case "子" : s = "+奴" : Case "亥" : s = "-奴"
          Case "辰" : s = "+官" : Case "未" : s = "-官"
          Case "寅" : s = "+父" : Case "卯" : s = "-父"
          Case "午" : s = "+兄" : Case "巳" : s = "-兄"
          Case "戌" : s = "+妻" : Case "丑" : s = "-妻"
          Case "申" : s = "+仔" : Case "酉" : s = "-仔"
        End Select
      Case "午" ' 火
        Select Case sBz
          Case "子" : s = "-奴" : Case "亥" : s = "+奴"
          Case "辰" : s = "-官" : Case "未" : s = "+官"
          Case "寅" : s = "-父" : Case "卯" : s = "+父"
          Case "午" : s = "-兄" : Case "巳" : s = "+兄"
          Case "戌" : s = "-妻" : Case "丑" : s = "+妻"
          Case "申" : s = "-仔" : Case "酉" : s = "+仔"
        End Select

      Case "申" ' 金
        Select Case sBz
          Case "子" : s = "-妻" : Case "亥" : s = "+妻"
          Case "辰" : s = "-仔" : Case "未" : s = "+仔"
          Case "寅" : s = "-奴" : Case "卯" : s = "+奴"
          Case "午" : s = "-官" : Case "巳" : s = "+官"
          Case "戌" : s = "-父" : Case "丑" : s = "+父"
          Case "申" : s = "-兄" : Case "酉" : s = "+兄"
        End Select
      Case "酉" ' 金
        Select Case sBz
          Case "子" : s = "+妻" : Case "亥" : s = "-妻"
          Case "辰" : s = "+仔" : Case "未" : s = "-仔"
          Case "寅" : s = "+奴" : Case "卯" : s = "-奴"
          Case "午" : s = "+官" : Case "巳" : s = "-官"
          Case "戌" : s = "+父" : Case "丑" : s = "-父"
          Case "申" : s = "+兄" : Case "酉" : s = "-兄"
        End Select
        '---------------------------------------------------
      Case "丑" ' 土
        Select Case sBz
          Case "子" : s = "+仔" : Case "亥" : s = "-仔"
          Case "辰" : s = "+奴" : Case "未" : s = "-奴"
          Case "寅" : s = "+官" : Case "卯" : s = "-官"
          Case "午" : s = "+父" : Case "巳" : s = "-父"
          Case "戌" : s = "+兄" : Case "丑" : s = "-兄"
          Case "申" : s = "+妻" : Case "酉" : s = "-妻"
        End Select
      Case "戌" ' 土
        Select Case sBz
          Case "子" : s = "-仔" : Case "亥" : s = "+仔"
          Case "辰" : s = "-奴" : Case "未" : s = "+奴"
          Case "寅" : s = "-官" : Case "卯" : s = "+官"
          Case "午" : s = "-父" : Case "巳" : s = "+父"
          Case "戌" : s = "-兄" : Case "丑" : s = "+兄"
          Case "申" : s = "-妻" : Case "酉" : s = "+妻"
        End Select
        '---------------------------------------------------
      Case "辰" ＇ 元
        Select Case sBz
          Case "子" : s = "-父" : Case "亥" : s = "+父"
          Case "辰" : s = "-兄" : Case "未" : s = "+兄"
          Case "寅" : s = "-妻" : Case "卯" : s = "+妻"
          Case "午" : s = "+仔" : Case "巳" : s = "+仔"
          Case "戌" : s = "-奴" : Case "丑" : s = "+奴"
          Case "申" : s = "-官" : Case "酉" : s = "+官"
        End Select
      Case "未" ＇ 燥土
        Select Case sBz
          Case "子" : s = "+父" : Case "亥" : s = "-父"
          Case "辰" : s = "+兄" : Case "未" : s = "-兄"
          Case "寅" : s = "+妻" : Case "卯" : s = "-妻"
          Case "午" : s = "+仔" : Case "巳" : s = "-仔"
          Case "戌" : s = "+奴" : Case "丑" : s = "-奴"
          Case "申" : s = "+官" : Case "酉" : s = "-官"
        End Select
    End Select
    Return s
  End Function
  Public Function uf六親_支vs支_六行_未戌元(sAz As String, sBz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    ' 辰.土  已.火  午.火  未.元
    ' 卯.木               申.金
    ' 寅.木               酉.金
    ' 丑.土  子.水  亥.水  戌.元
    '------------------------------
    ' 父兄妻　　木火土
    ' 仔奴官　　金水元
    Dim s = ""

    Select Case sAz ' 卦宮(生年)
      Case "亥" ' 水爻(流年)
        Select Case sBz
          Case "亥" : s = "-兄" : Case "子" : s = "+兄" ' 水
          Case "未" : s = "-妻" : Case "戌" : s = "+妻" ' 元
          Case "寅" : s = "+仔" : Case "卯" : s = "-仔" ' 木
          Case "巳" : s = "-奴" : Case "午" : s = "+奴" ' 火
          Case "丑" : s = "-官" : Case "辰" : s = "+官" ' 土
          Case "申" : s = "+父" : Case "酉" : s = "-父" ' 金
        End Select
      Case "子"
        Select Case sBz
          Case "亥" : s = "+兄" : Case "子" : s = "-兄"
          Case "未" : s = "+妻" : Case "戌" : s = "-妻"
          Case "寅" : s = "-仔" : Case "卯" : s = "+仔"
          Case "巳" : s = "+奴" : Case "午" : s = "-奴"
          Case "丑" : s = "+官" : Case "辰" : s = "-官"
          Case "申" : s = "-父" : Case "酉" : s = "+父"
        End Select

      Case "寅"
        Select Case sBz
          Case "亥" : s = "+官" : Case "子" : s = "-官" ' 水
          Case "未" : s = "+父" : Case "戌" : s = "-父" ' 元
          Case "寅" : s = "-兄" : Case "卯" : s = "+兄" ' 木
          Case "巳" : s = "+妻" : Case "午" : s = "-妻" ' 火
          Case "丑" : s = "+仔" : Case "辰" : s = "-仔" ' 土
          Case "申" : s = "-奴" : Case "酉" : s = "+奴" ' 金
        End Select
      Case "卯"
        Select Case sBz
          Case "亥" : s = "-官" : Case "子" : s = "+官" ' 水
          Case "未" : s = "-父" : Case "戌" : s = "+父" ' 元
          Case "寅" : s = "+兄" : Case "卯" : s = "-兄" ' 木
          Case "巳" : s = "-妻" : Case "午" : s = "+妻" ' 火
          Case "丑" : s = "-仔" : Case "辰" : s = "+仔" ' 土
          Case "申" : s = "+奴" : Case "酉" : s = "-奴" ' 金
        End Select

      Case "巳" ' 火
        ' 爻生[宮](父)，爻同[宮](兄)，爻克[宮](官)，爻沖[宮](奴)
        ' [宮]生爻(妻)，[宮]克爻(仔)
        Select Case sBz
          Case "亥" : s = "+奴" : Case "子" : s = "-奴" ' 水
          Case "未" : s = "+官" : Case "戌" : s = "-官" ' 元
          Case "寅" : s = "-父" : Case "卯" : s = "+父" ' 木
          Case "巳" : s = "+兄" : Case "午" : s = "-兄" ' 火
          Case "丑" : s = "+妻" : Case "辰" : s = "-妻" ' 土
          Case "申" : s = "-仔" : Case "酉" : s = "+仔" ' 金
        End Select
      Case "午" ' 火
        Select Case sBz
          Case "亥" : s = "+奴" : Case "子" : s = "-奴" ' 水
          Case "未" : s = "+官" : Case "戌" : s = "-官" ' 元
          Case "寅" : s = "-父" : Case "卯" : s = "+父" ' 木
          Case "巳" : s = "+兄" : Case "午" : s = "-兄" ' 火
          Case "丑" : s = "+妻" : Case "辰" : s = "-妻" ' 土
          Case "申" : s = "-仔" : Case "酉" : s = "+仔" ' 金
        End Select

      Case "申" ' 金
        Select Case sBz
          Case "亥" : s = "+妻" : Case "子" : s = "-妻" ' 水
          Case "未" : s = "+仔" : Case "戌" : s = "-仔" ' 元
          Case "寅" : s = "-奴" : Case "卯" : s = "+奴" ' 木
          Case "巳" : s = "+官" : Case "午" : s = "-官" ' 火
          Case "丑" : s = "+父" : Case "辰" : s = "-父" ' 土
          Case "申" : s = "-兄" : Case "酉" : s = "+兄" ' 金
        End Select
      Case "酉" ' 金
        Select Case sBz
          Case "亥" : s = "-妻" : Case "子" : s = "+妻" ' 水
          Case "未" : s = "-仔" : Case "戌" : s = "+仔" ' 元
          Case "寅" : s = "+奴" : Case "卯" : s = "-奴" ' 木
          Case "巳" : s = "-官" : Case "午" : s = "+官" ' 火
          Case "丑" : s = "-父" : Case "辰" : s = "+父" ' 土
          Case "申" : s = "+兄" : Case "酉" : s = "-兄" ' 金
        End Select
        '---------------------------------------------------
      Case "丑" ' 土
        Select Case sBz
          Case "亥" : s = "-仔" : Case "子" : s = "+仔" ' 水
          Case "未" : s = "-奴" : Case "戌" : s = "+奴" ' 元
          Case "寅" : s = "+官" : Case "卯" : s = "-官" ' 木
          Case "巳" : s = "-父" : Case "午" : s = "+父" ' 火
          Case "丑" : s = "-兄" : Case "辰" : s = "+兄" ' 土
          Case "申" : s = "+妻" : Case "酉" : s = "-妻" ' 金
        End Select
      Case "辰" ＇ 土
        Select Case sBz
          Case "亥" : s = "+仔" : Case "子" : s = "-仔" ' 水
          Case "未" : s = "+奴" : Case "戌" : s = "-奴" ' 元
          Case "寅" : s = "-官" : Case "卯" : s = "+官" ' 木
          Case "巳" : s = "+父" : Case "午" : s = "-父" ' 火
          Case "丑" : s = "+兄" : Case "辰" : s = "-兄" ' 土
          Case "申" : s = "-妻" : Case "酉" : s = "+妻" ' 金
        End Select
        '---------------------------------------------------
      Case "未" ＇ 元
        Select Case sBz
          Case "亥" : s = "-父" : Case "子" : s = "+父" ' 水
          Case "未" : s = "-兄" : Case "戌" : s = "+兄" ' 元
          Case "寅" : s = "+妻" : Case "卯" : s = "-妻" ' 木
          Case "巳" : s = "-仔" : Case "午" : s = "+仔" ' 火
          Case "丑" : s = "-奴" : Case "辰" : s = "+奴" ' 土
          Case "申" : s = "+官" : Case "酉" : s = "-官" ' 金
        End Select
      Case "戌" ' 元
        Select Case sBz
          Case "亥" : s = "+父" : Case "子" : s = "-父" ' 水
          Case "未" : s = "+兄" : Case "戌" : s = "-兄" ' 元
          Case "寅" : s = "-妻" : Case "卯" : s = "+妻" ' 木
          Case "巳" : s = "+仔" : Case "午" : s = "-仔" ' 火
          Case "丑" : s = "+奴" : Case "辰" : s = "-奴" ' 土
          Case "申" : s = "-官" : Case "酉" : s = "+官" ' 金
        End Select
    End Select
    Return s
  End Function
  Public Function uf父兄對仔妻官_支vs支_六行(sAz As String, sBz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    'A=世支, B=爻支  
    '  已.火  午.火  未.土  申.土  
    '  辰.木               酉.金
    '  卯.木               戌.金
    '  寅.元  丑.元  子.水  亥.水
    '------------------------------
    ' 父兄对  => 己不用
    ' 仔妻官
    Dim s = ""

    Select Case sAz
      Case "亥"
        Select Case sBz
          Case "亥" : s = "兄" : Case "子" : s = "兄"
          Case "丑" : s = "对" : Case "寅" : s = "对"
          Case "卯" : s = "仔" : Case "辰" : s = "仔"
          Case "巳" : s = "妻" : Case "午" : s = "妻"
          Case "未" : s = "官" : Case "申" : s = "官"
          Case "酉" : s = "父" : Case "戌" : s = "父"
        End Select
      Case "子"
        Select Case sBz
          Case "亥" : s = "兄" : Case "子" : s = "兄"
          Case "丑" : s = "对" : Case "寅" : s = "对"
          Case "卯" : s = "仔" : Case "辰" : s = "仔"
          Case "巳" : s = "妻" : Case "午" : s = "妻"
          Case "未" : s = "官" : Case "申" : s = "官"
          Case "酉" : s = "父" : Case "戌" : s = "父"
        End Select
      Case "丑"
        Select Case sBz
          Case "亥" : s = "父" : Case "子" : s = "父"
          Case "丑" : s = "兄" : Case "寅" : s = "兄"
          Case "卯" : s = "对" : Case "辰" : s = "对"
          Case "巳" : s = "仔" : Case "午" : s = "仔"
          Case "未" : s = "妻" : Case "申" : s = "妻"
          Case "酉" : s = "官" : Case "戌" : s = "官"
        End Select
      Case "寅"
        Select Case sBz
          Case "亥" : s = "父" : Case "子" : s = "父"
          Case "丑" : s = "兄" : Case "寅" : s = "兄"
          Case "卯" : s = "对" : Case "辰" : s = "对"
          Case "巳" : s = "仔" : Case "午" : s = "仔"
          Case "未" : s = "妻" : Case "申" : s = "妻"
          Case "酉" : s = "官" : Case "戌" : s = "官"
        End Select
      Case "卯"
        Select Case sBz
          Case "亥" : s = "官" : Case "子" : s = "官"
          Case "丑" : s = "父" : Case "寅" : s = "父"
          Case "卯" : s = "兄" : Case "辰" : s = "兄"
          Case "巳" : s = "对" : Case "午" : s = "对"
          Case "未" : s = "仔" : Case "申" : s = "仔"
          Case "酉" : s = "妻" : Case "戌" : s = "妻"
        End Select
      Case "辰"
        Select Case sBz
          Case "亥" : s = "官" : Case "子" : s = "官"
          Case "丑" : s = "父" : Case "寅" : s = "父"
          Case "卯" : s = "兄" : Case "辰" : s = "兄"
          Case "巳" : s = "对" : Case "午" : s = "对"
          Case "未" : s = "仔" : Case "申" : s = "仔"
          Case "酉" : s = "妻" : Case "戌" : s = "妻"
        End Select
      Case "巳"
        Select Case sBz
          Case "亥" : s = "妻" : Case "子" : s = "妻"
          Case "丑" : s = "官" : Case "寅" : s = "官"
          Case "卯" : s = "父" : Case "辰" : s = "父"
          Case "巳" : s = "兄" : Case "午" : s = "兄"
          Case "未" : s = "对" : Case "申" : s = "对"
          Case "酉" : s = "仔" : Case "戌" : s = "仔"
        End Select
      Case "午"
        Select Case sBz
          Case "亥" : s = "妻" : Case "子" : s = "妻"
          Case "丑" : s = "官" : Case "寅" : s = "官"
          Case "卯" : s = "父" : Case "辰" : s = "父"
          Case "巳" : s = "兄" : Case "午" : s = "兄"
          Case "未" : s = "对" : Case "申" : s = "对"
          Case "酉" : s = "仔" : Case "戌" : s = "仔"
        End Select
      Case "未"
        Select Case sBz
          Case "亥" : s = "仔" : Case "子" : s = "仔"
          Case "丑" : s = "妻" : Case "寅" : s = "妻"
          Case "卯" : s = "官" : Case "辰" : s = "官"
          Case "巳" : s = "父" : Case "午" : s = "父"
          Case "未" : s = "兄" : Case "申" : s = "兄"
          Case "酉" : s = "对" : Case "戌" : s = "对"
        End Select
      Case "申"
        Select Case sBz
          Case "亥" : s = "仔" : Case "子" : s = "仔"
          Case "丑" : s = "妻" : Case "寅" : s = "妻"
          Case "卯" : s = "官" : Case "辰" : s = "官"
          Case "巳" : s = "父" : Case "午" : s = "父"
          Case "未" : s = "兄" : Case "申" : s = "兄"
          Case "酉" : s = "对" : Case "戌" : s = "对"
        End Select
      Case "酉"
        Select Case sBz
          Case "亥" : s = "对" : Case "子" : s = "对"
          Case "丑" : s = "仔" : Case "寅" : s = "仔"
          Case "卯" : s = "妻" : Case "辰" : s = "妻"
          Case "巳" : s = "官" : Case "午" : s = "官"
          Case "未" : s = "父" : Case "申" : s = "父"
          Case "酉" : s = "兄" : Case "戌" : s = "兄"
        End Select
      Case "戌"
        Select Case sBz
          Case "亥" : s = "对" : Case "子" : s = "对"
          Case "丑" : s = "仔" : Case "寅" : s = "仔"
          Case "卯" : s = "妻" : Case "辰" : s = "妻"
          Case "巳" : s = "官" : Case "午" : s = "官"
          Case "未" : s = "父" : Case "申" : s = "父"
          Case "酉" : s = "兄" : Case "戌" : s = "兄"
        End Select
    End Select
    Return s
  End Function

  Public Function uf八親_支vs支(sAz As String, sBz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    ' 辰.杰  已.火  午.火  未.鈥
    ' 卯.木               申.金
    ' 寅.木               酉.金
    ' 丑.沐  子.水  亥.水  戌.淦
    '------------------------------
    ' 父母兄妻　　木杰火鈥
    ' 仔媳官奴　　金淦水沐
    Dim s = ""

    Select Case sAz
      Case "亥" ' 水爻(流年)
        Select Case sBz
          Case "亥" : s = "-兄"
          Case "子" : s = "+兄"
          Case "丑" : s = "-妻" ' 沐
          Case "寅" : s = "+仔"
          Case "卯" : s = "-仔"
          Case "辰" : s = "+媳" ' 杰
          Case "巳" : s = "-官"
          Case "午" : s = "+官"
          Case "未" : s = "-奴" ' 鈥
          Case "申" : s = "+父"
          Case "酉" : s = "-父"
          Case "戌" : s = "+母" ' 淦
        End Select
      Case "子"
        Select Case sBz
          Case "亥" : s = "+兄"
          Case "子" : s = "-兄"
          Case "丑" : s = "+妻" ' 沐
          Case "寅" : s = "-仔"
          Case "卯" : s = "+仔"
          Case "辰" : s = "-媳" ' 杰
          Case "巳" : s = "+官"
          Case "午" : s = "-官"
          Case "未" : s = "+奴" ' 鈥
          Case "申" : s = "-父"
          Case "酉" : s = "+父"
          Case "戌" : s = "-母" ' 淦
        End Select
    ' 父母兄妻　　木杰火鈥
    ' 仔媳官奴　　金淦水沐
      Case "寅"
        Select Case sBz
          Case "亥" : s = "+父"
          Case "子" : s = "-父"
          Case "丑" : s = "+母" ' 沐
          Case "寅" : s = "-兄"
          Case "卯" : s = "+兄"
          Case "辰" : s = "-妻" ' 杰
          Case "巳" : s = "+仔"
          Case "午" : s = "-仔"
          Case "未" : s = "+媳" ' 鈥
          Case "申" : s = "-官"
          Case "酉" : s = "+官"
          Case "戌" : s = "-奴" ' 淦
        End Select
      Case "卯"
        Select Case sBz
          Case "亥" : s = "-父"
          Case "子" : s = "+父"
          Case "丑" : s = "-母" ' 沐
          Case "寅" : s = "+兄"
          Case "卯" : s = "-兄"
          Case "辰" : s = "+妻" ' 杰
          Case "巳" : s = "-仔"
          Case "午" : s = "+仔"
          Case "未" : s = "-媳" ' 鈥
          Case "申" : s = "+官"
          Case "酉" : s = "-官"
          Case "戌" : s = "+奴" ' 淦
        End Select

      Case "巳" ' 火
        Select Case sBz
          Case "亥" : s = "-官"
          Case "子" : s = "+官"
          Case "丑" : s = "-奴" ' 沐
          Case "寅" : s = "+父"
          Case "卯" : s = "-父"
          Case "辰" : s = "+母" ' 杰
          Case "巳" : s = "-兄"
          Case "午" : s = "+兄"
          Case "未" : s = "-妻" ' 鈥
          Case "申" : s = "+仔"
          Case "酉" : s = "-仔"
          Case "戌" : s = "+媳" ' 淦
        End Select
    ' 父母兄妻　　木杰火鈥
    ' 仔媳官奴　　金淦水沐
      Case "午" ' 火
        Select Case sBz
          Case "亥" : s = "+官"
          Case "子" : s = "-官"
          Case "丑" : s = "+奴" ' 沐
          Case "寅" : s = "-父"
          Case "卯" : s = "+父"
          Case "辰" : s = "-母" ' 杰
          Case "巳" : s = "+兄"
          Case "午" : s = "-兄"
          Case "未" : s = "+妻" ' 鈥
          Case "申" : s = "-仔"
          Case "酉" : s = "+仔"
          Case "戌" : s = "-媳" ' 淦
        End Select
      Case "申" ' 金
        Select Case sBz
          Case "亥" : s = "+仔"
          Case "子" : s = "-仔"
          Case "丑" : s = "+媳" ' 沐
          Case "寅" : s = "-官"
          Case "卯" : s = "+官"
          Case "辰" : s = "-奴" ' 杰
          Case "巳" : s = "+父"
          Case "午" : s = "-父"
          Case "未" : s = "+母" ' 鈥
          Case "申" : s = "-兄"
          Case "酉" : s = "+兄"
          Case "戌" : s = "-妻" ' 淦
        End Select
      Case "酉" ' 金
        Select Case sBz
          Case "亥" : s = "-仔"
          Case "子" : s = "+仔"
          Case "丑" : s = "-媳" ' 沐
          Case "寅" : s = "+官"
          Case "卯" : s = "-官"
          Case "辰" : s = "+奴" ' 杰
          Case "巳" : s = "-父"
          Case "午" : s = "+父"
          Case "未" : s = "-母" ' 鈥
          Case "申" : s = "+兄"
          Case "酉" : s = "-兄"
          Case "戌" : s = "+妻" ' 淦
        End Select
        '---------------------------------------------------
    ' 父母兄妻　　木杰火鈥
    ' 仔媳官奴　　金淦水沐
      Case "丑" ' 沐
        Select Case sBz
          Case "亥" : s = "-母"
          Case "子" : s = "+母"
          Case "丑" : s = "-兄" ' 沐
          Case "寅" : s = "+妻"
          Case "卯" : s = "-妻"
          Case "辰" : s = "+仔" ' 杰
          Case "巳" : s = "-媳"
          Case "午" : s = "+媳"
          Case "未" : s = "-官" ' 鈥
          Case "申" : s = "+奴"
          Case "酉" : s = "-奴"
          Case "戌" : s = "+父" ' 淦
        End Select
      Case "辰" ＇ 灭
        Select Case sBz
          Case "亥" : s = "+奴"
          Case "子" : s = "-奴"
          Case "丑" : s = "+父" ' 沐
          Case "寅" : s = "-母"
          Case "卯" : s = "+母"
          Case "辰" : s = "-兄" ' 杰
          Case "巳" : s = "+妻"
          Case "午" : s = "-妻"
          Case "未" : s = "+仔" ' 鈥
          Case "申" : s = "-媳"
          Case "酉" : s = "+媳"
          Case "戌" : s = "-官" ' 淦
        End Select
        '---------------------------------------------------
      Case "未" ＇ 鈥
        Select Case sBz
          Case "亥" : s = "-媳"
          Case "子" : s = "+媳"
          Case "丑" : s = "-官" ' 沐
          Case "寅" : s = "+奴"
          Case "卯" : s = "-奴"
          Case "辰" : s = "+父" ' 杰
          Case "巳" : s = "-母"
          Case "午" : s = "+母"
          Case "未" : s = "-兄" ' 鈥
          Case "申" : s = "+妻"
          Case "酉" : s = "-妻"
          Case "戌" : s = "+仔" ' 淦
        End Select
      Case "戌" ' 淦
        Select Case sBz
          Case "亥" : s = "+妻"
          Case "子" : s = "-妻"
          Case "丑" : s = "+仔" ' 沐
          Case "寅" : s = "-媳"
          Case "卯" : s = "+媳"
          Case "辰" : s = "-官" ' 杰
          Case "巳" : s = "+奴"
          Case "午" : s = "-奴"
          Case "未" : s = "+父" ' 鈥
          Case "申" : s = "-母"
          Case "酉" : s = "+母"
          Case "戌" : s = "-兄" ' 淦
        End Select
    End Select
    Return s
  End Function
  '================================================================
  Public Sub us_空亡(s干支 As String, ByRef s空亡 As String, ByRef s出空 As String)
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------------
    Dim s1 = "甲子,乙丑,丙寅,丁卯,戊辰,己巳,庚午,辛未,壬申,癸酉"
    Dim s2 = "甲戌,乙亥,丙子,丁丑,戊寅,己卯,庚辰,辛巳,壬午,癸未"
    Dim s3 = "甲申,乙酉,丙戌,丁亥,戊子,己丑,庚寅,辛卯,壬辰,癸巳"
    Dim s4 = "甲午,乙未,丙申,丁酉,戊戌,己亥,庚子,辛丑,壬寅,癸卯"
    Dim s5 = "甲辰,乙巳,丙午,丁未,戊申,己酉,庚戌,辛亥,壬子,癸丑"
    Dim s6 = "甲寅,乙卯,丙辰,丁巳,戊午,己未,庚申,辛酉,壬戌,癸亥"

    Select Case True
      Case InStr(s1, s干支) > 0 : s空亡 = "戌亥" : s出空 = "甲戌"
      Case InStr(s2, s干支) > 0 : s空亡 = "申酉" : s出空 = "甲申"
      Case InStr(s3, s干支) > 0 : s空亡 = "午未" : s出空 = "甲午"
      Case InStr(s4, s干支) > 0 : s空亡 = "辰巳" : s出空 = "甲辰"
      Case InStr(s5, s干支) > 0 : s空亡 = "寅卯" : s出空 = "甲寅"
      Case InStr(s6, s干支) > 0 : s空亡 = "子丑" : s出空 = "甲子"
      Case Else : s空亡 = "ＸＸ" : s出空 = "ＸＸ" 'X寅 =>無法找出"空亡"
    End Select
  End Sub
  '================================================================
  Public Sub us_旺相死休囚_時間支vs爻支(sHz As String, s爻z As String, ByRef s結果 As String, ByRef i分數 As Int32)
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------------
    '防有人打錯字
    sHz = sHz.Replace("已", "巳") : s爻z = s爻z.Replace("已", "巳")

    s結果 = ""
    Select Case sHz
      Case "寅", "卯"
        Select Case s爻z
          Case "寅", "卯" : s結果 = "旺"
          Case "巳", "午" : s結果 = "相"
          Case "申", "酉" : s結果 = "囚"
          Case "亥", "子" : s結果 = "休"
          Case "丑", "辰", "未", "戌" : s結果 = "死"
        End Select
      Case "巳", "午"
        Select Case s爻z
          Case "寅", "卯" : s結果 = "休"
          Case "巳", "午" : s結果 = "旺"
          Case "申", "酉" : s結果 = "死"
          Case "亥", "子" : s結果 = "囚"
          Case "丑", "辰", "未", "戌" : s結果 = "相"
        End Select
      Case "申", "酉"
        Select Case s爻z
          Case "寅", "卯" : s結果 = "死"
          Case "巳", "午" : s結果 = "囚"
          Case "申", "酉" : s結果 = "旺"
          Case "亥", "子" : s結果 = "相"
          Case "丑", "辰", "未", "戌" : s結果 = "休"
        End Select
      Case "亥", "子"
        Select Case s爻z
          Case "寅", "卯" : s結果 = "相"
          Case "巳", "午" : s結果 = "死"
          Case "申", "酉" : s結果 = "休"
          Case "亥", "子" : s結果 = "旺"
          Case "丑", "辰", "未", "戌" : s結果 = "囚"
        End Select
      Case "丑", "辰", "未", "戌"
        Select Case s爻z
          Case "寅", "卯" : s結果 = "囚"
          Case "巳", "午" : s結果 = "休"
          Case "申", "酉" : s結果 = "相"
          Case "亥", "子" : s結果 = "死"
          Case "丑", "辰", "未", "戌" : s結果 = "旺"
        End Select
    End Select

    Select Case s結果
      Case "旺" : i分數 = 4
      Case "相" : i分數 = 3
      Case "囚" : i分數 = 2
      Case "休" : i分數 = 1
      Case "死" : i分數 = 0
    End Select
  End Sub
  Public Sub us_旺相死休囚_時間干vs爻干(sHg As String, sXg As String, ByRef s結果 As String, ByRef i分數 As Int32)
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------------
    '防有人打錯字
    sHg = sHg.Replace("已", "己") : sXg = sXg.Replace("已", "己")

    s結果 = ""
    Select Case sHg
      Case "子", "丑"
        Select Case sXg
          Case "子", "丑" : s結果 = "旺" '爻干=我，我去  比旺  時間干
          Case "寅", "卯" : s結果 = "相"
          Case "午", "未" : s結果 = "囚"
          Case "申", "酉" : s結果 = "休"
          Case "辰", "巳" : s結果 = "死"
        End Select
      Case "寅", "卯"
        Select Case sXg
          Case "子", "丑" : s結果 = "休"
          Case "寅", "卯" : s結果 = "旺"
          Case "午", "未" : s結果 = "死"
          Case "申", "酉" : s結果 = "囚"
          Case "辰", "巳" : s結果 = "相"
        End Select
      Case "午", "未"
        Select Case sXg
          Case "子", "丑" : s結果 = "死"
          Case "寅", "卯" : s結果 = "囚"
          Case "午", "未" : s結果 = "旺"
          Case "申", "酉" : s結果 = "相"
          Case "辰", "巳" : s結果 = "休"
        End Select
      Case "申", "酉"
        Select Case sXg
          Case "子", "丑" : s結果 = "相"
          Case "寅", "卯" : s結果 = "死"
          Case "午", "未" : s結果 = "休"
          Case "申", "酉" : s結果 = "旺"
          Case "辰", "巳" : s結果 = "囚"
        End Select
      Case "辰", "巳"
        Select Case sXg
          Case "子", "丑" : s結果 = "囚"
          Case "寅", "卯" : s結果 = "休"
          Case "午", "未" : s結果 = "相"
          Case "申", "酉" : s結果 = "死"
          Case "辰", "巳" : s結果 = "旺"
        End Select
    End Select

    Select Case s結果
      Case "旺" : i分數 = 4
      Case "相" : i分數 = 3
      Case "囚" : i分數 = 2
      Case "休" : i分數 = 1
      Case "死" : i分數 = 0
    End Select
  End Sub
  Public Sub us_旺相死休囚沖_時間支vs爻支_六行(sHz As String, sXz As String, ByRef sR As String, ByRef i分數 As Int32)
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------------
    sHz = sHz.Replace("已", "巳") : sXz = sXz.Replace("已", "巳") '防有人打錯字

    sR = ""
    Select Case sHz
      Case "寅", "卯"
        Select Case sXz
          Case "寅", "卯" : sR = "旺"
          Case "辰", "巳" : sR = "相"
          Case "午", "未" : sR = "死"
          Case "申", "酉" : sR = "沖"
          Case "戌", "亥" : sR = "囚"
          Case "子", "丑" : sR = "休"
        End Select
      Case "辰", "巳"
        Select Case sXz
          Case "寅", "卯" : sR = "休"
          Case "辰", "巳" : sR = "旺"
          Case "午", "未" : sR = "相"
          Case "申", "酉" : sR = "死"
          Case "戌", "亥" : sR = "沖"
          Case "子", "丑" : sR = "囚"
        End Select
      Case "午", "未"
        Select Case sXz
          Case "寅", "卯" : sR = "囚"
          Case "辰", "巳" : sR = "休"
          Case "午", "未" : sR = "旺"
          Case "申", "酉" : sR = "相"
          Case "戌", "亥" : sR = "死"
          Case "子", "丑" : sR = "沖"
        End Select
      Case "申", "酉"
        Select Case sXz
          Case "寅", "卯" : sR = "沖"
          Case "辰", "巳" : sR = "囚"
          Case "午", "未" : sR = "休"
          Case "申", "酉" : sR = "旺"
          Case "戌", "亥" : sR = "相"
          Case "子", "丑" : sR = "死"
        End Select
      Case "戌", "亥"
        Select Case sXz
          Case "寅", "卯" : sR = "死"
          Case "辰", "巳" : sR = "沖"
          Case "午", "未" : sR = "囚"
          Case "申", "酉" : sR = "休"
          Case "戌", "亥" : sR = "旺"
          Case "子", "丑" : sR = "相"
        End Select
      Case "子", "丑"
        Select Case sXz
          Case "寅", "卯" : sR = "相"
          Case "辰", "巳" : sR = "死"
          Case "午", "未" : sR = "沖"
          Case "申", "酉" : sR = "囚"
          Case "戌", "亥" : sR = "休"
          Case "子", "丑" : sR = "旺"
        End Select
    End Select

    Select Case sR
      Case "旺" : i分數 = 5
      Case "相" : i分數 = 4
      Case "囚" : i分數 = 3
      Case "休" : i分數 = 2
      Case "死" : i分數 = 1
      Case "沖" : i分數 = 0
    End Select
  End Sub
  '================================================================
  Public Sub us_刑沖合害比(sZ1 As String, sZ2 As String, ByRef s結果 As String, ByRef i分數 As Int32)
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------------
    s結果 = "　" '無關
    If Not sZ1.Equals(sZ2) Then
      Select Case True
        Case "子,丑".Contains(sZ1) AndAlso "子,丑".Contains(sZ2),
             "寅,亥".Contains(sZ1) AndAlso "寅,亥".Contains(sZ2),
             "卯,戌".Contains(sZ1) AndAlso "卯,戌".Contains(sZ2),
             "辰,酉".Contains(sZ1) AndAlso "辰,酉".Contains(sZ2),
             "巳,申".Contains(sZ1) AndAlso "巳,申".Contains(sZ2),
             "午,未".Contains(sZ1) AndAlso "午,未".Contains(sZ2)
          s結果 = "　" '六合

        Case "子,辰,申".Contains(sZ1) AndAlso "子,辰,申".Contains(sZ2),
             "丑,巳,酉".Contains(sZ1) AndAlso "丑,巳,酉".Contains(sZ2),
             "寅,午,戌".Contains(sZ1) AndAlso "寅,午,戌".Contains(sZ2),
             "卯,未,亥".Contains(sZ1) AndAlso "卯,未,亥".Contains(sZ2)
          s結果 = "　" '三合，不看
        Case "子,午".Contains(sZ1) AndAlso "子,午".Contains(sZ2),
             "丑,未".Contains(sZ1) AndAlso "丑,未".Contains(sZ2),
             "寅,申".Contains(sZ1) AndAlso "寅,申".Contains(sZ2),
             "卯,酉".Contains(sZ1) AndAlso "卯,酉".Contains(sZ2),
             "辰,戌".Contains(sZ1) AndAlso "辰,戌".Contains(sZ2),
             "巳,亥".Contains(sZ1) AndAlso "巳,亥".Contains(sZ2)
          s結果 = "沖" '六沖
        Case "亥,子,丑".Contains(sZ1) AndAlso "亥,子,丑".Contains(sZ2),
             "寅,卯,辰".Contains(sZ1) AndAlso "寅,卯,辰".Contains(sZ2),
             "巳,午,未".Contains(sZ1) AndAlso "巳,午,未".Contains(sZ2),
             "申,酉,戌".Contains(sZ1) AndAlso "申,酉,戌".Contains(sZ2)
          s結果 = "　" '三會，不看
        Case "子,卯".Contains(sZ1) AndAlso "子,卯".Contains(sZ2) : s結果 = "刑"  '無禮之刑
        Case "丑,戌".Contains(sZ1) AndAlso "丑,戌".Contains(sZ2) : s結果 = "刑" '持勢之刑
        Case "丑,未".Contains(sZ1) AndAlso "丑,未".Contains(sZ2) : s結果 = "刑" '持勢之刑
        Case "寅,巳".Contains(sZ1) AndAlso "寅,巳".Contains(sZ2) : s結果 = "刑" '無恩之刑
        Case "巳,申".Contains(sZ1) AndAlso "巳,申".Contains(sZ2) : s結果 = "刑" '無恩之刑
        Case "未,戌".Contains(sZ1) AndAlso "未,戌".Contains(sZ2) : s結果 = "刑" '持勢之刑
        Case "子,未".Contains(sZ1) AndAlso "子,未".Contains(sZ2),
             "丑,午".Contains(sZ1) AndAlso "丑,午".Contains(sZ2),
             "寅,巳".Contains(sZ1) AndAlso "寅,巳".Contains(sZ2),
             "卯,辰".Contains(sZ1) AndAlso "卯,辰".Contains(sZ2),
             "申,亥".Contains(sZ1) AndAlso "申,亥".Contains(sZ2),
             "酉,戌".Contains(sZ1) AndAlso "酉,戌".Contains(sZ2)
          s結果 = "害" '六害
      End Select

    Else ' sZ1=sZ2
      Select Case sZ1
        Case "辰", "午", "酉", "亥" : s結果 = "刑" '自刑
        Case Else : s結果 = "比" '比旺
      End Select
    End If

    Select Case s結果
      Case "合" : i分數 = 5
      Case "比" : i分數 = 4
      Case "無" : i分數 = 3
      Case "刑" : i分數 = 2
      Case "害" : i分數 = 1
      Case "沖" : i分數 = 0
    End Select
  End Sub
  '================================================================
  Public Function uf_神煞_年支_短(s生Yz As String, s流Yz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Dim s神煞 = ""
    For i As Int32 = 0 To 11
      If __s年支_(i).Equals(s生Yz) Then
        If __s龍德_(i).Equals(s流Yz) Then s神煞 += "德"
        If __s金匱_(i).Equals(s流Yz) Then s神煞 += "匮"
        If __s紅鶯_(i).Equals(s流Yz) Then s神煞 += "莺"
        If __s天狗_(i).Equals(s流Yz) Then s神煞 += "狗"
        If __s勾紋_(i).Equals(s流Yz) Then s神煞 += "紋"
        'If __s岁破_(i).Equals(s流Yz) Then s神煞 += "岁"
        If __s破碎_(i).Equals(s流Yz) Then s神煞 += "碎"
        If __s大耗_(i).Equals(s流Yz) Then s神煞 += "耗" ' 大耗
        If __s五鬼_(i).Equals(s流Yz) Then s神煞 += "鬼" ' 五鬼
        If __s年桃_(i).Equals(s流Yz) Then s神煞 += "桃" ' 年桃
        If __s年刃_(i).Equals(s流Yz) Then s神煞 += "刃" ' 年刃
        'If __s年马_(i).Equals(s流Yz) Then s神煞 += "马"
        Exit For
      End If
    Next
    '去掉尾巴"、"
    If s神煞.Length >= 2 AndAlso Mid(s神煞, s神煞.Length, 1).Equals("、") Then
      s神煞 = Mid(s神煞, 1, s神煞.Length - 1)
    End If
    Return s神煞
  End Function

  Public Function uf_神煞_年支_長(s生Yz As String, s流Yz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Dim s神煞 = ""
    For i As Int32 = 0 To 11
      If __s年支_(i).Equals(s生Yz) Then
        If __s龍德_(i).Equals(s流Yz) Then s神煞 += "y龙德、"
        If __s金匱_(i).Equals(s流Yz) Then s神煞 += "y金匮、"
        If __s紅鶯_(i).Equals(s流Yz) Then s神煞 += "y紅莺、"
        If __s天狗_(i).Equals(s流Yz) Then s神煞 += "y天狗、"
        If __s勾紋_(i).Equals(s流Yz) Then s神煞 += "y勾紋、"
        If __s岁破_(i).Equals(s流Yz) Then s神煞 += "y岁破、"
        If __s破碎_(i).Equals(s流Yz) Then s神煞 += "y破碎、"
        If __s大耗_(i).Equals(s流Yz) Then s神煞 += "y大耗、"
        If __s五鬼_(i).Equals(s流Yz) Then s神煞 += "y五鬼、"
        If __s年桃_(i).Equals(s流Yz) Then s神煞 += "y年桃、"
        If __s年刃_(i).Equals(s流Yz) Then s神煞 += "y年刃、"
        If __s年马_(i).Equals(s流Yz) Then s神煞 += "y年马"
        Exit For
      End If
    Next
    '去掉尾巴"、"
    If s神煞.Length >= 2 AndAlso Mid(s神煞, s神煞.Length, 1).Equals("、") Then
      s神煞 = Mid(s神煞, 1, s神煞.Length - 1)
    End If
    Return s神煞
  End Function
  Public Function uf_神煞_月支(s生月支 As String, s流月支 As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Dim i As Int32, s神煞 = ""
    For i = 0 To 11
      If __s月支_(i).Equals(s生月支) Then
        If __s月刃_(i).Equals(s流月支) Then s神煞 += "m刃"
        If __s月马_(i).Equals(s流月支) Then s神煞 += "m馬"
        If __s月桃_(i).Equals(s流月支) Then s神煞 += "m挑"
        Exit For
      End If
    Next
    '去掉"、"
    If False AndAlso s神煞.Length >= 2 Then
      If Mid(s神煞, s神煞.Length, 1).Equals("、") Then
        s神煞 = Mid(s神煞, 1, s神煞.Length - 1)
      End If
    End If
    Return s神煞
  End Function
  Public Function uf_神煞_日干(sDg As String, sXz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Dim i As Int32, s神煞 = ""
    For i = 0 To 9
      If __s日干_(i).Equals(sDg) Then
        If __s天乙_(i).Equals(sXz) Then s神煞 += "dg乙"
        If __s文昌_(i).Equals(sXz) Then s神煞 += "dg昌"
        If __s金車_(i).Equals(sXz) Then s神煞 += "dg車"
        If __s紅艷_(i).Equals(sXz) Then s神煞 += "dg紅"
        If __s羊刃_(i).Equals(sXz) Then s神煞 += "dg羊"
        If __s禄神_(i).Equals(sXz) Then s神煞 += "dg禄"
        '李鵬多的
        'If __s国印_(i).Equals(sXz) Then s神煞 +=  "+国印"
        'If __s词馆_(i).Equals(sXz) Then s神煞 +=  "+词馆"
        'If __s太极_(i).Equals(sXz) Then s神煞 +=  "+太极"
        If __s福神_(i).Equals(sXz) Then s神煞 += "dg福"
        Exit For
      End If
    Next
    Return s神煞
  End Function
  Public Function uf_神煞_日支(sDZ As String, sXz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Dim i As Int32, s = ""
    For i = 0 To 11
      If __s日支_(i).Equals(sDZ) Then
        If __s将星_(i).Equals(sXz) Then s += "dz将" '将星
        If __s日马_(i).Equals(sXz) Then s += "dz马" '马星
        If __s日桃_(i).Equals(sXz) Then s += "dz桃" '桃花
        'If __s华盖_(i).Equals(sXz) Then s += "dz华" '华盖
        If __s日刃_(i).Equals(sXz) Then s += "dz刃"
        If __s劫煞_(i).Equals(sXz) Then s += "dz劫" '劫煞
        If __s亡神_(i).Equals(sXz) Then s += "dz亡" '亡神
        'If __s寡宿_(i).Equals(sXz) Then s += "dz寡" '寡宿
        'If __s孤辰_(i).Equals(sXz) Then s += "dz孤" '孤辰
        'If __s隔角_(i).Equals(sXz) Then s += "dz隔" '隔角
        If __s喪門_(i).Equals(sXz) Then s += "dz喪" '喪門
        '李鵬多的
        If __s灾星_(i).Equals(sXz) Then s += "dz灾" '灾星
        'If __s罗网_(i).Equals(sXz) Then s += "dz罗" '罗网
        Exit For
      End If
    Next
    Return s
  End Function
  '================================================================
  Public Function uf_日神煞_提示(sDg As String, sDz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Dim i As Int32, sR = ""
    '日干神煞
    For i = 0 To 9
      If __s日干_.Equals(sDg) Then
        sR += "日干－昌：" + __s文昌_(i) + "　禄：" + __s禄神_(i) + "　貴：" + __s天乙_(i) + vbCrLf
        Exit For
      End If
    Next

    '日支神煞

    For i = 0 To 9
      If __s日支_.Equals(sDz) Then
        sR += "日支－马：" + __s日马_(i) + "　桃：" + __s日桃_(i) + "　华：" + __s华盖_(i) + vbCrLf
        Exit For
      End If
    Next
    Return sR
  End Function
  '================================================================
  Public Function uf_五不遇爻(sDg As String, sXg As String) As String
    '五不遇時：时干去克日干 不宜用
    '五不遇爻：爻干去克日干 不宜用
    Dim s = ""

    Select Case sXg
      Case "子", "丑" : If sDg.Equals("辰") OrElse sDg.Equals("巳") Then s += "遇"
      Case "寅", "卯" : If sDg.Equals("午") OrElse sDg.Equals("未") Then s += "遇"
      Case "辰", "巳" : If sDg.Equals("申") OrElse sDg.Equals("酉") Then s += "遇"
      Case "午", "未" : If sDg.Equals("子") OrElse sDg.Equals("丑") Then s += "遇"
      Case "申", "酉" : If sDg.Equals("寅") OrElse sDg.Equals("卯") Then s += "遇"
    End Select
    Return s
  End Function
  '================================================================
  Public Function uf_12長生_干支_五行(sDg As String, sXz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Dim i As Int32, s = ""
    For i = 0 To 9
      If __s日干_(i).Equals(sDg) Then
        If __s长干_(i).Equals(sXz) Then : s = "长" '长生
        ElseIf __s沐干_(i).Equals(sXz) Then : s = "沐" '沐浴
        ElseIf __s冠干_(i).Equals(sXz) Then : s = "冠" '冠帶

        ElseIf __s臨干_(i).Equals(sXz) Then : s = "临" '临官=建禄
        ElseIf __s帝干_(i).Equals(sXz) Then : s = "旺" '帝旺
        ElseIf __s衰干_(i).Equals(sXz) Then : s = "衰"

        ElseIf __s病干_(i).Equals(sXz) Then : s = "病"
        ElseIf __s死干_(i).Equals(sXz) Then : s = "死"
        ElseIf __s墓干_(i).Equals(sXz) Then : s = "墓"

        ElseIf __s绝干_(i).Equals(sXz) Then : s = "绝"
        ElseIf __s胎干_(i).Equals(sXz) Then : s = "胎"
        ElseIf __s养干_(i).Equals(sXz) Then : s = "养"
        Else
          s = "　"
        End If

        Exit For
      End If
    Next

    Return s
  End Function
  Public Function uf_12長生_支支_五行(sXz As String, sDz As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Dim i As Int32, s = ""

    For i = 0 To 11
      If __s爻支_(i).Equals(sXz) Then '以A為準
        If __s长支_(i).Equals(sDz) Then : s = "长" '长生
        ElseIf __s沐支_(i).Equals(sDz) Then : s = "沐" '沐浴
        ElseIf __s冠支_(i).Equals(sDz) Then : s = "冠"
        ElseIf __s臨支_(i).Equals(sDz) Then : s = "临" '临官
        ElseIf __s帝支_(i).Equals(sDz) Then : s = "旺" '帝旺
        ElseIf __s衰支_(i).Equals(sDz) Then : s = "衰"
        ElseIf __s病支_(i).Equals(sDz) Then : s = "病"
        ElseIf __s死支_(i).Equals(sDz) Then : s = "死"
        ElseIf __s墓支_(i).Equals(sDz) Then : s = "墓"
        ElseIf __s绝支_(i).Equals(sDz) Then : s = "绝"
        ElseIf __s胎支_(i).Equals(sDz) Then : s = "胎"
        ElseIf __s养支_(i).Equals(sDz) Then : s = "养"
        End If
        Exit For
      End If
    Next
    Return s
  End Function
  Public Function uf_12長生_干支_六行(sGZ As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    Dim sG = Mid(sGZ, 1, 1), sZ = Mid(sGZ, 2, 1), s = ""
    Select Case sG
      Case "甲", "乙"
        Select Case sZ
          Case "亥" : s = "长" : Case "子" : s = "沐" : Case "丑" : s = "冠"
          Case "寅" : s = "临" : Case "卯" : s = "旺" : Case "辰" : s = "衰"
          Case "巳" : s = "病" : Case "午" : s = "死" : Case "未" : s = "墓"
          Case "申" : s = "绝" : Case "酉" : s = "胎" : Case "戌" : s = "养"
        End Select
      Case "丙", "丁", "戊"
        Select Case sZ
          Case "亥" : s = "绝" : Case "子" : s = "胎" : Case "丑" : s = "养"
          Case "寅" : s = "长" : Case "卯" : s = "沐" : Case "辰" : s = "冠"
          Case "巳" : s = "临" : Case "午" : s = "旺" : Case "未" : s = "衰"
          Case "申" : s = "病" : Case "酉" : s = "死" : Case "戌" : s = "墓"
        End Select
      Case "庚", "辛"
        Select Case sZ
          Case "亥" : s = "病" : Case "子" : s = "死" : Case "丑" : s = "墓"
          Case "寅" : s = "绝" : Case "卯" : s = "胎" : Case "辰" : s = "养"
          Case "巳" : s = "长" : Case "午" : s = "沐" : Case "未" : s = "冠"
          Case "申" : s = "临" : Case "酉" : s = "旺" : Case "戌" : s = "衰"
        End Select
      Case "壬", "癸", "己"
        Select Case sZ
          Case "亥" : s = "临" : Case "子" : s = "旺" : Case "丑" : s = "衰"
          Case "寅" : s = "病" : Case "卯" : s = "死" : Case "辰" : s = "墓"
          Case "巳" : s = "绝" : Case "午" : s = "胎" : Case "未" : s = "养"
          Case "申" : s = "长" : Case "酉" : s = "沐" : Case "戌" : s = "冠"
        End Select
    End Select
    Return s
  End Function

  Public Function uf_12長生_干支_五行_禄刃(sG As String, sZ As String) As String
    Dim s = "" ' 阴阳12长生
    Select Case sG
      Case "甲"
        Select Case sZ
          Case "亥" : s = "长" : Case "寅" : s = "禄" : Case "卯" : s = "刃"
          Case "午" : s = "死" : Case "未" : s = "墓" : Case "申" : s = "绝沖"
        End Select

      Case "丙", "戊"
        Select Case sZ
          Case "寅" : s = "长" : Case "巳" : s = "禄" : Case "午" : s = "刃"
          Case "酉" : s = "死" : Case "戌" : s = "墓" : Case "亥" : s = "绝沖"
        End Select

      Case "庚"
        Select Case sZ
          Case "子" : s = "死" : Case "丑" : s = "墓" : Case "寅" : s = "绝沖"
          Case "巳" : s = "长" : Case "申" : s = "禄" : Case "酉" : s = "刃"
        End Select

      Case "壬"
        Select Case sZ
          Case "亥" : s = "禄" : Case "子" : s = "刃" : Case "申" : s = "长"
          Case "卯" : s = "死" : Case "辰" : s = "墓" : Case "巳" : s = "绝沖"
        End Select
     '---------------------------------------------------
      Case "乙" ' 阴长生
        Select Case sZ
          Case "午" : s = "长" : Case "卯" : s = "禄" : Case "辰" : s = "刃"
          Case "亥" : s = "死" : Case "戌" : s = "墓" : Case "酉" : s = "绝沖"
        End Select

      Case "丁", "己"
        Select Case sZ
          Case "酉" : s = "长" : Case "午" : s = "禄" : Case "未" : s = "刃"
          Case "寅" : s = "死" : Case "丑" : s = "墓" : Case "子" : s = "绝沖"
        End Select

      Case "辛"
        Select Case sZ
          Case "子" : s = "长" : Case "酉" : s = "禄" : Case "戌" : s = "刃"
          Case "巳" : s = "死" : Case "辰" : s = "墓" : Case "卯" : s = "绝沖"
        End Select

      Case "癸"
        Select Case sZ
          Case "卯" : s = "长" : Case "子" : s = "禄" : Case "丑" : s = "刃"
          Case "申" : s = "死" : Case "未" : s = "墓" : Case "午" : s = "绝沖"
        End Select
    End Select
    If s.Length > 0 Then Return "(" + s + ")" Else Return ""
  End Function
  Public Function uf_12長生_干支_六行_禄刃(sG As String, sZ As String) As String
    Dim s = ""
    Select Case sG
      Case "甲"
        Select Case sZ
          Case "亥" : s = "长" : Case "寅" : s = "禄" : Case "卯" : s = "刃"
          Case "午" : s = "死" : Case "未" : s = "墓" : Case "申" : s = "绝沖"
        End Select
      Case "乙"
        Select Case sZ
          Case "亥" : s = "长" : Case "寅" : s = "刃" : Case "卯" : s = "禄"
          Case "午" : s = "死" : Case "未" : s = "墓" : Case "申" : s = "绝" : Case "酉" : s = "沖"
        End Select

      Case "丙"
        Select Case sZ
          Case "寅" : s = "长" : Case "巳" : s = "禄" : Case "午" : s = "刃"
          Case "酉" : s = "死" : Case "戌" : s = "墓" : Case "亥" : s = "绝沖"
        End Select
      Case "丁"
        Select Case sZ
          Case "寅" : s = "长" : Case "巳" : s = "刃" : Case "午" : s = "禄"
          Case "酉" : s = "死" : Case "戌" : s = "墓" : Case "亥" : s = "绝" : Case "子" : s = "沖"
        End Select
      Case "戊" '=丙未
        Select Case sZ
          Case "寅" : s = "长" : Case "巳" : s = "禄" : Case "未" : s = "刃" ' 跟丙不同
          Case "酉" : s = "死" : Case "戌" : s = "墓" : Case "亥" : s = "绝" : Case "丑" : s = "沖"
        End Select

      Case "庚"
        Select Case sZ
          Case "子" : s = "死" : Case "丑" : s = "墓" : Case "寅" : s = "绝沖"
          Case "巳" : s = "长" : Case "申" : s = "禄" : Case "酉" : s = "刃"
        End Select
      Case "辛"
        Select Case sZ
          Case "子" : s = "死" : Case "丑" : s = "墓" : Case "寅" : s = "绝" : Case "卯" : s = "沖"
          Case "巳" : s = "长" : Case "申" : s = "刃" : Case "酉" : s = "禄"
        End Select

      Case "壬"
        Select Case sZ
          Case "申" : s = "长" : Case "亥" : s = "禄" : Case "子" : s = "刃"
          Case "卯" : s = "死" : Case "辰" : s = "墓" : Case "巳" : s = "绝沖"
        End Select
      Case "癸"
        Select Case sZ
          Case "申" : s = "长" : Case "亥" : s = "刃" : Case "子" : s = "禄"
          Case "卯" : s = "死" : Case "辰" : s = "墓" : Case "巳" : s = "绝" : Case "午" : s = "沖"
        End Select
      Case "己" '=壬丑
        Select Case sZ
          Case "申" : s = "长" : Case "亥" : s = "禄" : Case "丑" : s = "刃"
          Case "卯" : s = "死" : Case "辰" : s = "墓" : Case "巳" : s = "绝" : Case "未" : s = "沖"
        End Select
    End Select
    If s.Length > 0 Then Return "(" + s + ")" Else Return ""
  End Function
  Public Function uf_12長生_支支_六行_禄刃(sZA As String, sZB As String) As String
    Dim s = ""
    Select Case sZA
      Case "寅" ' 甲
        Select Case sZB
          Case "亥" : s = "长" : Case "寅" : s = "禄" : Case "卯" : s = "刃"
          Case "午" : s = "死" : Case "未" : s = "墓" : Case "申" : s = "绝沖"
        End Select
      Case "卯" ' 乙
        Select Case sZB
          Case "亥" : s = "长" : Case "寅" : s = "刃" : Case "卯" : s = "禄"
          Case "午" : s = "死" : Case "未" : s = "墓" : Case "申" : s = "绝" : Case "酉" : s = "沖"
        End Select

      Case "巳" ' 丙
        Select Case sZB
          Case "寅" : s = "长" : Case "巳" : s = "禄" : Case "午" : s = "刃"
          Case "酉" : s = "死" : Case "戌" : s = "墓" : Case "亥" : s = "绝沖"
        End Select
      Case "午" ' 丁
        Select Case sZB
          Case "寅" : s = "长" : Case "巳" : s = "刃" : Case "午" : s = "禄"
          Case "酉" : s = "死" : Case "戌" : s = "墓" : Case "亥" : s = "绝" : Case "子" : s = "沖"
        End Select
      Case "未" '=丙戊未
        Select Case sZB
          Case "寅" : s = "长" : Case "巳" : s = "禄" : Case "未" : s = "刃" ' 跟丙不同
          Case "酉" : s = "死" : Case "戌" : s = "墓" : Case "亥" : s = "绝" : Case "丑" : s = "沖"
        End Select

      Case "申"
        Select Case sZB
          Case "子" : s = "死" : Case "丑" : s = "墓" : Case "寅" : s = "绝沖"
          Case "巳" : s = "长" : Case "申" : s = "禄" : Case "酉" : s = "刃"
        End Select
      Case "酉"
        Select Case sZB
          Case "子" : s = "死" : Case "丑" : s = "墓" : Case "寅" : s = "绝" : Case "卯" : s = "沖"
          Case "巳" : s = "长" : Case "申" : s = "刃" : Case "酉" : s = "禄"
        End Select

      Case "亥"
        Select Case sZB
          Case "申" : s = "长" : Case "亥" : s = "禄" : Case "子" : s = "刃"
          Case "卯" : s = "死" : Case "辰" : s = "墓" : Case "巳" : s = "绝沖"
        End Select
      Case "子"
        Select Case sZB
          Case "申" : s = "长" : Case "亥" : s = "刃" : Case "子" : s = "禄"
          Case "卯" : s = "死" : Case "辰" : s = "墓" : Case "巳" : s = "绝" : Case "午" : s = "沖"
        End Select
      Case "丑" '=壬己丑
        Select Case sZB
          Case "申" : s = "长" : Case "亥" : s = "禄" : Case "丑" : s = "刃"
          Case "卯" : s = "死" : Case "辰" : s = "墓" : Case "巳" : s = "绝" : Case "未" : s = "沖"
        End Select
    End Select
    If s.Length > 0 Then Return "(" + s + ")" Else Return ""
  End Function
  '================================================================
  Private Function r(s As String) As String
    ' 試著 代替 Retrun，如：Return "比" 改寫=> r("比")， 但其實沒用
    Return s

#If False Then
❌ 為何 不能 改寫 Return 本身？
   VB 規則：
   * Return 是保留字（reserved keyword）
   * 不能被 shadow、alias、override
   * 不能成為函式或變數名（例如不能定義 Sub Return()）
   * 不能被使用者取代成別名（VB 無巨集）

也就是說：VB
  Dim r = Return  ' ❌ 語法錯誤
  r "比"          ' ❌ 絕對不行
#End If
  End Function

  Public Function uf_印劫伤财官_卦支vs爻支_五行_改寫_dic法_徐不用(sAz As String, sBz As String) As String
    'copilot 改寫，缺點：
    '  * 沒比較容易看懂
    '  * 每次都要建 dic，不管元素用得到用不到，全部都要建，浪費記憶體
    Dim M As New Dictionary(Of String, String) From {
    {"亥亥", "比"}, {"亥子", "劫"}, ' 亥
    {"亥寅", "伤"}, {"亥卯", "食"},
    {"亥巳", "贝"}, {"亥午", "财"},
    {"亥申", "印"}, {"亥酉", "枭"},
    {"亥丑", "杀"}, {"亥未", "杀"},
    {"亥辰", "官"}, {"亥戌", "官"},
    {"a01", "--"}, {"b01", "--"}, ' 不能分隔為空白行，故人為弄出來的區隔線
    {"子亥", "劫"}, {"子子", "比"}, ' 子
    {"子寅", "伤"}, {"子卯", "食"},
    {"子巳", "财"}, {"子午", "贝"},
    {"子申", "枭"}, {"子酉", "印"},
    {"子丑", "官"}, {"子未", "官"},
    {"子辰", "杀"}, {"子戌", "杀"},
    {"a02", "--"}, {"b02", "--"}, ' 人為弄出來的區隔線
    {"寅亥", "印"}, {"寅子", "枭"}, ' 寅
    {"寅寅", "比"}, {"寅卯", "劫"},
    {"寅巳", "伤"}, {"寅午", "食"},
    {"寅申", "杀"}, {"寅酉", "官"},
    {"寅丑", "财"}, {"寅未", "财"},
    {"寅辰", "贝"}, {"寅戌", "贝"},
    {"卯亥", "枭"}, {"卯子", "印"}, ' 卯
    {"卯寅", "劫"}, {"卯卯", "比"},
    {"卯巳", "食"}, {"卯午", "伤"}, {"卯申", "官"}, {"卯酉", "杀"},
    {"卯丑", "贝"}, {"卯未", "贝"}, {"卯辰", "财"}, {"卯戌", "财"},
    {"巳亥", "杀"}, {"巳子", "官"}, {"巳寅", "印"}, {"巳卯", "枭"},
    {"巳巳", "比"}, {"巳午", "劫"}, {"巳申", "财"}, {"巳酉", "贝"},
    {"巳丑", "食"}, {"巳未", "食"}, {"巳辰", "伤"}, {"巳戌", "伤"},
    {"午亥", "官"}, {"午子", "杀"}, {"午寅", "枭"}, {"午卯", "印"},
    {"午巳", "劫"}, {"午午", "比"}, {"午申", "贝"}, {"午酉", "财"},
    {"午丑", "伤"}, {"午未", "伤"}, {"午辰", "食"}, {"午戌", "食"},
    {"申亥", "伤"}, {"申子", "食"}, {"申寅", "贝"}, {"申卯", "财"},
    {"申巳", "官"}, {"申午", "杀"}, {"申申", "比"}, {"申酉", "劫"},
    {"申丑", "印"}, {"申未", "印"}, {"申辰", "枭"}, {"申戌", "枭"},
    {"酉亥", "食"}, {"酉子", "伤"}, {"酉寅", "财"}, {"酉卯", "贝"}, ' 酉 行
    {"酉巳", "杀"}, {"酉午", "官"}, {"酉申", "劫"}, {"酉酉", "比"},
    {"酉丑", "枭"}, {"酉未", "枭"}, {"酉辰", "印"}, {"酉戌", "印"},
    {"丑亥", "贝"}, {"丑子", "财"}, {"丑寅", "官"}, {"丑卯", "杀"}, ' 丑 行
    {"丑巳", "枭"}, {"丑午", "印"}, {"丑申", "伤"}, {"丑酉", "食"},
    {"丑丑", "比"}, {"丑未", "比"}, {"丑辰", "劫"}, {"丑戌", "劫"},
    {"未亥", "贝"}, {"未子", "财"}, {"未寅", "官"}, {"未卯", "杀"},' 未 行（同丑）
    {"未巳", "枭"}, {"未午", "印"}, {"未申", "伤"}, {"未酉", "食"},
    {"未丑", "比"}, {"未未", "比"}, {"未辰", "劫"}, {"未戌", "劫"},
    {"辰亥", "财"}, {"辰子", "贝"}, {"辰寅", "杀"}, {"辰卯", "官"},' 辰 行
    {"辰巳", "印"}, {"辰午", "枭"}, {"辰申", "食"}, {"辰酉", "伤"},
    {"辰丑", "劫"}, {"辰未", "劫"}, {"辰辰", "比"}, {"辰戌", "比"},
    {"戌亥", "财"}, {"戌子", "贝"}, {"戌寅", "杀"}, {"戌卯", "官"},' 戌 行（同辰）
    {"戌巳", "印"}, {"戌午", "枭"}, {"戌申", "食"}, {"戌酉", "伤"},
    {"戌丑", "劫"}, {"戌未", "劫"}, {"戌辰", "比"}, {"戌戌", "比"}
  }

    If True Then
      Dim s = "" : Return If(M.TryGetValue(sAz & sBz, s), s, "") ' 更短版
    Else
      Dim s = "" : If M.TryGetValue(sAz + sBz, s) Then Return s Else Return ""
    End If
  End Function
  Public Function uf_印劫伤财官_卦支vs爻支_五行_改寫_單層SelectCase(sAz As String, sBz As String) As String
    ' copilot 改寫： 省掉了很多 Select Case sBz ...End Select
    'A=卦宮支＝世爻支=我, B=爻支
    Dim s = ""
    Select Case sAz + sBz
      Case "亥亥" : Return "比" : Case "亥子" : Return "劫"
      Case "亥寅" : Return "伤" : Case "亥卯" : Return "食"
      Case "亥巳" : s = "贝" : Case "亥午" : s = "财"
      Case "亥申" : s = "印" : Case "亥酉" : s = "枭"
      Case "亥丑" : s = "杀" : Case "亥未" : s = "杀"
      Case "亥辰" : s = "官" : Case "亥戌" : s = "官"

      Case "子亥" : s = "劫" : Case "子子" : s = "比"  '我
      Case "子寅" : s = "伤" : Case "子卯" : s = "食"
      Case "子巳" : s = "财" : Case "子午" : s = "贝"
      Case "子申" : s = "枭" : Case "子酉" : s = "印"
      Case "子丑" : s = "官" : Case "子未" : s = "官"
      Case "子辰" : s = "杀" : Case "子戌" : s = "杀"
    End Select
    Return s
  End Function
  Public Function uf_印劫伤财官_卦支vs爻支_五行_改寫_字串法(sAz As String, sBz As String) As String
#If False Then ' VB的大量註解法，類似C# /*...*/
    copilot 方案 3：使用“字串表” + Mid 比對（最密、極短、不用 Dictionary），這是我認為 最適合你風格 的寫法
  * 不用：Dictionary、Select Case、ElseIf
  * 全部資料可以「擠在一個大字串」裡
  * VS 完全不會自動格式化
  * 超快 + 超密 + VS 不會亂排
  * 可設定成 常數 (Const)
  非常適合你目前的八字 / 易經 對照表資料量。
  
  Dim s =""  <--被包起來後，不會顯示顏色
  Console.WriteLine("測試")
#End If

    If False Then '唯一能讓 程或碼顯示顏色+不被執行 的方法
      Dim s1 = "" '< --被包起來後， 不會顯示顏色
      Console.WriteLine("測試")
    End If
    '------------------------------------------------------
    Const T =
"亥亥比 亥子劫 亥寅伤 亥卯食 亥巳贝 亥午财 亥申印 亥酉枭 亥丑杀 亥未杀 亥辰官 亥戌官 " &
"子亥劫 子子比 子寅伤 子卯食 子巳财 子午贝 子申枭 子酉印 子丑官 子未官 子辰杀 子戌杀 " &
"寅亥印 寅子枭 寅寅比 寅卯劫 寅巳伤 寅午食 寅申杀 寅酉官 寅丑财 寅未财 寅辰贝 寅戌贝 " &
"卯亥枭 卯子印 卯寅劫 卯卯比 卯巳食 卯午伤 卯申官 卯酉杀 卯丑贝 卯未贝 卯辰财 卯戌财 " &
"巳亥杀 巳子官 巳寅印 巳卯枭 巳巳比 巳午劫 巳申财 巳酉贝 巳丑食 巳未食 巳辰伤 巳戌伤 " &
"午亥官 午子杀 午寅枭 午卯印 午巳劫 午午比 午申贝 午酉财 午丑伤 午未伤 午辰食 午戌食 " &
"申亥伤 申子食 申寅贝 申卯财 申巳官 申午杀 申申比 申酉劫 申丑印 申未印 申辰枭 申戌枭 " &
"酉亥食 酉子伤 酉寅财 酉卯贝 酉巳杀 酉午官 酉申劫 酉酉比 酉丑枭 酉未枭 酉辰印 酉戌印 " &
"丑亥贝 丑子财 丑寅官 丑卯杀 丑巳枭 丑午印 丑申伤 丑酉食 丑丑比 丑未比 丑辰劫 丑戌劫 " &
"未亥贝 未子财 未寅官 未卯杀 未巳枭 未午印 未申伤 未酉食 未丑比 未未比 未辰劫 未戌劫 " &
"辰亥财 辰子贝 辰寅杀 辰卯官 辰巳印 辰午枭 辰申食 辰酉伤 辰丑劫 辰未劫 辰辰比 辰戌比 " &
"戌亥财 戌子贝 戌寅杀 戌卯官 戌巳印 戌午枭 戌申食 戌酉伤 戌丑劫 戌未劫 戌辰比 戌戌比 "
    Dim p = InStr(T, sAz & sBz) : Return If(p < 0, "", Mid(T, p + 2, 1))

    ' 改寫：用 常數 => 好處：點1個變數，其它相同者 高亮標示
    Const T2 =
亥亥 + 比 + 亥子 + 劫 + 亥寅 + 伤 + 亥卯 + 食 + 亥巳 + 贝 + 亥午 + 财 + 亥申 + 印 + 亥酉 + 枭 + 亥丑 + 杀 + 亥未 + 杀 + 亥辰 + 官 + 亥戌 + 官 +
子亥 + 劫 + 子子 + 比 + 子寅 + 伤 + 子卯 + 食 + 子巳 + 财 + 子午 + 贝 + 子申 + 枭 + 子酉 + 印 + 子丑 + 官 + 子未 + 官 + 子辰 + 杀 + 子戌 + 杀 +
寅亥 + 印 + 寅子 + 枭 + 寅寅 + 比 + 寅卯 + 劫 + 寅巳 + 伤 + 寅午 + 食 + 寅申 + 杀 + 寅酉 + 官 + 寅丑 + 财 + 寅未 + 财 + 寅辰 + 贝 + 寅戌 + 贝 +
卯亥 + 枭 + 卯子 + 印 + 卯寅 + 劫 + 卯卯 + 比 + 卯巳 + 食 + 卯午 + 伤 + 卯申 + 官 + 卯酉 + 杀 + 卯丑 + 贝 + 卯未 + 贝 + 卯辰 + 财 + 卯戌 + 财 +
巳亥 + 杀 + 巳子 + 官 + 巳寅 + 印 + 巳卯 + 枭 + 巳巳 + 比 + 巳午 + 劫 + 巳申 + 财 + 巳酉 + 贝 + 巳丑 + 食 + 巳未 + 食 + 巳辰 + 伤 + 巳戌 + 伤 +
午亥 + 官 + 午子 + 杀 + 午寅 + 枭 + 午卯 + 印 + 午巳 + 劫 + 午午 + 比 + 午申 + 贝 + 午酉 + 财 + 午丑 + 伤 + 午未 + 伤 + 午辰 + 食 + 午戌 + 食 +
申亥 + 伤 + 申子 + 食 + 申寅 + 贝 + 申卯 + 财 + 申巳 + 官 + 申午 + 杀 + 申申 + 比 + 申酉 + 劫 + 申丑 + 印 + 申未 + 印 + 申辰 + 枭 + 申戌 + 枭 +
酉亥 + 食 + 酉子 + 伤 + 酉寅 + 财 + 酉卯 + 贝 + 酉巳 + 杀 + 酉午 + 官 + 酉申 + 劫 + 酉酉 + 比 + 酉丑 + 枭 + 酉未 + 枭 + 酉辰 + 印 + 酉戌 + 印 +
丑亥 + 贝 + 丑子 + 财 + 丑寅 + 官 + 丑卯 + 杀 + 丑巳 + 枭 + 丑午 + 印 + 丑申 + 伤 + 丑酉 + 食 + 丑丑 + 比 + 丑未 + 比 + 丑辰 + 劫 + 丑戌 + 劫 +
未亥 + 贝 + 未子 + 财 + 未寅 + 官 + 未卯 + 杀 + 未巳 + 枭 + 未午 + 印 + 未申 + 伤 + 未酉 + 食 + 未丑 + 比 + 未未 + 比 + 未辰 + 劫 + 未戌 + 劫 +
辰亥 + 财 + 辰子 + 贝 + 辰寅 + 杀 + 辰卯 + 官 + 辰巳 + 印 + 辰午 + 枭 + 辰申 + 食 + 辰酉 + 伤 + 辰丑 + 劫 + 辰未 + 劫 + 辰辰 + 比 + 辰戌 + 比 +
戌亥 + 财 + 戌子 + 贝 + 戌寅 + 杀 + 戌卯 + 官 + 戌巳 + 印 + 戌午 + 枭 + 戌申 + 食 + 戌酉 + 伤 + 戌丑 + 劫 + 戌未 + 劫 + 戌辰 + 比 + 戌戌 + 比
    p = InStr(T2, sAz & sBz) : Return If(p < 0, "", Mid(T2, p + 2, 1))
  End Function

  Public Function uf_印劫伤财官_卦支vs爻支_五行(sAz As String, sBz As String) As String
    'A=卦宮支＝世爻支=我, B=爻支
    Dim s = ""
    Select Case sAz
      Case "亥"
        Select Case sBz
          Case "亥" : s = "比" : Case "子" : s = "劫"  '我
          Case "寅" : s = "伤" : Case "卯" : s = "食"
          Case "巳" : s = "贝" : Case "午" : s = "财"
          Case "申" : s = "印" : Case "酉" : s = "枭"
          Case "丑" : s = "杀" : Case "未" : s = "杀"
          Case "辰" : s = "官" : Case "戌" : s = "官"
        End Select

      Case "子" '卦宮
        Select Case sBz
          Case "亥" : s = "劫" : Case "子" : s = "比"  '我
          Case "寅" : s = "伤" : Case "卯" : s = "食"
          Case "巳" : s = "财" : Case "午" : s = "贝"
          Case "申" : s = "枭" : Case "酉" : s = "印" ' 可分隔成空白行，方便閱讀

          Case "丑" : s = "官" : Case "未" : s = "官"
          Case "辰" : s = "杀" : Case "戌" : s = "杀"
        End Select

      Case "寅"
        Select Case sBz
          Case "亥" : s = "印" : Case "子" : s = "枭"
          Case "寅" : s = "比" : Case "卯" : s = "劫"
          Case "巳" : s = "伤" : Case "午" : s = "食"
          Case "申" : s = "杀" : Case "酉" : s = "官"
          Case "丑" : s = "财" : Case "未" : s = "财"
          Case "辰" : s = "贝" : Case "戌" : s = "贝"
        End Select

      Case "卯"
        Select Case sBz
          Case "亥" : s = "枭" : Case "子" : s = "印"
          Case "寅" : s = "劫" : Case "卯" : s = "比"
          Case "巳" : s = "食" : Case "午" : s = "伤"
          Case "申" : s = "官" : Case "酉" : s = "杀"
          Case "丑" : s = "贝" : Case "未" : s = "贝"
          Case "辰" : s = "财" : Case "戌" : s = "财"
        End Select

      Case "巳"
        Select Case sBz
          Case "亥" : s = "杀" : Case "子" : s = "官"  '我
          Case "寅" : s = "印" : Case "卯" : s = "枭"
          Case "巳" : s = "比" : Case "午" : s = "劫"
          Case "申" : s = "财" : Case "酉" : s = "贝"
          Case "丑" : s = "食" : Case "未" : s = "食"
          Case "辰" : s = "伤" : Case "戌" : s = "伤"
        End Select
      Case "午"
        Select Case sBz
          Case "亥" : s = "官" : Case "子" : s = "杀"  '我
          Case "寅" : s = "枭" : Case "卯" : s = "印"
          Case "巳" : s = "劫" : Case "午" : s = "比"
          Case "申" : s = "贝" : Case "酉" : s = "财"
          Case "丑" : s = "伤" : Case "未" : s = "伤"
          Case "辰" : s = "食" : Case "戌" : s = "食"
        End Select

      Case "申"
        Select Case sBz
          Case "亥" : s = "伤" : Case "子" : s = "食"  '我
          Case "寅" : s = "贝" : Case "卯" : s = "财"
          Case "巳" : s = "官" : Case "午" : s = "杀"
          Case "申" : s = "比" : Case "酉" : s = "劫"
          Case "丑" : s = "印" : Case "未" : s = "印"
          Case "辰" : s = "枭" : Case "戌" : s = "枭"
        End Select
      Case "酉"
        Select Case sBz
          Case "亥" : s = "食" : Case "子" : s = "伤"  '我
          Case "寅" : s = "财" : Case "卯" : s = "贝"
          Case "巳" : s = "杀" : Case "午" : s = "官"
          Case "申" : s = "劫" : Case "酉" : s = "比"
          Case "丑" : s = "枭" : Case "未" : s = "枭"
          Case "辰" : s = "印" : Case "戌" : s = "印"
        End Select

      Case "丑"
        Select Case sBz
          Case "亥" : s = "贝" : Case "子" : s = "财"  '我
          Case "寅" : s = "官" : Case "卯" : s = "杀"
          Case "巳" : s = "枭" : Case "午" : s = "印"
          Case "申" : s = "伤"
          Case "酉" : s = "食"

          Case "丑" : s = "比"
          Case "未" : s = "比"
          Case "辰" : s = "劫"
          Case "戌" : s = "劫"
        End Select

      Case "未"
        Select Case sBz
          Case "亥" : s = "贝" : Case "子" : s = "财"
          Case "寅" : s = "官" : Case "卯" : s = "杀"
          Case "巳" : s = "枭" : Case "午" : s = "印"
          Case "申" : s = "伤" : Case "酉" : s = "食"

          Case "丑" : s = "比" : Case "未" : s = "比"
          Case "辰" : s = "劫" : Case "戌" : s = "劫"
        End Select

      Case "辰"
        Select Case sBz
          Case "亥" : s = "财" : Case "子" : s = "贝"  '我
          Case "寅" : s = "杀" : Case "卯" : s = "官"
          Case "巳" : s = "印" : Case "午" : s = "枭"
          Case "申" : s = "食" : Case "酉" : s = "伤"

          Case "丑" : s = "劫" : Case "未" : s = "劫"
          Case "辰" : s = "比" : Case "戌" : s = "比"
        End Select

      Case "戌"
        Select Case sBz
          Case "亥" : s = "财" : Case "子" : s = "贝"  '我
          Case "寅" : s = "杀" : Case "卯" : s = "官"
          Case "巳" : s = "印" : Case "午" : s = "枭"
          Case "申" : s = "食" : Case "酉" : s = "伤"

          Case "丑" : s = "劫" : Case "未" : s = "劫"
          Case "辰" : s = "比" : Case "戌" : s = "比"
        End Select
    End Select
    Return s
  End Function
  Public Function uf_印劫伤财官_卦支vs爻支_六行(sZa As String, sZb As String) As String
    On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Function
0:  '---------------------------------------------------
    '亥子水，丑寅元，卯辰木，巳午火，未申土，酉戌金
    'A=卦宮支＝世支=我, B=爻支
    ' 早期：父兄仔妻财官
    Dim s = ""

    Select Case sZa
      Case "丑"
        Select Case sZb
          Case "丑" : s = "比" : Case "寅" : s = "劫"
          Case "卯" : s = "对" : Case "辰" : s = "吸"
          Case "巳" : s = "食" : Case "午" : s = "伤"
          Case "未" : s = "贝" : Case "申" : s = "财"
          Case "酉" : s = "杀" : Case "戌" : s = "官"
          Case "亥" : s = "枭" : Case "子" : s = "印"
        End Select
      Case "寅" '木
        Select Case sZb
          Case "丑" : s = "劫" : Case "寅" : s = "比"
          Case "卯" : s = "吸" : Case "辰" : s = "对"
          Case "巳" : s = "伤" : Case "午" : s = "食"
          Case "未" : s = "财" : Case "申" : s = "贝"
          Case "酉" : s = "官" : Case "戌" : s = "杀"
          Case "亥" : s = "印" : Case "子" : s = "枭"
        End Select
      Case "卯" '木
        Select Case sZb
          Case "丑" : s = "枭" : Case "寅" : s = "印"
          Case "卯" : s = "比" : Case "辰" : s = "劫"
          Case "巳" : s = "对" : Case "午" : s = "吸"
          Case "未" : s = "食" : Case "申" : s = "伤"
          Case "酉" : s = "贝" : Case "戌" : s = "财"
          Case "亥" : s = "杀" : Case "子" : s = "官"
        End Select
      Case "辰" '火
        Select Case sZb
          Case "丑" : s = "印" : Case "寅" : s = "枭"
          Case "卯" : s = "劫" : Case "辰" : s = "比"
          Case "巳" : s = "吸" : Case "午" : s = "对"
          Case "未" : s = "伤" : Case "申" : s = "食"
          Case "酉" : s = "财" : Case "戌" : s = "贝"
          Case "亥" : s = "官" : Case "子" : s = "杀"
        End Select
      Case "巳" '火
        Select Case sZb
          Case "丑" : s = "杀" : Case "寅" : s = "官"
          Case "卯" : s = "枭" : Case "辰" : s = "印"
          Case "巳" : s = "比" : Case "午" : s = "劫"
          Case "未" : s = "对" : Case "申" : s = "吸"
          Case "酉" : s = "食" : Case "戌" : s = "伤"
          Case "亥" : s = "贝" : Case "子" : s = "财"
        End Select
      Case "午" '土
        Select Case sZb
          Case "丑" : s = "官" : Case "寅" : s = "杀"
          Case "卯" : s = "印" : Case "辰" : s = "枭"
          Case "巳" : s = "劫" : Case "午" : s = "比"
          Case "未" : s = "吸" : Case "申" : s = "对"
          Case "酉" : s = "伤" : Case "戌" : s = "食"
          Case "亥" : s = "财" : Case "子" : s = "贝"
        End Select
      Case "未" '土
        Select Case sZb
          Case "丑" : s = "贝" : Case "寅" : s = "财"
          Case "卯" : s = "杀" : Case "辰" : s = "官"
          Case "巳" : s = "枭" : Case "午" : s = "印"
          Case "未" : s = "比" : Case "申" : s = "劫"
          Case "酉" : s = "对" : Case "戌" : s = "吸"
          Case "亥" : s = "食" : Case "子" : s = "伤"
        End Select
      Case "申" '土
        Select Case sZb
          Case "丑" : s = "财" : Case "寅" : s = "贝"
          Case "卯" : s = "官" : Case "辰" : s = "杀"
          Case "巳" : s = "印" : Case "午" : s = "枭"
          Case "未" : s = "劫" : Case "申" : s = "比"
          Case "酉" : s = "吸" : Case "戌" : s = "对"
          Case "亥" : s = "伤" : Case "子" : s = "食"
        End Select
      Case "酉" '金
        Select Case sZb
          Case "丑" : s = "食" : Case "寅" : s = "伤"
          Case "卯" : s = "贝" : Case "辰" : s = "财"
          Case "巳" : s = "杀" : Case "午" : s = "官"
          Case "未" : s = "枭" : Case "申" : s = "印"
          Case "酉" : s = "比" : Case "戌" : s = "劫"
          Case "亥" : s = "对" : Case "子" : s = "吸"
        End Select
      Case "戌" '金
        Select Case sZb
          Case "丑" : s = "伤" : Case "寅" : s = "食"
          Case "卯" : s = "财" : Case "辰" : s = "贝"
          Case "巳" : s = "官" : Case "午" : s = "杀"
          Case "未" : s = "印" : Case "申" : s = "枭"
          Case "酉" : s = "劫" : Case "戌" : s = "比"
          Case "亥" : s = "吸" : Case "子" : s = "对"
        End Select
      Case "亥" '水
        Select Case sZb
          Case "丑" : s = "对" : Case "寅" : s = "吸"
          Case "卯" : s = "食" : Case "辰" : s = "伤"
          Case "巳" : s = "贝" : Case "午" : s = "财"
          Case "未" : s = "杀" : Case "申" : s = "官"
          Case "酉" : s = "枭" : Case "戌" : s = "印"
          Case "亥" : s = "比" : Case "子" : s = "劫"
        End Select
      Case "子"
        Select Case sZb
          Case "丑" : s = "吸" : Case "寅" : s = "对"
          Case "卯" : s = "伤" : Case "辰" : s = "食"
          Case "巳" : s = "财" : Case "午" : s = "贝"
          Case "未" : s = "官" : Case "申" : s = "杀"
          Case "酉" : s = "印" : Case "戌" : s = "枭"
          Case "亥" : s = "劫" : Case "子" : s = "比"
        End Select
    End Select
    Return s
  End Function
  '================================================================
  Public Function uf_印綬_干vs干_五行(sA As String, sB As String) As String
    Dim s = ""
    Select Case sA
      Case "甲"
        Select Case sB
          Case "甲" : s = "比" : Case "乙" : s = "劫"
          Case "丙" : s = "食" : Case "丁" : s = "伤"
          Case "戊" : s = "贝" : Case "己" : s = "财"
          Case "庚" : s = "杀" : Case "辛" : s = "官"
          Case "壬" : s = "枭" : Case "癸" : s = "印"
        End Select
      Case "乙"
        Select Case sB
          Case "甲" : s = "劫" : Case "乙" : s = "比"
          Case "丙" : s = "伤" : Case "丁" : s = "食"
          Case "戊" : s = "财" : Case "己" : s = "贝"
          Case "庚" : s = "官" : Case "辛" : s = "杀"
          Case "壬" : s = "印" : Case "癸" : s = "枭"
        End Select

      Case "丙"
        Select Case sB
          Case "甲" : s = "枭" : Case "乙" : s = "印"
          Case "丙" : s = "比" : Case "丁" : s = "劫"
          Case "戊" : s = "食" : Case "己" : s = "伤"
          Case "庚" : s = "贝" : Case "辛" : s = "财"
          Case "壬" : s = "杀" : Case "癸" : s = "官"
        End Select
      Case "丁"
        Select Case sB
          Case "甲" : s = "印" : Case "乙" : s = "枭"
          Case "丙" : s = "劫" : Case "丁" : s = "比"
          Case "戊" : s = "伤" : Case "己" : s = "食"
          Case "庚" : s = "财" : Case "辛" : s = "贝"
          Case "壬" : s = "官" : Case "癸" : s = "杀"
        End Select

      Case "戊"
        Select Case sB
          Case "甲" : s = "杀" : Case "乙" : s = "官"
          Case "丙" : s = "枭" : Case "丁" : s = "印"
          Case "戊" : s = "比" : Case "己" : s = "劫"
          Case "庚" : s = "食" : Case "辛" : s = "伤"
          Case "壬" : s = "贝" : Case "癸" : s = "财"
        End Select
      Case "己"
        Select Case sB
          Case "甲" : s = "官" : Case "乙" : s = "杀"
          Case "丙" : s = "印" : Case "丁" : s = "枭"
          Case "戊" : s = "劫" : Case "己" : s = "比"
          Case "庚" : s = "伤" : Case "辛" : s = "食"
          Case "壬" : s = "财" : Case "癸" : s = "贝"
        End Select

      Case "庚"
        Select Case sB
          Case "甲" : s = "贝" : Case "乙" : s = "财"
          Case "丙" : s = "杀" : Case "丁" : s = "官"
          Case "戊" : s = "枭" : Case "己" : s = "印"
          Case "庚" : s = "比" : Case "辛" : s = "劫"
          Case "壬" : s = "食" : Case "癸" : s = "伤"
        End Select
      Case "辛"
        Select Case sB
          Case "甲" : s = "财" : Case "乙" : s = "贝"
          Case "丙" : s = "官" : Case "丁" : s = "杀"
          Case "戊" : s = "印" : Case "己" : s = "枭"
          Case "庚" : s = "劫" : Case "辛" : s = "比"
          Case "壬" : s = "伤" : Case "癸" : s = "食"
        End Select

      Case "壬"
        Select Case sB
          Case "甲" : s = "食" : Case "乙" : s = "伤"
          Case "丙" : s = "贝" : Case "丁" : s = "财"
          Case "戊" : s = "杀" : Case "己" : s = "官"
          Case "庚" : s = "枭" : Case "辛" : s = "印"
          Case "壬" : s = "比" : Case "癸" : s = "劫"
        End Select
      Case "癸"
        Select Case sB
          Case "甲" : s = "伤" : Case "乙" : s = "食"
          Case "丙" : s = "财" : Case "丁" : s = "贝"
          Case "戊" : s = "官" : Case "己" : s = "杀"
          Case "庚" : s = "印" : Case "辛" : s = "枭"
          Case "壬" : s = "劫" : Case "癸" : s = "比"
        End Select
    End Select
    Return s
  End Function
  Public Function uf_印綬_干vs干_六行(sA As String, sB As String) As String
    ' 戊=燥土，己=湿土
    '              枭印 比劫 食伤
    ' 甲乙丙丁戊    父   兄   仔     父兄仔财奴官
    ' 庚辛壬癸己    财   奴   官     木火土金水元
    ' 6 7 8 9     贝财  奴迁 杀官
    Dim s = ""
    Select Case sA
      Case "甲"        ' 兄仔财奴官父
        Select Case sB ' 木火土金水元
          Case "甲" : s = "比" : Case "乙" : s = "劫"
          Case "丙" : s = "食" : Case "丁" : s = "伤"
          Case "戊" : s = "贝" '-> 此法缺點：12神中缺了：正财、偏印
          Case "庚" : s = "迁" : Case "辛" : s = "奴"
          Case "壬" : s = "杀" : Case "癸" : s = "官"
          Case "己" : s = "印"
        End Select
      Case "乙"
        Select Case sB
          Case "甲" : s = "劫" : Case "乙" : s = "比"
          Case "丙" : s = "伤" : Case "丁" : s = "食"
          Case "戊" : s = "财"
          Case "庚" : s = "奴" : Case "辛" : s = "迁"
          Case "壬" : s = "官" : Case "癸" : s = "杀"
          Case "己" : s = "枭"
        End Select

      Case "丙"        ' 父兄仔财奴官
        Select Case sB ' 木火土金水元
          Case "甲" : s = "枭" : Case "乙" : s = "印"
          Case "丙" : s = "比" : Case "丁" : s = "劫"
          Case "戊" : s = "食"
          Case "庚" : s = "贝" : Case "辛" : s = "财"
          Case "壬" : s = "迁" : Case "癸" : s = "奴"
          Case "己" : s = "官"
        End Select
      Case "丁"
        Select Case sB
          Case "甲" : s = "印" : Case "乙" : s = "枭"
          Case "丙" : s = "劫" : Case "丁" : s = "比"
          Case "戊" : s = "伤"
          Case "庚" : s = "财" : Case "辛" : s = "贝"
          Case "壬" : s = "奴" : Case "癸" : s = "迁"
          Case "己" : s = "杀"
        End Select

      Case "戊"        ' 官父兄仔财奴
        Select Case sB ' 木火土金水元
          Case "甲" : s = "杀" : Case "乙" : s = "官"
          Case "丙" : s = "枭" : Case "丁" : s = "印"
          Case "戊" : s = "比"
          Case "庚" : s = "食" : Case "辛" : s = "伤"
          Case "壬" : s = "贝" : Case "癸" : s = "财"
          Case "己" : s = "迁"
        End Select
      Case "己"        ' 仔财奴官父兄
        Select Case sB ' 木火土金水元
          Case "甲" : s = "伤" : Case "乙" : s = "食"
          Case "丙" : s = "财" : Case "丁" : s = "贝"
          Case "戊" : s = "迁"
          Case "庚" : s = "官" : Case "辛" : s = "杀"
          Case "壬" : s = "印" : Case "癸" : s = "枭"
          Case "己" : s = "比"
        End Select

      Case "庚"        ' 奴官父兄仔财
        Select Case sB ' 木火土金水元
          Case "甲" : s = "迁" : Case "乙" : s = "奴"
          Case "丙" : s = "杀" : Case "丁" : s = "官"
          Case "戊" : s = "枭"
          Case "庚" : s = "比" : Case "辛" : s = "劫"
          Case "壬" : s = "食" : Case "癸" : s = "伤"
          Case "己" : s = "财"
        End Select
      Case "辛"
        Select Case sB
          Case "甲" : s = "奴" : Case "乙" : s = "迁"
          Case "丙" : s = "官" : Case "丁" : s = "杀"
          Case "戊" : s = "印"
          Case "庚" : s = "劫" : Case "辛" : s = "比"
          Case "壬" : s = "伤" : Case "癸" : s = "食"
          Case "己" : s = "贝"
        End Select

      Case "壬"        ' 财奴官父兄仔
        Select Case sB ' 木火土金水元
          Case "甲" : s = "贝" : Case "乙" : s = "财"
          Case "丙" : s = "迁" : Case "丁" : s = "奴"
          Case "戊" : s = "杀"
          Case "庚" : s = "枭" : Case "辛" : s = "印"
          Case "壬" : s = "比" : Case "癸" : s = "劫"
          Case "己" : s = "伤"
        End Select
      Case "癸"
        Select Case sB
          Case "甲" : s = "财" : Case "乙" : s = "贝"
          Case "丙" : s = "奴" : Case "丁" : s = "迁"
          Case "戊" : s = "官"
          Case "庚" : s = "印" : Case "辛" : s = "枭"
          Case "壬" : s = "劫" : Case "癸" : s = "比"
          Case "己" : s = "食"
        End Select
    End Select

    'Select Case s
    '  Case "奴" : s = "迁"
    '  Case "迁" : s = "奴"
    'End Select
    Return s
  End Function
  Public Function uf_印綬_干vs干_六行_奇隅(sA As String, sB As String) As String
    ' 戊=燥土，己=湿土
    ' 枭印 比劫 食伤
    ' 甲乙 丙丁 戊
    ' 父   兄   仔     父兄仔财奴官
    ' 财   奴   官     木火土金水元
    ' 贝财 奴迁 杀官
    ' 庚辛 壬癸 己      
    Dim s = ""
    Select Case sA
      Case "甲"        ' 兄仔财奴官父
        Select Case sB ' 木火土金水元
          Case "甲" : s = "比" : Case "乙" : s = "劫"
          Case "丙" : s = "食" : Case "丁" : s = "伤"
          Case "戊" : s = "贝" '-> 此法缺點：12神中缺了：正财、偏印
          Case "庚" : s = "迁" : Case "辛" : s = "奴"
          Case "壬" : s = "官" : Case "癸" : s = "杀" ' 变
          Case "己" : s = "印"
        End Select
      Case "乙"
        Select Case sB
          Case "甲" : s = "劫" : Case "乙" : s = "比"
          Case "丙" : s = "伤" : Case "丁" : s = "食"
          Case "戊" : s = "财"
          Case "庚" : s = "奴" : Case "辛" : s = "迁"
          Case "壬" : s = "杀" : Case "癸" : s = "官" ' 变
          Case "己" : s = "枭"
        End Select

      Case "丙"        ' 父兄仔财奴官
        Select Case sB ' 木火土金水元
          Case "甲" : s = "枭" : Case "乙" : s = "印"
          Case "丙" : s = "比" : Case "丁" : s = "劫"
          Case "戊" : s = "食"
          Case "庚" : s = "财" : Case "辛" : s = "贝" ' 变
          Case "壬" : s = "迁" : Case "癸" : s = "奴"
          Case "己" : s = "官"
        End Select
      Case "丁"
        Select Case sB
          Case "甲" : s = "印" : Case "乙" : s = "枭"
          Case "丙" : s = "劫" : Case "丁" : s = "比"
          Case "戊" : s = "伤"
          Case "庚" : s = "贝" : Case "辛" : s = "财" ' 变
          Case "壬" : s = "奴" : Case "癸" : s = "迁"
          Case "己" : s = "杀"
        End Select

      Case "戊"        ' 官父兄仔财奴
        Select Case sB ' 木火土金水元
          Case "甲" : s = "杀" : Case "乙" : s = "官"
          Case "丙" : s = "枭" : Case "丁" : s = "印"
          Case "戊" : s = "比"
          Case "庚" : s = "伤" : Case "辛" : s = "食" ' 变
          Case "壬" : s = "财" : Case "癸" : s = "贝" ' 变
          Case "己" : s = "迁"
        End Select
      Case "己" ' 10   ' 仔财奴官父兄
        Select Case sB ' 木火土金水元
          Case "甲" : s = "伤" : Case "乙" : s = "食"
          Case "丙" : s = "财" : Case "丁" : s = "贝"
          Case "戊" : s = "迁"
          Case "庚" : s = "杀" : Case "辛" : s = "官" ' 变
          Case "壬" : s = "枭" : Case "癸" : s = "印" ' 变
          Case "己" : s = "比"
        End Select

      Case "庚" ' 6    ' 奴官父兄仔财
        Select Case sB ' 木火土金水元
          Case "甲" : s = "迁" : Case "乙" : s = "奴"
          Case "丙" : s = "官" : Case "丁" : s = "杀" ' 变
          Case "戊" : s = "印" ' 变
          Case "庚" : s = "比" : Case "辛" : s = "劫"
          Case "壬" : s = "食" : Case "癸" : s = "伤"
          Case "己" : s = "贝" ' 变
        End Select
      Case "辛" ' 7
        Select Case sB
          Case "甲" : s = "奴" : Case "乙" : s = "迁"
          Case "丙" : s = "杀" : Case "丁" : s = "官" ' 变
          Case "戊" : s = "枭" ' 变
          Case "庚" : s = "劫" : Case "辛" : s = "比"
          Case "壬" : s = "伤" : Case "癸" : s = "食"
          Case "己" : s = "财" ' 变
        End Select

      Case "壬" ' 8    ' 财奴官父兄仔
        Select Case sB ' 木火土金水元
          Case "甲" : s = "财" : Case "乙" : s = "贝" ' 变
          Case "丙" : s = "迁" : Case "丁" : s = "奴"
          Case "戊" : s = "官" ' 变
          Case "庚" : s = "枭" : Case "辛" : s = "印"
          Case "壬" : s = "比" : Case "癸" : s = "劫"
          Case "己" : s = "食" ' 变
        End Select
      Case "癸" ' 9
        Select Case sB
          Case "甲" : s = "贝" : Case "乙" : s = "财" ' 变
          Case "丙" : s = "奴" : Case "丁" : s = "迁"
          Case "戊" : s = "杀" ' 变
          Case "庚" : s = "印" : Case "辛" : s = "枭"
          Case "壬" : s = "劫" : Case "癸" : s = "比"
          Case "己" : s = "伤" ' 变
        End Select
    End Select

    'Select Case s
    '  Case "奴" : s = "迁"
    '  Case "迁" : s = "奴"
    'End Select
    Return s
  End Function
  Public Function us父2梟印(s As String) As String
    s = s.Replace("比", "").Replace("沖", "").Replace("合", "").Replace("墓", "　").Replace("　", "").TrimEnd
    s = s.Replace("-父", "枭").Replace("+父", "印")
    s = s.Replace("-兄", "比").Replace("+兄", "劫")
    s = s.Replace("-妻", "贝").Replace("+妻", "财")
    s = s.Replace("-仔", "食").Replace("+仔", "伤")
    s = s.Replace("-奴", "迁").Replace("+奴", "奴")
    s = s.Replace("-官", "杀").Replace("+官", "官")
    Return s
  End Function
End Module

