Module m字串

	Public Sub us_六爻_替換不要的字(ByRef s As String)
		On Error GoTo Er : GoTo 0
Er: m1.us_ErrMsg(New StackFrame(0, True), Err.Description) : Exit Sub
0:  '---------------------------------------------------
		'替換太長的字串，因从网络上(京房易)貼过來的卦例，都太長了
		's = s.Replace("^l", vbNewLine).Replace("^m", vbNewLine).Replace("^v", vbNewLine) ' 沒效
		's = s.Replace("^%", vbNewLine)' 沒效
		s = s.Replace("戍", "戌")
		s = s.Replace("* ", "＊").Replace("*", "＊").Replace("!", "！").Replace(",", "，").Replace("?", "？").Replace(":", "：")
		's显示 = s显示.Replace(":", "：") '不用,因時間 24:12 不要让它变
		s = s.Replace("？ ", "？").Replace(" ？", "？").Replace("　？", "？")
		s = s.Replace("， ", "，").Replace(" ，", "，").Replace("　，", "，").Replace("　，", "，")
		s = s.Replace("。 ", "。").Replace(" 。", "。").Replace("　。", "。")
		s = s.Replace("、 ", "、").Replace(" 、", "、").Replace("　、", "、")
		s = s.Replace("： ", "：").Replace(" ：", "：").Replace("　：", "：")

		s = s.Replace("宫 ：", "宮：").Replace("宫：", "宮：") '坎宫 ：雷火丰 =>「宫」不一样

		'乾
		s = s.Replace("乾宮：乾为天", "乾：乾")
		s = s.Replace("父母壬戌土", "父壬戌").Replace("兄弟壬申金", "兄壬申")
		s = s.Replace("官鬼壬午火", "官壬午").Replace("父母甲辰土", "父甲辰")
		s = s.Replace("妻财甲寅木", "妻甲寅").Replace("子孙甲子水", "仔甲子")
		'姤
		s = s.Replace("乾宮：天风姤", "乾：姤")
		s = s.Replace("父母壬戌土", "父壬戌").Replace("兄弟壬申金", "兄壬申")
		s = s.Replace("官鬼壬午火", "官壬午").Replace("兄弟辛酉金", "兄辛酉")
		s = s.Replace("子孙辛亥水", "仔辛亥").Replace("父母辛丑土", "父辛丑")
		s = s.Replace("妻财甲寅木", "妻甲寅") '伏
		'遁
		s = s.Replace("乾宮：天山遁", "乾：遁")
		s = s.Replace("父母壬戌土", "父壬戌").Replace("兄弟壬申金", "兄壬申")
		s = s.Replace("官鬼壬午火", "官壬午").Replace("兄弟丙申金", "兄丙申")
		s = s.Replace("官鬼丙午火", "官丙午").Replace("父母丙辰土", "父丙辰")
		s = s.Replace("妻财甲寅木", "妻甲寅").Replace("子孙甲子水", "仔甲子") '伏
		'否
		s = s.Replace("乾宮：天地否", "乾：否")
		s = s.Replace("父母壬戌土", "父壬戌").Replace("兄弟壬申金", "兄壬申")
		s = s.Replace("官鬼壬午火", "官壬午").Replace("妻财乙卯木", "妻乙卯")
		s = s.Replace("官鬼乙巳火", "官乙巳").Replace("父母乙未土", "父乙未")
		s = s.Replace("子孙甲子水", "仔甲子")
		'观
		s = s.Replace("乾宮：风地观", "乾：观")
		s = s.Replace("妻财辛卯木", "妻辛卯").Replace("官鬼辛巳火", "官辛巳")
		s = s.Replace("父母辛未土", "父辛未").Replace("妻财乙卯木", "妻乙卯")
		s = s.Replace("官鬼乙巳火", "官乙巳").Replace("父母乙未土", "父乙未")
		s = s.Replace("兄弟壬申金", "兄壬申").Replace("子孙甲子水", "仔甲子") '伏
		'剥　　　
		s = s.Replace("乾宮：山地剥", "乾：剥")
		s = s.Replace("妻财丙寅木", "妻丙寅").Replace("子孙丙子水", "仔丙子")
		s = s.Replace("父母丙戌土", "父丙戌").Replace("妻财乙卯木", "妻乙卯")
		s = s.Replace("官鬼乙巳火", "官乙巳").Replace("父母乙未土", "父乙未")
		s = s.Replace("兄弟壬申金", "兄壬申")
		'晋
		s = s.Replace("乾宮：火地晋", "乾：晋")
		s = s.Replace("官鬼己巳火", "官己巳").Replace("父母己未土", "父己未")
		s = s.Replace("兄弟己酉金", "兄己酉").Replace("妻财乙卯木", "妻乙卯")
		s = s.Replace("官鬼乙巳火", "官乙巳").Replace("父母乙未土", "父乙未")
		s = s.Replace("子孙甲子水", "仔甲子")
		'大有    
		s = s.Replace("乾宮：火天大有", "乾：大有")
		s = s.Replace("官鬼己巳火", "官己巳").Replace("父母己未土", "父己未")
		s = s.Replace("兄弟己酉金", "兄己酉").Replace("父母甲辰土", "父甲辰")
		s = s.Replace("妻财甲寅木", "妻甲寅").Replace("子孙甲子水", "仔甲子")
		'=====================================================
		'兑　　　
		s = s.Replace("兑宮：兑为泽", "兑：兑")
		s = s.Replace("父母丁未土", "父丁未").Replace("兄弟丁酉金", "兄丁酉")
		s = s.Replace("子孙丁亥水", "仔丁亥").Replace("父母丁丑土", "父丁丑")
		s = s.Replace("妻财丁卯木", "妻丁卯").Replace("官鬼丁巳火", "官丁巳")
		'困
		s = s.Replace("兑宮：泽水困", "兑：困")
		s = s.Replace("父母丁未土", "父丁未").Replace("兄弟丁酉金", "兄丁酉")
		s = s.Replace("子孙丁亥水", "仔丁亥").Replace("官鬼戊午火", "官戊午")
		s = s.Replace("父母戊辰土", "父戊辰").Replace("妻财戊寅木", "妻戊寅")
		'萃
		s = s.Replace("兑宮：泽地萃", "兑：萃")
		s = s.Replace("父母丁未土", "父丁未").Replace("兄弟丁酉金", "兄丁酉")
		s = s.Replace("子孙丁亥水", "仔丁亥").Replace("妻财乙卯木", "妻乙卯")
		s = s.Replace("官鬼乙巳火", "官乙巳").Replace("父母乙未土", "父乙未")

		s = s.Replace("兑宮：泽山咸", "兑：咸")
		s = s.Replace("父母丁未土", "父丁未").Replace("兄弟丁酉金", "兄丁酉")
		s = s.Replace("子孙丁亥水", "仔丁亥").Replace("兄弟丙申金", "兄丙申")
		s = s.Replace("官鬼丙午火", "官丙午").Replace("父母丙辰土", "父丙辰")
		s = s.Replace("妻财丁卯木", "妻丁卯") '伏
		'蹇25
		s = s.Replace("兑宮：水山蹇", "兑：蹇")
		s = s.Replace("子孙戊子水", "仔戊子").Replace("父母戊戌土", "父戊戌")
		s = s.Replace("兄弟戊申金", "兄戊申").Replace("兄弟丙申金", "兄丙申")
		s = s.Replace("官鬼丙午火", "官丙午").Replace("父母丙辰土", "父丙辰")
		s = s.Replace("妻财丁卯木", "妻丁卯") '伏
		'谦26
		s = s.Replace("兑宮：地山谦", "兑：谦")
		s = s.Replace("兄弟癸酉金", "兄癸酉").Replace("子孙癸亥水", "仔癸亥")
		s = s.Replace("父母癸丑土", "父癸丑").Replace("兄弟丙申金", "兄丙申")
		s = s.Replace("官鬼丙午火", "官丙午").Replace("父母丙辰土", "父丙辰")
		s = s.Replace("妻财丁卯木", "妻丁卯") '伏
		'小过27
		s = s.Replace("兑宮：雷山小过", "兑：小过")
		s = s.Replace("父母庚戌土", "父庚戌").Replace("兄弟庚申金", "兄庚申")
		s = s.Replace("官鬼庚午火", "官庚午").Replace("兄弟丙申金", "兄丙申")
		s = s.Replace("官鬼丙午火", "官丙午").Replace("父母丙辰土", "父丙辰")
		s = s.Replace("子孙丁亥水", "仔丁亥").Replace("妻财丁卯木", "妻丁卯") '伏
		'归妹
		s = s.Replace("兑宮：雷泽归妹", "兑：归妹")
		s = s.Replace("父母庚戌土", "父庚戌").Replace("兄弟庚申金", "兄庚申")
		s = s.Replace("官鬼庚午火", "官庚午").Replace("父母丁丑土", "父丁丑")
		s = s.Replace("妻财丁卯木", "妻丁卯").Replace("官鬼丁巳火", "官丁巳")
		s = s.Replace("子孙丁亥水", "仔丁亥") '伏
		'==================================================
		'离31
		s = s.Replace("离宮：离为火", "离：离")
		s = s.Replace("兄弟己巳火", "兄己巳").Replace("子孙己未土", "仔己未")
		s = s.Replace("妻财己酉金", "妻己酉").Replace("官鬼己亥水", "官己亥")
		s = s.Replace("子孙己丑土", "仔己丑").Replace("父母己卯木", "父己卯")
		'旅32
		s = s.Replace("离宮：火山旅", "离：旅")
		s = s.Replace("兄弟己巳火", "兄己巳").Replace("子孙己未土", "仔己未")
		s = s.Replace("妻财己酉金", "妻己酉").Replace("妻财丙申金", "妻丙申")
		s = s.Replace("兄弟丙午火", "兄丙午").Replace("子孙丙辰土", "仔丙辰")
		s = s.Replace("官鬼己亥水", "官己亥").Replace("父母己卯木", "父己卯") '伏
		'鼎33
		s = s.Replace("离宮：火风鼎", "离：鼎")
		s = s.Replace("兄弟己巳火", "兄己巳").Replace("子孙己未土", "仔己未")
		s = s.Replace("妻财己酉金", "妻己酉").Replace("妻财辛酉金", "妻辛酉")
		s = s.Replace("官鬼辛亥水", "官辛亥").Replace("子孙辛丑土", "仔辛丑")
		s = s.Replace("父母己卯木", "父己卯") '伏
		'未济34
		s = s.Replace("离宮：火水未济", "离：未济")
		s = s.Replace("兄弟己巳火", "兄己巳").Replace("子孙己未土", "仔己未")
		s = s.Replace("妻财己酉金", "妻己酉").Replace("兄弟戊午火", "兄戊午")
		s = s.Replace("子孙戊辰土", "仔戊辰").Replace("父母戊寅木", "父戊寅")
		s = s.Replace("官鬼己亥水", "官己亥") '伏
		'蒙35
		s = s.Replace("离宮：山水蒙", "离：蒙")
		s = s.Replace("父母丙寅木", "父丙寅").Replace("官鬼丙子水", "官丙子")
		s = s.Replace("子孙丙戌土", "仔丙戌").Replace("兄弟戊午火", "兄戊午")
		s = s.Replace("子孙戊辰土", "仔戊辰").Replace("父母戊寅木", "父戊寅")
		s = s.Replace("妻财己酉金", "妻己酉") '伏
		'涣36
		s = s.Replace("离宮：风水涣", "离：涣")
		s = s.Replace("父母辛卯木", "父辛卯").Replace("兄弟辛巳火", "兄辛巳")
		s = s.Replace("子孙辛未土", "仔辛未").Replace("兄弟戊午火", "兄戊午")
		s = s.Replace("子孙戊辰土", "仔戊辰").Replace("父母戊寅木", "父戊寅")
		s = s.Replace("妻财己酉金", "妻己酉").Replace("官鬼己亥水", "官己亥") '伏
		'讼37
		s = s.Replace("离宮：天水讼", "离：讼")
		s = s.Replace("子孙壬戌土", "仔壬戌").Replace("妻财壬申金", "妻壬申")
		s = s.Replace("兄弟壬午火", "兄壬午").Replace("兄弟戊午火", "兄戊午")
		s = s.Replace("子孙戊辰土", "仔戊辰").Replace("父母戊寅木", "父戊寅")
		s = s.Replace("官鬼己亥水", "官己亥") '伏
		'同人38
		s = s.Replace("离宮：天火同人", "离：同人")
		s = s.Replace("子孙壬戌土", "仔壬戌").Replace("妻财壬申金", "妻壬申")
		s = s.Replace("兄弟壬午火", "兄壬午").Replace("官鬼己亥水", "官己亥")
		s = s.Replace("子孙己丑土", "仔己丑").Replace("父母己卯木", "父己卯")
		'================================================
		'震41
		s = s.Replace("震宮：震为雷", "震：震")
		s = s.Replace("妻财庚戌土", "妻庚戌").Replace("官鬼庚申金", "官庚申")
		s = s.Replace("子孙庚午火", "仔庚午").Replace("妻财庚辰土", "妻庚辰")
		s = s.Replace("兄弟庚寅木", "兄庚寅").Replace("父母庚子水", "父庚子")
		'豫42:
		s = s.Replace("震宮：雷地豫", "震：豫")
		s = s.Replace("妻财庚戌土", "妻庚戌").Replace("官鬼庚申金", "官庚申")
		s = s.Replace("子孙庚午火", "仔庚午").Replace("兄弟乙卯木", "兄乙卯")
		s = s.Replace("子孙乙巳火", "仔乙巳").Replace("妻财乙未土", "妻乙未")
		s = s.Replace("父母庚子水", "父庚子") '伏
		'解43
		s = s.Replace("震宮：雷水解", "震：解")
		s = s.Replace("妻财庚戌土", "妻庚戌").Replace("官鬼庚申金", "官庚申")
		s = s.Replace("子孙庚午火", "仔庚午").Replace("子孙戊午火", "仔戊午")
		s = s.Replace("妻财戊辰土", "妻戊辰").Replace("兄弟戊寅木", "兄戊寅")
		s = s.Replace("父母庚子水", "父庚子") '伏
		'恒44
		s = s.Replace("震宮：雷风恒", "震：恒")
		s = s.Replace("妻财庚戌土", "妻庚戌").Replace("官鬼庚申金", "官庚申")
		s = s.Replace("子孙庚午火", "仔庚午").Replace("官鬼辛酉金", "官辛酉")
		s = s.Replace("父母辛亥水", "父辛亥").Replace("妻财辛丑土", "妻辛丑")
		s = s.Replace("兄弟庚寅木", "兄庚寅") '伏
		'升45
		s = s.Replace("震宮：地风升", "震：升")
		s = s.Replace("官鬼癸酉金", "官癸酉").Replace("父母癸亥水", "父癸亥")
		s = s.Replace("妻财癸丑土", "妻癸丑").Replace("官鬼辛酉金", "官辛酉")
		s = s.Replace("父母辛亥水", "父辛亥").Replace("妻财辛丑土", "妻辛丑")
		s = s.Replace("子孙庚午火", "仔庚午").Replace("兄弟庚寅木", "兄庚寅") '伏
		'井46
		s = s.Replace("震宮：水风井", "震：井")
		s = s.Replace("父母戊子水", "父戊子").Replace("妻财戊戌土", "妻戊戌")
		s = s.Replace("官鬼戊申金", "官戊申").Replace("官鬼辛酉金", "官辛酉")
		s = s.Replace("父母辛亥水", "父辛亥").Replace("妻财辛丑土", "妻辛丑")
		s = s.Replace("子孙庚午火", "仔庚午").Replace("兄弟庚寅木", "兄庚寅") '伏
		'大过47
		s = s.Replace("震宮：泽风大过", "震：大过")
		s = s.Replace("妻财丁未土", "妻丁未").Replace("官鬼丁酉金", "官丁酉")
		s = s.Replace("父母丁亥水", "父丁亥").Replace("官鬼辛酉金", "官辛酉")
		s = s.Replace("父母辛亥水", "父辛亥").Replace("妻财辛丑土", "妻辛丑")
		s = s.Replace("子孙庚午火", "仔庚午").Replace("兄弟庚寅木", "兄庚寅") '伏
		'随48
		s = s.Replace("震宮：泽雷随", "震：随")
		s = s.Replace("妻财丁未土", "妻丁未").Replace("官鬼丁酉金", "官丁酉")
		s = s.Replace("父母丁亥水", "父丁亥").Replace("妻财庚辰土", "妻庚辰")
		s = s.Replace("兄弟庚寅木", "兄庚寅").Replace("父母庚子水", "父庚子")
		s = s.Replace("子孙庚午火", "仔庚午") '伏
		'============================================
		'巽51
		s = s.Replace("巽宮：巽为风", "巽：巽")
		s = s.Replace("兄弟辛卯木", "兄辛卯").Replace("子孙辛巳火", "仔辛巳")
		s = s.Replace("妻财辛未土", "妻辛未").Replace("官鬼辛酉金", "官辛酉")
		s = s.Replace("父母辛亥水", "父辛亥").Replace("妻财辛丑土", "妻辛丑")
		'小畜52
		s = s.Replace("巽宮：风天小畜", "巽：小畜")
		s = s.Replace("兄弟辛卯木", "兄辛卯").Replace("子孙辛巳火", "仔辛巳")
		s = s.Replace("妻财辛未土", "妻辛未").Replace("妻财甲辰土", "妻甲辰")
		s = s.Replace("兄弟甲寅木", "兄甲寅").Replace("父母甲子水", "父甲子")
		s = s.Replace("官鬼辛酉金", "官辛酉") '伏
		'家人53
		s = s.Replace("巽宮：风火家人", "巽：家人")
		s = s.Replace("兄弟辛卯木", "兄辛卯").Replace("子孙辛巳火", "仔辛巳")
		s = s.Replace("妻财辛未土", "妻辛未").Replace("父母己亥水", "父己亥")
		s = s.Replace("妻财己丑土", "妻己丑").Replace("兄弟己卯木", "兄己卯")
		s = s.Replace("官鬼辛酉金", "官辛酉") '伏
		'益54
		s = s.Replace("巽宮：风雷益", "巽：益")
		s = s.Replace("兄弟辛卯木", "兄辛卯").Replace("子孙辛巳火", "仔辛巳")
		s = s.Replace("妻财辛未土", "妻辛未").Replace("妻财庚辰土", "妻庚辰")
		s = s.Replace("兄弟庚寅木", "兄庚寅").Replace("父母庚子水", "父庚子")
		s = s.Replace("官鬼辛酉金", "官辛酉") '伏
		'无妄55
		s = s.Replace("巽宮：天雷无妄", "巽：无妄")
		s = s.Replace("妻财壬戌土", "妻壬戌").Replace("官鬼壬申金", "官壬申")
		s = s.Replace("子孙壬午火", "仔壬午").Replace("妻财庚辰土", "妻庚辰")
		s = s.Replace("兄弟庚寅木", "兄庚寅").Replace("父母庚子水", "父庚子")
		'噬嗑56
		s = s.Replace("巽宮：火雷噬嗑", "巽：噬嗑")
		s = s.Replace("子孙己巳火", "仔己巳").Replace("妻财己未土", "妻己未")
		s = s.Replace("官鬼己酉金", "官己酉").Replace("妻财庚辰土", "妻庚辰")
		s = s.Replace("兄弟庚寅木", "兄庚寅").Replace("父母庚子水", "父庚子")
		'颐57
		s = s.Replace("巽宮：山雷颐", "巽：颐")
		s = s.Replace("兄弟丙寅木", "兄丙寅").Replace("父母丙子水", "父丙子")
		s = s.Replace("妻财丙戌土", "妻丙戌").Replace("妻财庚辰土", "妻庚辰")
		s = s.Replace("兄弟庚寅木", "兄庚寅").Replace("父母庚子水", "父庚子")
		s = s.Replace("子孙辛巳火", "仔辛巳").Replace("官鬼辛酉金", "官辛酉") '伏
		'蛊58
		s = s.Replace("巽宮：山风蛊", "巽：蛊")
		s = s.Replace("兄弟丙寅木", "兄丙寅").Replace("父母丙子水", "父丙子")
		s = s.Replace("妻财丙戌土", "妻丙戌").Replace("官鬼辛酉金", "官辛酉")
		s = s.Replace("父母辛亥水", "父辛亥").Replace("妻财辛丑土", "妻辛丑")
		s = s.Replace("子孙辛巳火", "仔辛巳") '伏
		'================================================
		'坎61
		s = s.Replace("坎宮：坎为水", "坎：坎")
		s = s.Replace("兄弟戊子水", "兄戊子").Replace("官鬼戊戌土", "官戊戌")
		s = s.Replace("父母戊申金", "父戊申").Replace("妻财戊午火", "妻戊午")
		s = s.Replace("官鬼戊辰土", "官戊辰").Replace("子孙戊寅木", "仔戊寅")
		'节62
		s = s.Replace("坎宮：水泽节", "坎：节")
		s = s.Replace("兄弟戊子水", "兄戊子").Replace("官鬼戊戌土", "官戊戌")
		s = s.Replace("父母戊申金", "父戊申").Replace("官鬼丁丑土", "官丁丑")
		s = s.Replace("子孙丁卯木", "仔丁卯").Replace("妻财丁巳火", "妻丁巳")
		'屯63
		s = s.Replace("坎宮：水雷屯", "坎：屯")
		s = s.Replace("兄弟戊子水", "兄戊子").Replace("兄弟戊子水", "兄戊子")
		s = s.Replace("官鬼戊戌土", "官戊戌").Replace("父母戊申金", "父戊申")
		s = s.Replace("官鬼庚辰土", "官庚辰").Replace("子孙庚寅木", "仔庚寅")
		s = s.Replace("兄弟庚子水", "兄庚子").Replace("妻财戊午火", "妻戊午") '伏
		'既济64
		s = s.Replace("坎宮：水火既济", "坎：既济")
		s = s.Replace("兄弟戊子水", "兄戊子").Replace("官鬼戊戌土", "官戊戌")
		s = s.Replace("父母戊申金", "父戊申").Replace("兄弟己亥水", "兄己亥")
		s = s.Replace("官鬼己丑土", "官己丑").Replace("子孙己卯木", "仔己卯")
		s = s.Replace("妻财戊午火", "妻戊午") '伏
		'革65
		s = s.Replace("坎宮：泽火革", "坎：革")
		s = s.Replace("官鬼丁未土", "官丁未").Replace("父母丁酉金", "父丁酉")
		s = s.Replace("兄弟丁亥水", "兄丁亥").Replace("兄弟己亥水", "兄己亥")
		s = s.Replace("官鬼己丑土", "官己丑").Replace("子孙己卯木", "仔己卯")
		s = s.Replace("妻财戊午火", "妻戊午")
		'丰66                  坎宫：雷火丰
		s = s.Replace("坎宮：雷火丰", "坎：丰")
		s = s.Replace("官鬼庚戌土", "官庚戌").Replace("父母庚申金", "父庚申")
		s = s.Replace("妻财庚午火", "妻庚午").Replace("兄弟己亥水", "兄己亥")
		s = s.Replace("官鬼己丑土", "官己丑").Replace("子孙己卯木", "仔己卯")
		'明夷67
		s = s.Replace("坎宮：地火明夷", "坎：明夷")
		s = s.Replace("父母癸酉金", "父癸酉").Replace("兄弟癸亥水", "兄癸亥")
		s = s.Replace("官鬼癸丑土", "官癸丑").Replace("兄弟己亥水", "兄己亥")
		s = s.Replace("官鬼己丑土", "官己丑").Replace("子孙己卯木", "仔己卯")
		s = s.Replace("妻财戊午火", "妻戊午")
		'师68
		s = s.Replace("坎宮：地水师", "坎：师")
		s = s.Replace("父母癸酉金", "父癸酉").Replace("兄弟癸亥水", "兄癸亥")
		s = s.Replace("官鬼癸丑土", "官癸丑").Replace("妻财戊午火", "妻戊午")
		s = s.Replace("官鬼戊辰土", "官戊辰").Replace("子孙戊寅木", "仔戊寅")
		'================================================
		'艮71
		s = s.Replace("艮宮：艮为山", "艮：艮")
		s = s.Replace("官鬼丙寅木", "官丙寅").Replace("妻财丙子水", "妻丙子")
		s = s.Replace("兄弟丙戌土", "兄丙戌").Replace("子孙丙申金", "仔丙申")
		s = s.Replace("父母丙午火", "父丙午").Replace("兄弟丙辰土", "兄丙辰")
		'贲72
		s = s.Replace("艮宮：山火贲", "艮：贲")
		s = s.Replace("官鬼丙寅木", "官丙寅").Replace("妻财丙子水", "妻丙子")
		s = s.Replace("兄弟丙戌土", "兄丙戌").Replace("妻财己亥水", "妻己亥")
		s = s.Replace("兄弟己丑土", "兄己丑").Replace("官鬼己卯木", "官己卯")
		s = s.Replace("子孙庚申金", "仔庚申").Replace("父母丙午火", "父丙午") '伏
		'大畜73
		s = s.Replace("艮宮：山天大畜", "艮：大畜")
		s = s.Replace("官鬼丙寅木", "官丙寅").Replace("妻财丙子水", "妻丙子")
		s = s.Replace("兄弟丙戌土", "兄丙戌").Replace("兄弟甲辰土", "兄甲辰")
		s = s.Replace("官鬼甲寅木", "官甲寅").Replace("妻财甲子水", "妻甲子")
		s = s.Replace("子孙庚申金", "仔庚申").Replace("父母丙午火", "父丙午") '伏
		'损74
		s = s.Replace("艮宮：山泽损", "艮：损")
		s = s.Replace("官鬼丙寅木", "官丙寅").Replace("妻财丙子水", "妻丙子")
		s = s.Replace("兄弟丙戌土", "兄丙戌").Replace("兄弟丁丑土", "兄丁丑")
		s = s.Replace("官鬼丁卯木", "官丁卯").Replace("父母丁巳火", "父丁巳")
		s = s.Replace("子孙丙申金", "仔丙申") '伏
		'睽75
		s = s.Replace("艮宮：火泽睽", "艮：睽")
		s = s.Replace("父母己巳火", "父己巳").Replace("兄弟己未土", "兄己未")
		s = s.Replace("子孙己酉金", "仔己酉").Replace("兄弟丁丑土", "兄丁丑")
		s = s.Replace("官鬼丁卯木", "官丁卯").Replace("父母丁巳火", "父丁巳")
		s = s.Replace("妻财丙子水", "妻丙子") '伏
		'履76
		s = s.Replace("艮宮：天泽履", "艮：履")
		s = s.Replace("兄弟壬戌土", "兄壬戌").Replace("子孙壬申金", "仔壬申")
		s = s.Replace("父母壬午火", "父壬午").Replace("兄弟丁丑土", "兄丁丑")
		s = s.Replace("官鬼丁卯木", "官丁卯").Replace("父母丁巳火", "父丁巳")
		s = s.Replace("妻财丙子水", "妻丙子") '伏
		'中孚77
		s = s.Replace("艮宮：风泽中孚", "艮：中孚")
		s = s.Replace("官鬼辛卯木", "官辛卯").Replace("父母辛巳火", "父辛巳")
		s = s.Replace("兄弟辛未土", "兄辛未").Replace("兄弟丁丑土", "兄丁丑")
		s = s.Replace("官鬼丁卯木", "官丁卯").Replace("父母丁巳火", "父丁巳")
		s = s.Replace("妻财丙子水", "妻丙子").Replace("子孙丙申金", "仔丙申") '伏
		'渐78
		s = s.Replace("艮宮：风山渐", "艮：渐")
		s = s.Replace("官鬼辛卯木", "官辛卯").Replace("父母辛巳火", "父辛巳")
		s = s.Replace("兄弟辛未土", "兄辛未").Replace("子孙丙申金", "仔丙申")
		s = s.Replace("父母丙午火", "父丙午").Replace("兄弟丙辰土", "兄丙辰")
		s = s.Replace("妻财丙子水", "妻丙子") '伏
		'================================================
		'坤81
		s = s.Replace("坤宮：坤为地", "坤：坤")
		s = s.Replace("子孙癸酉金", "仔癸酉").Replace("妻财癸亥水", "妻癸亥")
		s = s.Replace("兄弟癸丑土", "兄癸丑").Replace("官鬼乙卯木", "官乙卯")
		s = s.Replace("父母乙巳火", "父乙巳").Replace("兄弟乙未土", "兄乙未")
		'复82
		s = s.Replace("坤宮：地雷复", "坤：复")
		s = s.Replace("子孙癸酉金", "仔癸酉").Replace("妻财癸亥水", "妻癸亥")
		s = s.Replace("兄弟癸丑土", "兄癸丑").Replace("兄弟庚辰土", "兄庚辰")
		s = s.Replace("官鬼庚寅木", "官庚寅").Replace("妻财庚子水", "妻庚子")
		s = s.Replace("父母乙巳火", "父乙巳") '伏
		'临83
		s = s.Replace("坤宮：地泽临", "坤：临")
		s = s.Replace("子孙癸酉金", "仔癸酉").Replace("妻财癸亥水", "妻癸亥")
		s = s.Replace("兄弟癸丑土", "兄癸丑").Replace("兄弟丁丑土", "兄丁丑")
		s = s.Replace("官鬼丁卯木", "官丁卯").Replace("父母丁巳火", "父丁巳")
		'泰84
		s = s.Replace("坤宮：地天泰", "坤：泰")
		s = s.Replace("子孙癸酉金", "仔癸酉").Replace("妻财癸亥水", "妻癸亥")
		s = s.Replace("兄弟癸丑土", "兄癸丑").Replace("兄弟甲辰土", "兄甲辰")
		s = s.Replace("官鬼甲寅木", "官甲寅").Replace("妻财甲子水", "妻甲子")
		s = s.Replace("父母乙巳火", "父乙巳") '伏
		'大壮85
		s = s.Replace("坤宮：雷天大壮", "坤：大壮")
		s = s.Replace("兄弟庚戌土", "兄庚戌").Replace("子孙庚申金", "仔庚申")
		s = s.Replace("父母庚午火", "父庚午").Replace("兄弟甲辰土", "兄甲辰")
		s = s.Replace("官鬼甲寅木", "官甲寅").Replace("妻财甲子水", "妻甲子")
		'夬86
		s = s.Replace("坤宮：泽天夬", "坤：夬")
		s = s.Replace("兄弟丁未土", "兄丁未").Replace("子孙丁酉金", "仔丁酉")
		s = s.Replace("妻财丁亥水", "妻丁亥").Replace("兄弟甲辰土", "兄甲辰")
		s = s.Replace("官鬼甲寅木", "官甲寅").Replace("妻财甲子水", "妻甲子")
		s = s.Replace("父母乙巳火", "父乙巳")
		'需87
		s = s.Replace("坤宮：水天需", "坤：需")
		s = s.Replace("妻财戊子水", "妻戊子").Replace("兄弟戊戌土", "兄戊戌")
		s = s.Replace("子孙戊申金", "仔戊申").Replace("兄弟甲辰土", "兄甲辰")
		s = s.Replace("官鬼甲寅木", "官甲寅").Replace("妻财甲子水", "妻甲子")
		s = s.Replace("父母乙巳火", "父乙巳") '伏
		'比88
		s = s.Replace("坤宮：水地比", "坤：比")
		s = s.Replace("妻财戊子水", "妻戊子").Replace("兄弟戊戌土", "兄戊戌")
		s = s.Replace("子孙戊申金", "仔戊申").Replace("官鬼乙卯木", "官乙卯")
		s = s.Replace("父母乙巳火", "父乙巳").Replace("兄弟乙未土", "兄乙未")
		'============================================
		'变卦处的替換，視需要加入
		s = s.Replace("父母甲寅木", "父甲寅")
		s = s.Replace("父母乙卯木", "父乙卯")
		s = s.Replace("父母丁卯木", "父丁卯")
		s = s.Replace("父母戊午火", "父戊午")
		s = s.Replace("父母己丑土", "父己丑").Replace("父母己酉金", "父己酉")
		s = s.Replace("父母庚辰土", "父庚辰").Replace("父母庚寅木", "父庚寅")
		s = s.Replace("父母辛酉金", "父辛酉")
		s = s.Replace("父母壬申金", "父壬申")

		s = s.Replace("兄弟甲子水", "兄甲子")
		s = s.Replace("兄弟乙巳火", "兄乙巳")
		s = s.Replace("兄弟丙子水", "兄丙子")
		s = s.Replace("兄弟丁巳火", "兄丁巳").Replace("兄弟丁卯木", "兄丁卯")
		s = s.Replace("兄弟戊辰土", "兄戊辰")
		s = s.Replace("兄弟庚午火", "兄庚午")
		s = s.Replace("兄弟辛亥水", "兄辛亥")

		s = s.Replace("子孙乙未土", "仔乙未")
		s = s.Replace("子孙丙午火", "仔丙午").Replace("子孙丙寅木", "仔丙寅")
		s = s.Replace("子孙丁丑土", "仔丁丑").Replace("子孙丁巳火", "仔丁巳")
		s = s.Replace("子孙己亥水", "仔己亥")
		s = s.Replace("子孙庚戌土", "仔庚戌")

		s = s.Replace("妻财乙巳火", "妻乙巳")
		s = s.Replace("官鬼丙申金", "官丙申").Replace("妻财丙辰土", "妻丙辰")
		s = s.Replace("妻财丁丑土", "妻丁丑").Replace("妻财丁酉金", "妻丁酉")
		s = s.Replace("妻财己卯木", "妻己卯").Replace("妻财己巳火", "妻己巳")
		s = s.Replace("妻财庚申金", "妻庚申").Replace("妻财庚寅木", "妻庚寅")
		s = s.Replace("妻财辛巳火", "妻辛巳")
		s = s.Replace("妻财壬午火", "妻壬午")
		s = s.Replace("妻财癸酉金", "妻癸酉")

		s = s.Replace("官鬼甲子水", "官甲子").Replace("官鬼甲辰土", "官甲辰")
		s = s.Replace("官鬼乙未土", "官乙未")
		s = s.Replace("官鬼丙戌土", "官丙戌")
		s = s.Replace("官鬼丁亥水", "官丁亥")
		s = s.Replace("官鬼戊寅木", "官戊寅")
		s = s.Replace("官鬼己未土", "官己未")
		s = s.Replace("官鬼庚子水", "官庚子")
		s = s.Replace("官鬼辛丑土", "官辛丑")
		s = s.Replace("官鬼壬戌土", "官壬戌")
		s = s.Replace("官鬼癸亥水", "官癸亥")

		s = s.Replace("子孙甲寅木", "仔甲寅").Replace("子孙甲辰土", "仔甲辰")
		s = s.Replace("子孙乙卯木", "仔乙卯")
		s = s.Replace("子孙丁未土", "仔丁未")
		s = s.Replace("子孙庚子水", "仔庚子").Replace("子孙庚辰土", "仔庚辰")
		s = s.Replace("子孙辛卯木", "仔辛卯")
		s = s.Replace("子孙癸丑土", "仔癸丑")
		'============================================
		s = s.Replace(" [发布者已设置隐藏，任务完成之前回复仅专家和标主可见]", "")

		s = s.Replace("起卦方式：手动摇卦 　　龙隐测客()六爻线上排盘系统", "")
		s = s.Replace("起卦方式：手工指定 　　龙隐网(www.longyin.net)六爻线上排盘系统 ", "")
		s = s.Replace("起卦方式：手动摇卦 　　龙隐网(www.longyin.net)六爻线上排盘系统 ", "")

		s = s.Replace("六神 　伏　　神 【本　　卦】　　　　　　　　　【变　　卦】", "")
		s = s.Replace("六神 　伏　　神 【本　　卦】　　　　　　　【变　　卦】", "")
		s = s.Replace("六神　 伏　　神　【本　　卦】 　　　　 【变　　卦】", "")
		s = s.Replace("六神　 伏　　神　【本　　卦】 　　　　　　 【变　　卦】", "")
		s = s.Replace("六神 　伏　　神　【本　　卦】", "")
		s = s.Replace("六神 　伏　　神　本　　卦", "")
		s = s.Replace("六神　　伏神　　　本　 　　卦　　 　 　　　　变　 　　卦", "")
		s = s.Replace("六神　伏神　 　本　 　　卦　　 　 　　 　　变　 　　卦", "")
		s = s.Replace("六神 【本　　卦】　　　　　　　　　【变　　卦】", "")
		s = s.Replace("六神 【本　　卦】　　　　　　　【变　　卦】", "")
		s = s.Replace("六神 【本　　卦】　　　　　　【变　　卦】", "")
		s = s.Replace("六神 【本　　卦】", "")

		s = s.Replace("▅▅▅▅▅", "１").Replace("▄▄▄▄▄", "１").Replace("▅▅▅", "１")
		s = s.Replace("▅▅　▅▅", "２").Replace("▄▄ ▄▄", "２").Replace("▅　▅", "２").Replace("▄▄  ▄▄", "２")
		s = s.Replace("━━━", "１")
		s = s.Replace("━　━", "２")

		s = s.Replace("(六合)", "(合)").Replace("(六冲)", "(冲)")
		s = s.Replace("(游魂)", "(游)").Replace("(归魂)", "(归)")
		s = s.Replace("（六合）", "（合）").Replace("（六冲）", "（冲）")
		s = s.Replace("（游魂）", "（游）").Replace("（归魂）", "（归）")
		's = s.Replace("　　　　　", "　　　") '伏卦区5個空白，第2次用時，又会再縮成3格，关掉不用
		s = s.Replace("公历起卦时间", "公历").Replace("公历时间", "公历")
		s = s.Replace("以下是回复内容", "")

		s = s.Replace("干　　支：", "干支：").Replace("旬　　空：", "旬空：").Replace("神　　煞：", "神煞：")

		s = s.Replace("0.	", "0. ").Replace("1.	", "1. ").Replace("2.	", "2. ").Replace("3.	", "3. ")
		s = s.Replace("4.	", "4. ").Replace("5.	", "5. ").Replace("6.	", "6. ").Replace("7.	", "7. ")
		s = s.Replace("8.	", "8. ").Replace("9.	", "9. ")
		'==================================================
		s = s.Replace("勾陈　", "勾　").Replace("勾陈 ", "勾　")
		s = s.Replace("螣蛇　", "蛇　").Replace("螣蛇 ", "蛇　").Replace("腾蛇　", "蛇　").Replace("腾蛇 ", "蛇　")
		s = s.Replace("白虎　", "虎　").Replace("白虎 ", "虎　")
		s = s.Replace("玄武　", "玄　").Replace("玄武 ", "玄　")
		s = s.Replace("青龙　", "龙　").Replace("青龙 ", "龙　")
		s = s.Replace("朱雀　", "雀　").Replace("朱雀 ", "雀　")
		'==================================================
	End Sub
End Module
