Module m四柱
	Public Function uf單柱引字_五行(sGZ As String) As String
		Dim s = "" ' 天干s地支=帶象=象
		Select Case sGZ
			Case "甲子" : s += "g丑[子],象" ' 甲(寅)子->丑[子]
			Case "甲寅" : s += "L" ' 甲(寅)寅L
			Case "甲辰" : s += "[木],头" ' 甲(寅)辰[木]
			Case "甲午" : s += "g戌[午],象" ' 甲(寅)午->戌[午]
			Case "甲申" : s += "^,脚" ' 甲(寅)申^
			Case "甲戌" : s += "g午,头" ' 甲(寅)戌->午

			Case "乙丑" : s += "g寅,头" ' 乙(卯)丑->寅
			Case "乙卯" : s += "L" ' 乙(卯)卯L
			Case "乙巳" : s += "zz增,g辰[乙],象" ' 乙(卯)巳zz增->辰[乙]
			Case "乙未" : s += "[乙],g亥,头" ' 乙(卯)未[乙]->亥
			Case "乙酉" : s += "^,脚" ' 乙(卯)酉^
			Case "乙亥" : s += "g未[乙],象" ' 乙(卯)亥->未[乙]

			Case "丙子" : s += "脚"  ' 丙(巳)子
			Case "丙寅" : s += "#,象" ' 丙(巳)寅#
			Case "丙辰" : s += "象" ' 丙(巳)辰
			Case "丙午" : s += "R" ' 丙(巳)午R
			Case "丙申" : s += "#,头" ' 丙(巳)申#
			Case "丙戌" : s += "[丙]zz,象" ' 丙(巳)戌zz[丙]

			Case "丁丑" : s += "t,象" ' 丁(午)丑t
			Case "丁卯" : s += "p" ' 丁(午)卯p
			Case "丁巳" : s += "R" ' 丁(午)巳R
			Case "丁未" : s += "[丁],象" ' 丁(午)未[丁]
			Case "丁酉" : s += "p,头" ' 丁(午)酉p
			Case "丁亥" : s += "zz,脚" ' 丁(午)亥zz

			Case "戊子" : s += "zz,头" ' 戊(巳)子zz
			Case "戊寅" : s += "#,脚" ' 戊(巳)寅#
			Case "戊辰" : s += "zz增" ' 戊(巳)辰zz增
			Case "戊午" : s += "R,象" ' 戊(巳)午R
			Case "戊申" : s += "#kv,象" ' 戊(巳)申#kv
			Case "戊戌" : s += "[戊]" ' 戊(巳)戌[戊]
				'---------------------------------------------
		' 要分上下半部的原因，两者不同
		' 五行 丁.禄=午，刃=未，午:=丁己
		' 六行 己.禄=亥，刃=子     
			Case "己丑" : s += "t" ' 己(午)丑t
			Case "己卯" : s += "p,脚" ' 己(午)卯p
			Case "己巳" : s += "R,象" ' 己(午)巳R
			Case "己未" : s += "[己]" ' 己(午)未[己]
			Case "己酉" : s += "p,象" ' 己(午)酉p
			Case "己亥" : s += "zz,头" ' 己(午)亥zz

			Case "庚子" : s += "g辰[子],象" ' 庚(申)子->辰[子]
			Case "庚寅" : s += "^,头" ' 庚(申)寅^
			Case "庚辰" : s += "zz增,g子" ' 庚(申)辰zz增->子
			Case "庚午" : s += "g未,脚"  ' 庚(申)午
			Case "庚申" : s += "L" ' 庚(申)申L
			Case "庚戌" : s += "[金],g酉,象" ' 庚(申)戌[]->酉

			Case "辛丑" : s += "[辛],g巳" ' 辛(酉)丑[辛]->巳
			Case "辛卯" : s += "^,头" ' 辛(酉)卯^
			Case "辛巳" : s += "zz,g丑[辛],脚" ' 辛(酉)巳zz->丑[辛]
			Case "辛未" : s += "g申,象" ' 辛(酉)未->申
			Case "辛酉" : s += "L" ' 辛(酉)酉L
			Case "辛亥" : s += "g戌[辛],象" ' 辛(酉)亥->戌[辛]

			Case "壬子" : s += "R" ' 壬(亥)子R
			Case "壬寅" : s += "v,象" ' 壬(亥)寅v
			Case "壬辰" : s += "[水],脚" ' 壬(亥)辰[水]
			Case "壬午" : s += "zz,头" ' 壬(亥)午zz
			Case "壬申" : s += "#,象" ' 壬(亥)申#
			Case "壬戌" : s += "zz,脚" ' 壬(亥)戌zz

			Case "癸丑" : s += "[癸],v,脚" ' 癸(子)丑[癸]v
			Case "癸卯" : s += "p,象" ' 癸(子)卯p
			Case "癸巳" : s += "zz,头" ' 癸(子)巳zz
			Case "癸未" : s += "t,脚" ' 癸(子)未t
			Case "癸酉" : s += "p,象" ' 癸(子)酉p
			Case "癸亥" : s += "R" ' 癸(子)亥R
		End Select
		Return s
	End Function
	Public Function uf單柱引字_六行(sGZ As String) As String
		Dim s = ""
		Select Case sGZ
			Case "甲子" : s += "j丑[子]"  ' 甲(寅)子->丑[子]
			Case "甲寅" : s += "L"        ' 甲(寅)寅L
			Case "甲辰" : s += "zz=木,[木]"     ' 甲(寅)辰[木]
			Case "甲午" : s += "g戌[午]"  ' 甲(寅)午->戌[午]
			Case "甲申" : s += "^"       ' 甲(寅)申^
			Case "甲戌" : s += "zz=0,g午"     ' 甲(寅)戌->午

			Case "乙丑" : s += "zz=木,j寅"  '  乙(卯)丑zz->寅
			Case "乙卯" : s += "L"       ' 乙(卯)卯L
			Case "乙巳" : s += "zz=0,j辰[乙]"  ' 乙(卯)巳zz->辰[乙]
			Case "乙未" : s += "[乙],g亥"     ' 乙(卯)未[乙]->亥
			Case "乙酉" : s += "^"       ' 乙(卯)酉^
			Case "乙亥" : s += "g未[乙]"  ' 乙(卯)亥->未[乙]

			Case "丙子" : s += "zz=0"  ' 丙(巳)子zz=0
			Case "丙寅" : s += "#"     ' 丙(巳)寅#
			Case "丙辰" : s += "zz=0"  ' 丙(巳)辰zz=0
			Case "丙午" : s += "R"     ' 丙(巳)午R
			Case "丙申" : s += "#kv"   ' 丙(巳)申#kv
			Case "丙戌" : s += "[丙]"  ' 丙(巳)戌[丙]

			Case "丁丑" : s += "t"      ' 丁(午)丑t
			Case "丁卯" : s += "p"      ' 丁(午)卯p
			Case "丁巳" : s += "R"      ' 丁(午)巳R
			Case "丁未" : s += "[丁],午未v"   ' 丁(午)未[丁]
			Case "丁酉" : s += "p"      ' 丁(午)酉p
			Case "丁亥" : s += "zz=0"  ' 丁(午)亥zz=木

			Case "戊子" : s += ""             ' 戊(巳)[未戌]子
			Case "戊寅" : s += "#,未[寅],g午"  ' 戊(巳)[未戌]u寅#，未[寅]，->午
			Case "戊辰" : s += "3#未戌辰,戌辰^" ' 戊(巳)[未戌]u辰，3#未戌辰,戌辰^
			Case "戊午" : s += "未午v、戌午3v"  ' 戊(巳)[未戌]u午R、戌午3v
			Case "戊申" : s += "#kv,j酉"       ' 戊(巳)[未戌]u申#->酉
			Case "戊戌" : s += "[戊]"          ' 戊(巳)[未戌]u戌[戊]
		' =壬  ----------------------------------------------
			Case "己丑" : s += "j子"           ' 己(亥)[丑辰]u丑[己癸辛]->子
			Case "己卯" : s += "zz=木,g未[卯],j寅" ' 己(亥)[丑辰]u卯[乙]zz=木，->未[卯]
			Case "己巳" : s += "^,g酉"         ' 己(亥)[丑辰]u巳[丙庚戊]^
			Case "己未" : s += "zz=木,^#,g卯"   ' 己(亥)[丑辰]u未[戊丁乙]^#zz=木，->卯
			Case "己酉" : s += "j戌[辛],g巳,辰酉v" ' 己(亥)[丑辰]u酉[辛]->戌[辛]、->巳、辰酉v
			Case "己亥" : s += "L##"           ' 己(亥)[丑辰]亥L##

			Case "庚丑" : s += "[金]" ' 庚(申)丑[金]
			Case "庚卯" : s += "zz=0" ' 庚(申)卯zz=戊
			Case "庚巳" : s += "#"     ' 庚(申)巳#
			Case "庚未" : s += "zz=0" ' 庚(申)未zz=戊，庚未[戊]zz=金
			Case "庚酉" : s += "R"     ' 庚(申)酉R
			Case "庚亥" : s += "#"     ' 庚(申)亥

			Case "辛子" : s += "p" ' 辛(酉)子p
			Case "辛寅" : s += "zz=0" ' 辛(酉)寅zz=己
			Case "辛辰" : s += "v" ' 辛(酉)辰v
			Case "辛午" : s += "p" ' 辛(酉)午p
			Case "辛申" : s += "R" ' 辛(酉)申R
			Case "辛戌" : s += "zz=金,[辛]" ' 辛(酉)戌[辛]

			Case "壬丑" : s += "[水],j子" ' 壬(亥)丑[水]->子
			Case "壬卯" : s += "g未[卯]"  ' 壬(亥)卯->未[卯]
			Case "壬巳" : s += "^"       ' 壬(亥)巳^
			Case "壬未" : s += "zz=0,g卯"  ' 壬(亥)未zz->卯
			Case "壬酉" : s += "j戌[酉]" ' 壬(亥)酉->戌[酉]
			Case "壬亥" : s += "L##"     ' 壬(亥)亥L##

			Case "癸子" : s += "L"          '癸(子)子L
			Case "癸寅" : s += "zz=0,j丑[癸]" ' 癸(子)寅zz->丑[癸]
			Case "癸辰" : s += "[癸],g申"  ' 癸(子)辰[癸]->申
			Case "癸午" : s += "zz=0,^"  ' 癸(子)午zz=0,^
			Case "癸申" : s += "g辰[癸]"  ' 癸(子)申->辰[癸]
			Case "癸戌" : s += "j亥"      ' 癸(子)戌->亥
		End Select
		Return s
	End Function
	Public Function uf單柱引字_六行_配置4_早期(sGZ As String) As String
		Dim s = ""
		Select Case sGZ
			Case "甲子" : s += "g丑[子]" ' 甲(寅)子->丑[子]
			Case "甲寅" : s += "L"       ' 甲(寅)寅L
			Case "甲辰" : s += "[木]"    ' 甲(寅)辰[木]
			Case "甲午" : s += "g戌[午]" ' 甲(寅)午->戌[午]
			Case "甲申" : s += "^"      ' 甲(寅)申^
			Case "甲戌" : s += "g午"    ' 甲(寅)戌->午

			Case "乙丑" : s += "zz,g寅"    ' 乙(卯)丑zz->寅
			Case "乙卯" : s += "L"         ' 乙(卯)卯L
			Case "乙巳" : s += "g辰[乙]"   ' 乙(卯)巳zz->辰[乙]
			Case "乙未" : s += "zz=木,g亥" ' 乙(卯)未[乙]zz=木->亥
			Case "乙酉" : s += "^"        ' 乙(卯)酉^
			Case "乙亥" : s += "g未[乙]"   ' 乙(卯)亥->未[乙]

			Case "丙子" : s += "zz=金" ' 丙(巳)子zz=金
			Case "丙寅" : s += "#"     ' 丙(巳)寅#
			Case "丙辰" : s += "zz=金" ' 丙(巳)辰zz=金
			Case "丙午" : s += "R"     ' 丙(巳)午R
			Case "丙申" : s += "#kv"   ' 丙(巳)申#kv
			Case "丙戌" : s += "[丙]"  ' 丙(巳)戌[丙]

			Case "丁丑" : s += "t"     ' 丁(午)丑t
			Case "丁卯" : s += "p"     ' 丁(午)卯p
			Case "丁巳" : s += "R"     ' 丁(午)巳R
			Case "丁未" : s += "[丁]"   ' 丁(午)未[丁]
			Case "丁酉" : s += "p"     ' 丁(午)酉p
			Case "丁亥" : s += "zz=木" ' 丁(午)亥zz=木

			Case "戊子" ': s += ""           ' 戊子
			Case "戊寅" : s += "#,未[寅],g午" ' 戊(巳)[未戌]u寅#，未[寅]
			Case "戊辰" : s += "zz,3#未戌辰"  ' 戊(巳)[未戌]u辰zz、3#未戌辰
			Case "戊午" : s += "R,戌午3v"    ' 戊(巳)[未戌]u午R、戌午3v
			Case "戊申" : s += "#,g酉"       ' 戊(巳)[未戌]u申#->酉
			Case "戊戌" : s += "[戊]"        ' 戊(巳)[未戌]u戌[戊]
		' =壬  ----------------------------------------------
			Case "己丑" : s += "g子,zz=金" ' 己(亥)[丑辰]u丑[己癸辛]->子,zz=金
			Case "己卯" : s += "g未[卯]"   ' 己(亥)[丑辰]u卯[乙]   ->未[卯]
			Case "己巳" : s += "^,g酉" ' 己(亥)[丑辰]u巳[丙庚戊]^
			Case "己未" : s += "^#,g卯"    ' 己(亥)[丑辰]u未[戊丁乙]^#->卯
			Case "己酉" : s += "zz=金,g戌[辛],g巳,辰酉v" ' 己(亥)[丑辰]u酉[辛]zz=金->戌[辛]、->巳、辰酉v
			Case "己亥" : s += "L##"       ' 己(亥)[丑辰]亥L##

			Case "庚丑" : s += "[金]"  ' 庚(申)丑[金]
			Case "庚卯" : s += "zz=水" ' 庚(申)卯zz=水
			Case "庚巳" : s += "#"     ' 庚(申)巳#
			Case "庚未" : s += "zz=水" ' 庚(申)未zz=水
			Case "庚酉" : s += "R"     ' 庚(申)酉R
			Case "庚亥" ': s += ""     ' 庚(申)亥

			Case "辛子" : s += "p"     ' 辛(酉)子p
			Case "辛寅" : s += "zz=火" ' 辛(酉)寅zz=火
			Case "辛辰" : s += "zz=金,v" ' 辛(酉)辰zz=金、v
			Case "辛午" : s += "p"     ' 辛(酉)午p
			Case "辛申" : s += "R"     ' 辛(酉)申R
			Case "辛戌" : s += "[辛]"  ' 辛(酉)戌[辛]

			Case "壬丑" : s += "g子"     ' 壬(亥)丑[水]->子
			Case "壬卯" : s += "g未[卯]" ' 壬(亥)卯->未[卯]
			Case "壬巳" : s += "^"      ' 壬(亥)巳^
			Case "壬未" : s += "zz,g卯" ' 壬(亥)未zz->卯
			Case "壬酉" : s += "g戌[酉]" ' 壬(亥)酉->戌[酉]
			Case "壬亥" : s += "L##"    ' 壬(亥)亥L##

			Case "癸子" : s += "L"          ' 癸(子)子L
			Case "癸寅" : s += "zz,g丑[癸]" ' 癸(子)寅zz->丑[癸]
			Case "癸辰" : s += "[癸],g申"   ' 癸(子)辰[癸]->申
			Case "癸午" : s += "^"         ' 癸(子)午^
			Case "癸申" : s += "g辰[癸]"    ' 癸(子)申->辰[癸]
			Case "癸戌" : s += "g亥"        ' 癸(子)戌->亥
		End Select
		Return s
	End Function

	Public Function uf干支巔倒_五行(sGZ As String) As String
		Dim s = ""
		Select Case sGZ ' 丑未=己，未戌=戊
			Case "甲子" : s += "癸寅" : Case "甲寅" : s += "甲寅" : Case "甲辰" : s += "戊寅" : Case "甲午" : s += "丁寅" : Case "甲申" : s += "庚寅" : Case "甲戌" : s += "戊寅"
			Case "丙子" : s += "癸巳" : Case "丙寅" : s += "甲巳" : Case "丙辰" : s += "戊巳" : Case "丙午" : s += "丁巳" : Case "丙申" : s += "庚巳" : Case "丙戌" : s += "戊巳"
			Case "戊子" : s += "癸未" : Case "戊寅" : s += "甲未" : Case "戊辰" : s += "戊未" : Case "戊午" : s += "丁未" : Case "戊申" : s += "庚未" : Case "戊戌" : s += "戊未"
			Case "庚子" : s += "癸申" : Case "庚寅" : s += "甲申" : Case "庚辰" : s += "戊申" : Case "庚午" : s += "丁申" : Case "庚申" : s += "庚申" : Case "庚戌" : s += "戊申"
			Case "壬子" : s += "癸亥" : Case "壬寅" : s += "甲亥" : Case "壬辰" : s += "戊亥" : Case "壬午" : s += "丁亥" : Case "壬申" : s += "庚亥" : Case "壬戌" : s += "戊亥"

			Case "乙丑" : s += "己卯" : Case "乙卯" : s += "乙卯" : Case "乙巳" : s += "丙卯" : Case "乙未" : s += "己卯" : Case "乙酉" : s += "辛卯" : Case "乙亥" : s += "壬卯"
			Case "丁丑" : s += "己午" : Case "丁卯" : s += "乙午" : Case "丁巳" : s += "丙午" : Case "丁未" : s += "己午" : Case "丁酉" : s += "辛午" : Case "丁亥" : s += "壬午"
			Case "己丑" : s += "己丑" : Case "己卯" : s += "乙丑" : Case "己巳" : s += "丙丑" : Case "己未" : s += "己丑" : Case "己酉" : s += "辛丑" : Case "己亥" : s += "壬丑"
			Case "辛丑" : s += "己酉" : Case "辛卯" : s += "乙酉" : Case "辛巳" : s += "丙酉" : Case "辛未" : s += "己酉" : Case "辛酉" : s += "辛酉" : Case "辛亥" : s += "壬酉"
			Case "癸丑" : s += "己子" : Case "癸卯" : s += "乙子" : Case "癸巳" : s += "丙子" : Case "癸未" : s += "己子" : Case "癸酉" : s += "辛子" : Case "癸亥" : s += "壬子"
		End Select
		Return s
	End Function
	Public Function uf干支巔倒_六行(sGZ As String) As String
		Dim s = ""
		Select Case sGZ
			' 支->干，不可變 未戌=戊，丑辰=己
			' 干->支，可  變 戊=未戌，己=丑辰
			Case "甲子" : s += "癸寅" : Case "甲寅" : s += "甲寅" : Case "甲辰" : s += "己寅"
			Case "甲午" : s += "丁寅?丁卯" : Case "甲申" : s += "庚寅?庚酉" : Case "甲戌" : s += "戊寅"

			Case "丙子" : s += "癸巳?癸午" : Case "丙寅" : s += "甲巳?甲午" : Case "丙辰" : s += "己巳"
			Case "丙午" : s += "丁巳" : Case "丙申" : s += "庚巳" : Case "丙戌" : s += "戊巳?戊午"

			Case "辛子" : s += "癸酉?癸申" : Case "辛寅" : s += "甲酉?甲申" : Case "辛辰" : s += "己酉"
			Case "辛午" : s += "丁酉" : Case "辛申" : s += "庚酉" : Case "辛戌" : s += "戊酉?戊申"

			Case "癸子" : s += "癸子" : Case "癸寅" : s += "甲子" : Case "癸辰" : s += "己子?己亥"
			Case "癸午" : s += "丁子?丁亥" : Case "癸申" : s += "庚子?庚亥" : Case "癸戌" : s += "戊子"

			Case "乙丑" : s += "己卯" : Case "乙卯" : s += "乙卯" : Case "乙巳" : s += "丙卯?丙寅"
			Case "乙未" : s += "戊卯?戊寅" : Case "乙酉" : s += "辛卯?辛寅" : Case "乙亥" : s += "壬卯"

			Case "丁丑" : s += "己午?己巳" : Case "丁卯" : s += "乙午?乙巳" : Case "丁巳" : s += "丙午"
			Case "丁未" : s += "戊午" : Case "丁酉" : s += "辛午" : Case "丁亥" : s += "壬午?壬巳"

			Case "庚丑" : s += "己申?己酉" : Case "庚卯" : s += "乙申?乙酉" : Case "庚巳" : s += "丙申"
			Case "庚未" : s += "戊申" : Case "庚酉" : s += "辛申" : Case "庚亥" : s += "壬申?壬酉"

			Case "壬丑" : s += "己亥" : Case "壬卯" : s += "乙亥" : Case "壬巳" : s += "丙亥?丙子"
			Case "壬未" : s += "戊亥?戊子" : Case "壬酉" : s += "辛亥?辛子" : Case "壬亥" : s += "壬亥"
				' 干-支 可變
			Case "戊子" : s += "癸戌" : Case "戊寅" : s += "甲戌" : Case "戊辰" : s += "己未"
			Case "戊午" : s += "丁未" : Case "戊申" : s += "庚未" : Case "戊戌" : s += "戊戌"

			Case "己丑" : s += "己丑" : Case "己卯" : s += "乙丑" : Case "己巳" : s += "丙辰"
			Case "己未" : s += "戊辰" : Case "己酉" : s += "辛辰" : Case "己亥" : s += "壬丑"
		End Select
		Return s
	End Function

	Public Function uf天地沖合_五行(s0 As String) As String
		Dim s = ""  ' s0=八字
		' 天合+地合
		If s0.Contains("甲子") AndAlso s0.Contains("己丑") Then s += "甲子zv己丑" + vbNL
		If s0.Contains("甲寅") AndAlso s0.Contains("己亥") Then s += "甲寅zv己亥" + vbNL
		If s0.Contains("甲辰") AndAlso s0.Contains("己酉") Then s += "甲辰zv己酉" + vbNL
		If s0.Contains("甲午") AndAlso s0.Contains("己未") Then s += "甲午zv己未" + vbNL
		If s0.Contains("甲申") AndAlso s0.Contains("己巳") Then s += "甲申zv己巳" + vbNL
		If s0.Contains("甲戌") AndAlso s0.Contains("己卯") Then s += "甲戌zv己卯" + vbNL

		If s0.Contains("丙子") AndAlso s0.Contains("辛丑") Then s += "丙子zv辛丑" + vbNL
		If s0.Contains("丙寅") AndAlso s0.Contains("辛亥") Then s += "丙寅zv辛亥" + vbNL
		If s0.Contains("丙辰") AndAlso s0.Contains("辛酉") Then s += "丙辰zv辛酉" + vbNL
		If s0.Contains("丙午") AndAlso s0.Contains("辛未") Then s += "丙午zv辛未" + vbNL
		If s0.Contains("丙申") AndAlso s0.Contains("辛巳") Then s += "丙申zv辛巳" + vbNL
		If s0.Contains("丙戌") AndAlso s0.Contains("辛卯") Then s += "丙戌zv辛卯" + vbNL

		If s0.Contains("庚子") AndAlso s0.Contains("乙丑") Then s += "庚子zv乙丑" + vbNL
		If s0.Contains("庚寅") AndAlso s0.Contains("乙亥") Then s += "庚寅zv乙亥" + vbNL
		If s0.Contains("庚辰") AndAlso s0.Contains("乙酉") Then s += "庚辰zv乙酉" + vbNL
		If s0.Contains("庚午") AndAlso s0.Contains("乙未") Then s += "庚午zv乙未" + vbNL
		If s0.Contains("庚申") AndAlso s0.Contains("乙巳") Then s += "庚申zv乙巳" + vbNL
		If s0.Contains("庚戌") AndAlso s0.Contains("乙卯") Then s += "庚戌zv乙卯" + vbNL

		If s0.Contains("壬子") AndAlso s0.Contains("丁丑") Then s += "壬子zv丁丑" + vbNL
		If s0.Contains("壬寅") AndAlso s0.Contains("丁亥") Then s += "壬寅zv丁亥" + vbNL
		If s0.Contains("壬辰") AndAlso s0.Contains("丁酉") Then s += "壬辰zv丁酉" + vbNL
		If s0.Contains("壬午") AndAlso s0.Contains("丁未") Then s += "壬午zv丁未" + vbNL
		If s0.Contains("壬申") AndAlso s0.Contains("丁巳") Then s += "壬申zv丁巳" + vbNL
		If s0.Contains("壬戌") AndAlso s0.Contains("丁卯") Then s += "壬戌zv丁卯" + vbNL

		If s0.Contains("戊子") AndAlso s0.Contains("癸丑") Then s += "戊子zv癸丑" + vbNL
		If s0.Contains("戊寅") AndAlso s0.Contains("癸亥") Then s += "戊寅zv癸亥" + vbNL
		If s0.Contains("戊辰") AndAlso s0.Contains("癸酉") Then s += "戊辰zv癸酉" + vbNL
		If s0.Contains("戊午") AndAlso s0.Contains("癸未") Then s += "戊午zv癸未" + vbNL
		If s0.Contains("戊申") AndAlso s0.Contains("癸巳") Then s += "戊申zv癸巳" + vbNL
		If s0.Contains("戊戌") AndAlso s0.Contains("癸卯") Then s += "戊戌zv癸卯" + vbNL
		' 天沖+地沖
		If s0.Contains("甲子") AndAlso s0.Contains("庚午") Then s += "甲子^^庚午" + vbNL
		If s0.Contains("甲寅") AndAlso s0.Contains("庚申") Then s += "甲寅^^庚申" + vbNL
		If s0.Contains("甲辰") AndAlso s0.Contains("庚戌") Then s += "甲辰^^庚戌" + vbNL
		If s0.Contains("甲午") AndAlso s0.Contains("庚子") Then s += "甲午^^庚子" + vbNL
		If s0.Contains("甲申") AndAlso s0.Contains("庚寅") Then s += "甲申^^庚寅" + vbNL
		If s0.Contains("甲戌") AndAlso s0.Contains("庚辰") Then s += "甲戌^^庚辰" + vbNL

		If s0.Contains("丙子") AndAlso s0.Contains("壬午") Then s += "丙子^^壬午" + vbNL
		If s0.Contains("丙寅") AndAlso s0.Contains("壬申") Then s += "丙寅^^壬申" + vbNL
		If s0.Contains("丙辰") AndAlso s0.Contains("壬戌") Then s += "丙辰^^壬戌" + vbNL
		If s0.Contains("丙午") AndAlso s0.Contains("壬子") Then s += "丙午^^壬子" + vbNL
		If s0.Contains("丙申") AndAlso s0.Contains("壬寅") Then s += "丙申^^壬寅" + vbNL
		If s0.Contains("丙戌") AndAlso s0.Contains("壬辰") Then s += "丙戌^^壬辰" + vbNL

		If s0.Contains("乙丑") AndAlso s0.Contains("辛未") Then s += "乙丑^^辛未" + vbNL
		If s0.Contains("乙亥") AndAlso s0.Contains("辛巳") Then s += "乙亥^^辛巳" + vbNL
		If s0.Contains("乙酉") AndAlso s0.Contains("辛卯") Then s += "乙酉^^辛卯" + vbNL
		If s0.Contains("乙未") AndAlso s0.Contains("辛丑") Then s += "乙未^^辛丑" + vbNL
		If s0.Contains("乙巳") AndAlso s0.Contains("辛亥") Then s += "乙巳^^辛亥" + vbNL
		If s0.Contains("乙卯") AndAlso s0.Contains("辛酉") Then s += "乙卯^^辛酉" + vbNL

		If s0.Contains("丁丑") AndAlso s0.Contains("癸未") Then s += "丁丑^^癸未" + vbNL
		If s0.Contains("丁亥") AndAlso s0.Contains("癸巳") Then s += "丁亥^^癸巳" + vbNL
		If s0.Contains("丁酉") AndAlso s0.Contains("癸卯") Then s += "丁酉^^癸卯" + vbNL
		If s0.Contains("丁未") AndAlso s0.Contains("癸丑") Then s += "丁未^^癸丑" + vbNL
		If s0.Contains("丁巳") AndAlso s0.Contains("癸亥") Then s += "丁巳^^癸亥" + vbNL
		If s0.Contains("丁卯") AndAlso s0.Contains("癸酉") Then s += "丁卯^^癸酉" + vbNL
		Return s
	End Function
	Public Function uf天地沖合_六行(s0 As String) As String
		Dim s = ""  ' s0=八字
		' 天合+地沖
		If s0.Contains("甲子") AndAlso s0.Contains("辛午") Then s += "甲子z^辛午" + vbNL
		If s0.Contains("甲寅") AndAlso s0.Contains("辛申") Then s += "甲寅z^辛申" + vbNL
		If s0.Contains("甲辰") AndAlso s0.Contains("辛戌") Then s += "甲辰z^辛戌" + vbNL
		If s0.Contains("甲午") AndAlso s0.Contains("辛子") Then s += "甲午z^辛子" + vbNL
		If s0.Contains("甲申") AndAlso s0.Contains("辛寅") Then s += "甲申z^辛寅" + vbNL
		If s0.Contains("甲戌") AndAlso s0.Contains("辛辰") Then s += "甲戌z^辛辰" + vbNL

		If s0.Contains("丙子") AndAlso s0.Contains("癸午") Then s += "丙子z^癸午" + vbNL
		If s0.Contains("丙寅") AndAlso s0.Contains("癸申") Then s += "丙寅z^癸申" + vbNL
		If s0.Contains("丙辰") AndAlso s0.Contains("癸戌") Then s += "丙辰z^癸戌" + vbNL
		If s0.Contains("丙午") AndAlso s0.Contains("癸子") Then s += "丙午z^癸子" + vbNL
		If s0.Contains("丙申") AndAlso s0.Contains("癸寅") Then s += "丙申z^癸寅" + vbNL
		If s0.Contains("丙戌") AndAlso s0.Contains("癸辰") Then s += "丙戌z^癸辰" + vbNL

		If s0.Contains("庚未") AndAlso s0.Contains("乙丑") Then s += "庚未z^乙丑" + vbNL
		If s0.Contains("庚巳") AndAlso s0.Contains("乙亥") Then s += "庚巳z^乙亥" + vbNL
		If s0.Contains("庚卯") AndAlso s0.Contains("乙酉") Then s += "庚卯z^乙酉" + vbNL
		If s0.Contains("庚丑") AndAlso s0.Contains("乙未") Then s += "庚丑z^乙未" + vbNL
		If s0.Contains("庚亥") AndAlso s0.Contains("乙巳") Then s += "庚亥z^乙巳" + vbNL
		If s0.Contains("庚酉") AndAlso s0.Contains("乙卯") Then s += "庚酉z^乙卯" + vbNL

		If s0.Contains("壬未") AndAlso s0.Contains("丁丑") Then s += "壬未z^丁丑" + vbNL
		If s0.Contains("壬巳") AndAlso s0.Contains("丁亥") Then s += "壬巳z^丁亥" + vbNL
		If s0.Contains("壬卯") AndAlso s0.Contains("丁酉") Then s += "壬卯z^丁酉" + vbNL
		If s0.Contains("壬丑") AndAlso s0.Contains("丁未") Then s += "壬丑z^丁未" + vbNL
		If s0.Contains("壬亥") AndAlso s0.Contains("丁巳") Then s += "壬亥z^丁巳" + vbNL
		If s0.Contains("壬酉") AndAlso s0.Contains("丁卯") Then s += "壬酉z^丁卯" + vbNL

		If s0.Contains("戊子") AndAlso s0.Contains("辛午") Then s += "戊子z^辛午" + vbNL
		If s0.Contains("戊寅") AndAlso s0.Contains("辛申") Then s += "戊寅z^辛申" + vbNL
		If s0.Contains("戊辰") AndAlso s0.Contains("辛戌") Then s += "戊辰z^辛戌" + vbNL
		If s0.Contains("戊午") AndAlso s0.Contains("辛子") Then s += "戊午z^辛子" + vbNL
		If s0.Contains("戊申") AndAlso s0.Contains("辛寅") Then s += "戊申z^辛寅" + vbNL
		If s0.Contains("戊戌") AndAlso s0.Contains("辛辰") Then s += "戊戌z^辛辰" + vbNL

		If s0.Contains("乙丑") AndAlso s0.Contains("己未") Then s += "乙丑z^己未" + vbNL
		If s0.Contains("乙卯") AndAlso s0.Contains("己酉") Then s += "乙卯z^己酉" + vbNL
		If s0.Contains("乙巳") AndAlso s0.Contains("己亥") Then s += "乙巳z^己亥" + vbNL
		If s0.Contains("乙未") AndAlso s0.Contains("己丑") Then s += "乙未z^己丑" + vbNL
		If s0.Contains("乙酉") AndAlso s0.Contains("己卯") Then s += "乙酉z^己卯" + vbNL
		If s0.Contains("乙亥") AndAlso s0.Contains("己巳") Then s += "乙亥z^己巳" + vbNL

		' 天沖+地合
		If s0.Contains("甲子") AndAlso s0.Contains("庚丑") Then s += "甲子^v庚丑" + vbNL
		If s0.Contains("甲寅") AndAlso s0.Contains("庚亥") Then s += "甲寅^v庚亥" + vbNL
		If s0.Contains("甲辰") AndAlso s0.Contains("庚酉") Then s += "甲辰^v庚酉" + vbNL
		If s0.Contains("甲午") AndAlso s0.Contains("庚未") Then s += "甲午^v庚未" + vbNL
		If s0.Contains("甲申") AndAlso s0.Contains("庚巳") Then s += "甲申^v庚巳" + vbNL
		If s0.Contains("甲戌") AndAlso s0.Contains("庚卯") Then s += "甲戌^v庚卯" + vbNL

		If s0.Contains("丙子") AndAlso s0.Contains("壬丑") Then s += "丙子^v壬丑" + vbNL
		If s0.Contains("丙寅") AndAlso s0.Contains("壬亥") Then s += "丙寅^v壬亥" + vbNL
		If s0.Contains("丙辰") AndAlso s0.Contains("壬酉") Then s += "丙辰^v壬酉" + vbNL
		If s0.Contains("丙午") AndAlso s0.Contains("壬未") Then s += "丙午^v壬未" + vbNL
		If s0.Contains("丙申") AndAlso s0.Contains("壬巳") Then s += "丙申^v壬巳" + vbNL
		If s0.Contains("丙戌") AndAlso s0.Contains("壬卯") Then s += "丙戌^v壬卯" + vbNL

		If s0.Contains("戊子") AndAlso s0.Contains("己丑") Then s += "戊子^v己丑" + vbNL
		If s0.Contains("戊寅") AndAlso s0.Contains("己亥") Then s += "戊寅^v己亥" + vbNL
		If s0.Contains("戊辰") AndAlso s0.Contains("己酉") Then s += "戊辰^v己酉" + vbNL
		If s0.Contains("戊午") AndAlso s0.Contains("己未") Then s += "戊午^v己未" + vbNL
		If s0.Contains("戊申") AndAlso s0.Contains("己巳") Then s += "戊申^v己巳" + vbNL
		If s0.Contains("戊戌") AndAlso s0.Contains("己卯") Then s += "戊戌^v己卯" + vbNL

		If s0.Contains("乙丑") AndAlso s0.Contains("辛子") Then s += "乙丑^v辛子" + vbNL
		If s0.Contains("乙亥") AndAlso s0.Contains("辛寅") Then s += "乙亥^v辛寅" + vbNL
		If s0.Contains("乙酉") AndAlso s0.Contains("辛辰") Then s += "乙酉^v辛辰" + vbNL
		If s0.Contains("乙未") AndAlso s0.Contains("辛午") Then s += "乙未^v辛午" + vbNL
		If s0.Contains("乙巳") AndAlso s0.Contains("辛申") Then s += "乙巳^v辛申" + vbNL
		If s0.Contains("乙卯") AndAlso s0.Contains("辛戌") Then s += "乙卯^v辛戌" + vbNL

		If s0.Contains("丁丑") AndAlso s0.Contains("癸子") Then s += "丁丑^v癸子" + vbNL
		If s0.Contains("丁亥") AndAlso s0.Contains("癸寅") Then s += "丁亥^v癸寅" + vbNL
		If s0.Contains("丁酉") AndAlso s0.Contains("癸辰") Then s += "丁酉^v癸辰" + vbNL
		If s0.Contains("丁未") AndAlso s0.Contains("癸午") Then s += "丁未^v癸午" + vbNL
		If s0.Contains("丁巳") AndAlso s0.Contains("癸申") Then s += "丁巳^v癸申" + vbNL
		If s0.Contains("丁卯") AndAlso s0.Contains("癸戌") Then s += "丁卯^v癸戌" + vbNL
		Return s
	End Function

	Public Function uf地支組合_五行(s0 As String) As String
		Dim s = ""
		If s0.Contains("子") AndAlso s0.Contains("丑") Then s += "子v丑->土" + vbNL
		If s0.Contains("寅") AndAlso s0.Contains("亥") Then s += "寅z亥->木" + vbNL
		If s0.Contains("卯") AndAlso s0.Contains("戌") Then s += "卯z戌->火" + vbNL
		If s0.Contains("辰") AndAlso s0.Contains("酉") Then s += "辰z酉->金" + vbNL
		If s0.Contains("巳") AndAlso s0.Contains("申") Then s += "巳z申->水" + vbNL
		If s0.Contains("午") AndAlso s0.Contains("未") Then s += "午z未->土" + vbNL

		If s0.Contains("子") AndAlso s0.Contains("卯") Then s += "子p卯" + vbNL
		If s0.Contains("卯") AndAlso s0.Contains("午") Then s += "卯p午" + vbNL
		If s0.Contains("午") AndAlso s0.Contains("酉") Then s += "午p酉" + vbNL
		If s0.Contains("酉") AndAlso s0.Contains("子") Then s += "酉p子" + vbNL

		If s0.Contains("子") AndAlso s0.Contains("午") Then s += "子^午" + vbNL
		If s0.Contains("丑") AndAlso s0.Contains("未") Then s += "丑^未" + vbNL
		If s0.Contains("寅") AndAlso s0.Contains("申") Then s += "寅^申" + vbNL
		If s0.Contains("卯") AndAlso s0.Contains("酉") Then s += "卯^酉" + vbNL
		If s0.Contains("辰") AndAlso s0.Contains("戌") Then s += "辰^戌" + vbNL
		If s0.Contains("巳") AndAlso s0.Contains("亥") Then s += "巳^亥" + vbNL

		If s0.Contains("子") AndAlso s0.Contains("未") Then s += "子t未" + vbNL
		If s0.Contains("丑") AndAlso s0.Contains("午") Then s += "丑t午" + vbNL
		If s0.Contains("寅") AndAlso s0.Contains("巳") Then s += "寅t巳" + vbNL
		If s0.Contains("卯") AndAlso s0.Contains("辰") Then s += "卯t辰" + vbNL
		If s0.Contains("申") AndAlso s0.Contains("亥") Then s += "申t亥" + vbNL
		If s0.Contains("酉") AndAlso s0.Contains("戌") Then s += "酉t戌" + vbNL

		If s0.Contains("未") AndAlso s0.Contains("戌") AndAlso s0.Contains("丑") Then s += "未戌丑3#" + vbNL
		If s0.Contains("戌") AndAlso s0.Contains("丑") AndAlso s0.Contains("辰") Then s += "戌丑辰3#" + vbNL
		If s0.Contains("丑") AndAlso s0.Contains("辰") AndAlso s0.Contains("未") Then s += "丑辰未3#" + vbNL
		If s0.Contains("辰") AndAlso s0.Contains("未") AndAlso s0.Contains("戌") Then s += "辰未戌3#" + vbNL

		If s0.Contains("寅") AndAlso s0.Contains("巳") AndAlso s0.Contains("申") Then s += "寅巳申3#" + vbNL
		If s0.Contains("巳") AndAlso s0.Contains("申") AndAlso s0.Contains("亥") Then s += "巳申亥3#" + vbNL
		If s0.Contains("申") AndAlso s0.Contains("亥") AndAlso s0.Contains("寅") Then s += "申亥寅3#" + vbNL
		If s0.Contains("亥") AndAlso s0.Contains("寅") AndAlso s0.Contains("巳") Then s += "亥寅巳3#" + vbNL

		If s0.Contains("丑") AndAlso s0.Contains("寅") Then s += "丑az寅" + vbNL
		If s0.Contains("卯") AndAlso s0.Contains("申") Then s += "卯az申" + vbNL
		If s0.Contains("酉") AndAlso s0.Contains("寅") Then s += "酉az寅" + vbNL
		If s0.Contains("午") AndAlso s0.Contains("亥") Then s += "午az亥" + vbNL
		If s0.Contains("子") AndAlso s0.Contains("巳") Then s += "子az巳" + vbNL
		'---------------------------------------------
		If s0.Contains("寅") AndAlso s0.Contains("午") AndAlso s0.Contains("戌") Then s += "寅{午}戌" + vbNL
		If s0.Contains("申") AndAlso s0.Contains("子") AndAlso s0.Contains("辰") Then s += "申{子}辰" + vbNL
		If s0.Contains("亥") AndAlso s0.Contains("卯") AndAlso s0.Contains("未") Then s += "亥{卯}未" + vbNL
		If s0.Contains("巳") AndAlso s0.Contains("酉") AndAlso s0.Contains("丑") Then s += "巳{酉}丑" + vbNL

		If Not s0.Contains("寅") AndAlso s0.Contains("午") AndAlso s0.Contains("戌") Then s += "午戌g寅" + vbNL
		If s0.Contains("寅") AndAlso Not s0.Contains("午") AndAlso s0.Contains("戌") Then s += "寅戌g午" + vbNL
		If s0.Contains("寅") AndAlso s0.Contains("午") AndAlso Not s0.Contains("戌") Then s += "寅午g戌" + vbNL

		If Not s0.Contains("申") AndAlso s0.Contains("子") AndAlso s0.Contains("辰") Then s += "子辰g申" + vbNL
		If s0.Contains("申") AndAlso Not s0.Contains("子") AndAlso s0.Contains("辰") Then s += "申辰g子" + vbNL
		If s0.Contains("申") AndAlso s0.Contains("子") AndAlso Not s0.Contains("辰") Then s += "申子g辰" + vbNL

		If Not s0.Contains("亥") AndAlso s0.Contains("卯") AndAlso s0.Contains("未") Then s += "卯未g亥" + vbNL
		If s0.Contains("亥") AndAlso Not s0.Contains("卯") AndAlso s0.Contains("未") Then s += "亥未g卯" + vbNL
		If s0.Contains("亥") AndAlso s0.Contains("卯") AndAlso Not s0.Contains("未") Then s += "亥卯g未" + vbNL

		If Not s0.Contains("巳") AndAlso s0.Contains("酉") AndAlso s0.Contains("丑") Then s += "酉丑g巳" + vbNL
		If s0.Contains("巳") AndAlso Not s0.Contains("酉") AndAlso s0.Contains("丑") Then s += "巳丑g酉" + vbNL
		If s0.Contains("巳") AndAlso s0.Contains("酉") AndAlso Not s0.Contains("丑") Then s += "巳酉g丑" + vbNL
		'---------------------------------------------
		If s0.Contains("亥") AndAlso s0.Contains("子") AndAlso s0.Contains("丑") Then s += "亥{子}丑" + vbNL
		If s0.Contains("寅") AndAlso s0.Contains("卯") AndAlso s0.Contains("辰") Then s += "寅{卯}辰" + vbNL
		If s0.Contains("巳") AndAlso s0.Contains("午") AndAlso s0.Contains("未") Then s += "巳{午}未" + vbNL
		If s0.Contains("申") AndAlso s0.Contains("酉") AndAlso s0.Contains("戌") Then s += "申{酉}戌" + vbNL

		If s0.Contains("亥") AndAlso Not s0.Contains("子") AndAlso s0.Contains("丑") Then s += "亥丑j子" + vbNL
		If s0.Contains("子") AndAlso Not s0.Contains("丑") AndAlso s0.Contains("寅") Then s += "子寅j丑" + vbNL
		If s0.Contains("丑") AndAlso Not s0.Contains("寅") AndAlso s0.Contains("卯") Then s += "丑卯j寅" + vbNL
		If s0.Contains("寅") AndAlso Not s0.Contains("卯") AndAlso s0.Contains("辰") Then s += "寅辰j卯" + vbNL
		If s0.Contains("卯") AndAlso Not s0.Contains("辰") AndAlso s0.Contains("巳") Then s += "卯巳j辰" + vbNL
		If s0.Contains("辰") AndAlso Not s0.Contains("巳") AndAlso s0.Contains("午") Then s += "辰午j巳" + vbNL
		If s0.Contains("巳") AndAlso Not s0.Contains("午") AndAlso s0.Contains("未") Then s += "巳未j午" + vbNL
		If s0.Contains("午") AndAlso Not s0.Contains("未") AndAlso s0.Contains("申") Then s += "午申j未" + vbNL
		If s0.Contains("未") AndAlso Not s0.Contains("申") AndAlso s0.Contains("酉") Then s += "未酉j申" + vbNL
		If s0.Contains("申") AndAlso Not s0.Contains("酉") AndAlso s0.Contains("戌") Then s += "申戌j酉" + vbNL
		If s0.Contains("酉") AndAlso Not s0.Contains("戌") AndAlso s0.Contains("亥") Then s += "酉亥j戌" + vbNL
		If s0.Contains("戌") AndAlso Not s0.Contains("亥") AndAlso s0.Contains("子") Then s += "戌子j亥" + vbNL
		'---------------------------------------------
		Return s
	End Function
	Public Function Uf八字組合_五行(s0 As String) As String
		Dim s = "--------------" + vbNL ' s0=八字
		If s0.Contains("甲") AndAlso s0.Contains("庚") Then s += "甲^庚" + vbNL
		If s0.Contains("乙") AndAlso s0.Contains("辛") Then s += "乙^辛" + vbNL
		If s0.Contains("丙") AndAlso s0.Contains("壬") Then s += "丙^壬" + vbNL
		If s0.Contains("丁") AndAlso s0.Contains("癸") Then s += "丁^癸" + vbNL
		' 合制
		If s0.Contains("甲") AndAlso s0.Contains("己") Then s += "甲z己->土" + vbNL
		If s0.Contains("乙") AndAlso s0.Contains("庚") Then s += "乙z庚->金" + vbNL
		If s0.Contains("丙") AndAlso s0.Contains("辛") Then s += "丙z辛->水" + vbNL
		If s0.Contains("丁") AndAlso s0.Contains("壬") Then s += "丁z壬->木" + vbNL
		If s0.Contains("戊") AndAlso s0.Contains("癸") Then s += "戊z癸->火" + vbNL
		' 天干夾
		If s0.Contains("甲") AndAlso Not s0.Contains("乙") AndAlso s0.Contains("丙") Then s += "甲丙g乙" + vbNL
		If s0.Contains("乙") AndAlso Not s0.Contains("丙") AndAlso s0.Contains("丁") Then s += "乙丁g丙" + vbNL
		If s0.Contains("丙") AndAlso Not s0.Contains("丁") AndAlso s0.Contains("戊") Then s += "丙戊g丁" + vbNL
		If s0.Contains("丁") AndAlso Not s0.Contains("戊") AndAlso s0.Contains("己") Then s += "丁己g戊" + vbNL
		If s0.Contains("戊") AndAlso Not s0.Contains("己") AndAlso s0.Contains("庚") Then s += "戊庚g己" + vbNL
		If s0.Contains("己") AndAlso Not s0.Contains("庚") AndAlso s0.Contains("辛") Then s += "己辛g庚" + vbNL
		If s0.Contains("庚") AndAlso Not s0.Contains("辛") AndAlso s0.Contains("壬") Then s += "庚壬g辛" + vbNL
		If s0.Contains("辛") AndAlso Not s0.Contains("壬") AndAlso s0.Contains("癸") Then s += "辛癸g壬" + vbNL
		If s0.Contains("壬") AndAlso Not s0.Contains("癸") AndAlso s0.Contains("甲") Then s += "壬甲g癸" + vbNL
		If s0.Contains("癸") AndAlso Not s0.Contains("甲") AndAlso s0.Contains("乙") Then s += "癸乙g甲" + vbNL
		Return s
	End Function
	Public Function Uf八字組合_六行(s0 As String) As String
		Dim s = "--------------" + vbNL ' s0=八字
		If s0.Contains("甲") AndAlso s0.Contains("庚") Then s += "甲^庚" + vbNL
		If s0.Contains("乙") AndAlso s0.Contains("辛") Then s += "乙^辛" + vbNL
		If s0.Contains("丙") AndAlso s0.Contains("壬") Then s += "丙^壬" + vbNL
		If s0.Contains("丁") AndAlso s0.Contains("癸") Then s += "丁^癸" + vbNL
		If s0.Contains("戊") AndAlso s0.Contains("己") Then s += "戊^己" + vbNL
		' 合制
		If False Then ' 早期 卯[卯辰] 推出的公式
			If s0.Contains("甲") AndAlso s0.Contains("辛") Then s += "甲z辛=火" + vbNL
			If s0.Contains("乙") AndAlso s0.Contains("庚") Then s += "乙z庚=水" + vbNL
			If s0.Contains("丙") AndAlso s0.Contains("癸") Then s += "丙z癸=金" + vbNL
			If s0.Contains("丁") AndAlso s0.Contains("壬") Then s += "丁z壬=木" + vbNL
			If s0.Contains("戊") AndAlso s0.Contains("乙") Then s += "戊z乙=木" + vbNL
			If s0.Contains("己") AndAlso s0.Contains("辛") Then s += "己z辛=金" + vbNL
		Else ' 最新 卯[卯未]
			If s0.Contains("甲") AndAlso s0.Contains("辛") Then s += "甲z辛=0" + vbNL ' 己
			If s0.Contains("乙") AndAlso s0.Contains("庚") Then s += "乙z庚=0" + vbNL ' 戊
			If s0.Contains("丙") AndAlso s0.Contains("癸") Then s += "丙z癸=0" + vbNL
			If s0.Contains("丁") AndAlso s0.Contains("壬") Then s += "丁z壬=0" + vbNL

			'If s0.Contains("戊") AndAlso s0.Contains("庚") Then s += "戊z庚=金" + vbNL ' 只取 阳z阴
			'If s0.Contains("戊") AndAlso s0.Contains("辛") Then s += "戊z辛=金" + vbNL
			'If s0.Contains("己") AndAlso s0.Contains("甲") Then s += "己z甲=木" + vbNL ' 只取 阳z阴
			'If s0.Contains("己") AndAlso s0.Contains("乙") Then s += "己z乙=木" + vbNL
		End If

		' 天干夾
		If s0.Contains("甲") AndAlso Not s0.Contains("乙") AndAlso s0.Contains("丙") Then s += "甲丙j乙" + vbNL
		If s0.Contains("乙") AndAlso Not s0.Contains("丙") AndAlso s0.Contains("丁") Then s += "乙丁j丙" + vbNL
		If s0.Contains("丙") AndAlso Not s0.Contains("丁") AndAlso s0.Contains("戊") Then s += "丙戊j丁" + vbNL
		If s0.Contains("丁") AndAlso Not s0.Contains("戊") AndAlso s0.Contains("庚") Then s += "丁庚j戊" + vbNL
		If s0.Contains("戊") AndAlso Not s0.Contains("庚") AndAlso s0.Contains("辛") Then s += "戊辛j庚" + vbNL
		If s0.Contains("庚") AndAlso Not s0.Contains("辛") AndAlso s0.Contains("壬") Then s += "庚壬j辛" + vbNL
		If s0.Contains("辛") AndAlso Not s0.Contains("壬") AndAlso s0.Contains("癸") Then s += "辛癸j壬" + vbNL
		If s0.Contains("壬") AndAlso Not s0.Contains("癸") AndAlso s0.Contains("己") Then s += "壬己j癸" + vbNL
		If s0.Contains("癸") AndAlso Not s0.Contains("己") AndAlso s0.Contains("甲") Then s += "癸甲j己" + vbNL
		If s0.Contains("己") AndAlso Not s0.Contains("甲") AndAlso s0.Contains("乙") Then s += "己乙j甲" + vbNL
		' az
		If s0.Contains("卯") AndAlso s0.Contains("申") Then s += "卯az申=水+己->己" + vbNL
		If s0.Contains("酉") AndAlso s0.Contains("寅") Then s += "酉az寅=火+戊->戊" + vbNL
		If s0.Contains("午") AndAlso s0.Contains("亥") Then s += "午az亥=己+木->木" + vbNL
		If s0.Contains("子") AndAlso s0.Contains("巳") Then s += "子az巳=戊+金->金" + vbNL
		'墓az下一支
		If s0.Contains("丑") AndAlso s0.Contains("寅") Then s += "丑az寅=0" + vbNL ' 己
		If s0.Contains("辰") AndAlso s0.Contains("巳") Then s += "辰az巳=0" + vbNL
		If s0.Contains("未") AndAlso s0.Contains("申") Then s += "未az申=0" + vbNL ' 戊
		If s0.Contains("戌") AndAlso s0.Contains("亥") Then s += "戌az亥=0" + vbNL
		' 四驛馬
		If False Then ' 配置4a
			If s0.Contains("亥") AndAlso s0.Contains("寅") Then s += "亥v寅=己s木->木" + vbNL
			If s0.Contains("巳") AndAlso s0.Contains("申") Then s += "巳v申=戊s金->金" + vbNL
			If s0.Contains("寅") AndAlso s0.Contains("巳") Then s += "寅t巳=火" + vbNL
			If s0.Contains("申") AndAlso s0.Contains("亥") Then s += "申t亥=水" + vbNL
		Else ' 配置5
			If s0.Contains("寅") AndAlso s0.Contains("巳") Then s += "寅t巳=火s戊->戊" + vbNL ' 长生方
			If s0.Contains("巳") AndAlso s0.Contains("申") Then s += "巳v申=金" + vbNL
			If s0.Contains("申") AndAlso s0.Contains("亥") Then s += "申t亥=水s己->己" + vbNL
			If s0.Contains("亥") AndAlso s0.Contains("寅") Then s += "亥v寅=木" + vbNL
		End If

		' 寅z亥->木  巳z申->水
		' 子v丑->土  卯z戌->火   古法，比较用
		' 辰z酉->金  午z未->土
		' 六合
		If s0.Contains("卯") AndAlso s0.Contains("戌") Then s += "卯v戌=火s戊->戊" + vbNL
		If s0.Contains("辰") AndAlso s0.Contains("酉") Then s += "辰v酉=水s己->己" + vbNL
		If s0.Contains("子") AndAlso s0.Contains("丑") Then s += "子v丑=金s水2s己->己" + vbNL
		If s0.Contains("午") AndAlso s0.Contains("未") Then s += "午v未=木s火2+戊->戊" + vbNL
		' 六穿
		If s0.Contains("子") AndAlso s0.Contains("未") Then s += "子t未=木k戊->？" + vbNL
		If s0.Contains("午") AndAlso s0.Contains("丑") Then s += "午t丑=金k己->？" + vbNL
		If s0.Contains("卯") AndAlso s0.Contains("辰") Then s += "卯t辰=水s己s木2->木" + vbNL
		If s0.Contains("酉") AndAlso s0.Contains("戌") Then s += "酉t戌=火s戊s金2->金" + vbNL
		' 三會
		'If s0.Contains("寅") AndAlso s0.Contains("卯") AndAlso s0.Contains("辰") Then s += "寅(卯)辰->木" + vbNL
		'If s0.Contains("巳") AndAlso s0.Contains("午") AndAlso s0.Contains("未") Then s += "巳(午)未->火" + vbNL
		'If s0.Contains("申") AndAlso s0.Contains("酉") AndAlso s0.Contains("戌") Then s += "申(酉)戌->金" + vbNL
		'If s0.Contains("亥") AndAlso s0.Contains("子") AndAlso s0.Contains("丑") Then s += "亥(子)丑->子" + vbNL
		' 土
		If s0.Contains("辰") AndAlso s0.Contains("未") Then s += "辰+未=木" + vbNL
		If s0.Contains("戌") AndAlso s0.Contains("丑") Then s += "戌+丑=金" + vbNL
		If s0.Contains("丑") AndAlso s0.Contains("辰") Then s += "丑+辰=水s己->己" + vbNL
		If s0.Contains("未") AndAlso s0.Contains("戌") Then s += "未+戌=火s戊->戊" + vbNL
		' 特殊三合
		If s0.Contains("甲") AndAlso s0.Contains("丁") AndAlso s0.Contains("戊") Then s += "●甲丁+戊=寅午+戌" + vbNL ' 甲s丁s戊
		If s0.Contains("庚") AndAlso s0.Contains("癸") AndAlso s0.Contains("己") Then s += "●庚癸+己=申子+辰" + vbNL ' 庚s癸s己
		If s0.Contains("丙") AndAlso s0.Contains("辛") AndAlso s0.Contains("己") Then s += "●丙辛+己=巳酉+丑" + vbNL ' 丙k辛k己
		If s0.Contains("壬") AndAlso s0.Contains("乙") AndAlso s0.Contains("戊") Then s += "●壬乙+戊=亥卯+未" + vbNL ' 壬k乙k戊

		If Not s0.Contains("甲") AndAlso s0.Contains("丁") AndAlso s0.Contains("戊") Then s += "●丁戊g甲=午戌g寅" + vbTab
		If s0.Contains("甲") AndAlso Not s0.Contains("丁") AndAlso s0.Contains("戊") Then s += "●甲戊g丁=寅戌g午" + vbTab
		If s0.Contains("甲") AndAlso s0.Contains("丁") AndAlso Not s0.Contains("戊") Then s += "●甲丁g戊=寅午g戌" + vbTab

		If Not s0.Contains("庚") AndAlso s0.Contains("癸") AndAlso s0.Contains("己") Then s += "●癸己g庚=子辰g申" + vbTab
		If s0.Contains("庚") AndAlso Not s0.Contains("癸") AndAlso s0.Contains("己") Then s += "●庚己g癸=申辰g子" + vbTab
		If s0.Contains("庚") AndAlso s0.Contains("癸") AndAlso Not s0.Contains("己") Then s += "●庚癸g己=申子g辰" + vbTab

		If Not s0.Contains("丙") AndAlso s0.Contains("辛") AndAlso s0.Contains("己") Then s += "●辛己g丙=酉丑g巳" + vbTab
		If s0.Contains("丙") AndAlso Not s0.Contains("辛") AndAlso s0.Contains("己") Then s += "●丙己g辛=巳丑g酉" + vbTab
		If s0.Contains("丙") AndAlso s0.Contains("辛") AndAlso Not s0.Contains("己") Then s += "●丙辛g己=巳酉g丑" + vbTab

		If Not s0.Contains("壬") AndAlso s0.Contains("乙") AndAlso s0.Contains("戊") Then s += "●乙戊g壬=卯未g亥" + vbTab
		If s0.Contains("壬") AndAlso Not s0.Contains("乙") AndAlso s0.Contains("戊") Then s += "●壬戊g乙=亥未g卯" + vbTab
		If s0.Contains("壬") AndAlso s0.Contains("乙") AndAlso Not s0.Contains("戊") Then s += "●壬乙g戊=亥卯g未" + vbTab
		Return s
	End Function
End Module
