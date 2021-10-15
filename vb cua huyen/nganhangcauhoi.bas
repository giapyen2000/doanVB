Attribute VB_Name = "Module1"
Sub nganhangcauhoi()
HB(1) = "Leänh java coù taùc duïng gì?"
TB(1, 1) = "Bieân dòch file .java thaønh file .class"
TB(1, 2) = "Chaïy chöông trình java coù ñuoâi .class"
TB(1, 3) = "Bieân dòch file .class thaønh .java"
TB(1, 4) = "Duøng leänh debug chöông trình"
TB(1, 5) = "Khoâng coù leänh naøy"
DA(1) = 1

HB(2) = "Leänh khai bao :Scanner in = new Scanner(System.in);.coù taùc duïng gì?"
TB(2, 1) = "Khai baùo ñoáiùi töôïng teân in, keát noái vôùi stream input chuaån"
TB(2, 2) = "Khai baùo ñoái töôïng Scanner ñeå in döõ lieäu"
TB(2, 3) = "Khai baùo ñoái töôïng in, keát noái vôùi stream output chuaån"
TB(2, 4) = "Khai baùo ñoái töôïng Scanner keát noái tôùi döõ lieäu"
TB(2, 5) = "Khoâng coù leänh treân"
DA(2) = 1

HB(3) = "Phöông thöùc nextLine() trong lôùp Scanner coù taùc duïng gì?"
TB(3, 1) = "ñoïc moät chuoãi kí töï töø baøn phím, keå caû daáu caùch"
TB(3, 2) = "ñoïc moät chuoãi kí töï töø baøn phím, khoâng keå daáu caùch"
TB(3, 3) = "ñoïc moät giaù trò soá nguyeân töø baøn phím"
TB(3, 4) = "ñoïc moät giaù trò soá thöïc töø baøn phím"
TB(3, 5) = "lôùp Scanner khoâng coù phöông thöùc naøy"
DA(3) = 1

HB(4) = "Phöông thöùc next() trong lôùp Scanner coù taùc duïng gì?"
TB(4, 1) = "ñoïc moät chuoãi kí töï töø baøn phím, keå caû daáu caùch"
TB(4, 2) = "ñoïc moät chuoãi kí töï töø baøn phím, khoâng keå daáu caùch"
TB(4, 3) = "ñoïc moät giaù trò soá nguyeân töø baøn phím"
TB(4, 4) = "ñoïc moät giaù trò soá thöïc töø baøn phím"
TB(4, 5) = "lôùp Scanner khoâng coù phöông thöùc naøy"
DA(4) = 1


HB(5) = "Phöông thöùc nextInt() trong lôùp Scanner  coù taùc duïng gì?"
TB(5, 1) = "ñoïc moät chuoãi kí töï töø baøn phím, keå caû daáu caùch"
TB(5, 2) = "ñoïc moät chuoãi kí töï töø baøn phím, khoâng keå daáu caùch"
TB(5, 3) = "ñoïc moät giaù trò soá nguyeân töø baøn phím"
TB(5, 4) = "ñoïc moät giaù trò soá thöïc töø baøn phím"
TB(5, 5) = "lôùp Scanner khoâng coù phöông thöùc naøy"
DA(5) = 1

HB(6) = "Phöông thöùc hasNext() trong lôùp Scanner coù taùc duïng gì?"
TB(6, 1) = "kieåm tra xem 1 chuoãi kí töï naèm trong input khoâng,keå caû daáu caùch"
TB(6, 2) = "kieåm tra xem moät chuoãi kí töï coù naèm trong input khoâng,khoâng keå daáu caùch"
TB(6, 3) = "kieåm tra xem moät soá nguyeân coù naèm trong input khoâng"
TB(6, 4) = "Kieåm tra xem moät soá thöïc coù trong input khoâng"
TB(6, 5) = "Lôùp Scanner khoâng coù phöông thöùc naøy"
DA(6) = 1

HB(7) = "Phöông thöùc hasNextLine() trong lôùp Scanner coù taùc duïng gì?"
TB(7, 1) = "kieåm tra xem 1 chuoãi kí töï naèm trong input khoâng,keå caû daáu caùch"
TB(7, 2) = "kieåm tra xem moät chuoãi kí töï coù naèm trong input khoâng,khoâng keå daáu caùch"
TB(7, 3) = "kieåm tra xem moät soá nguyeân coù naèm trong input khoâng"
TB(7, 4) = "Kieåm tra xem moät soá thöïc coù trong input khoâng"
TB(7, 5) = "Lôùp Scanner khoâng coù phöông thöùc naøy"
DA(7) = 1

HB(8) = "Phöông thöùc hasNextInt() trong lôùp Scanner coù taùc duïng gì?"
TB(8, 1) = "kieåm tra xem 1 chuoãi kí töï naèm trong input khoâng,keå caû daáu caùch"
TB(8, 2) = "kieåm tra xem moät chuoãi kí töï coù naèm trong input khoâng,khoâng keå daáu caùch"
TB(8, 3) = "kieåm tra xem moät soá nguyeân coù naèm trong input khoâng"
TB(8, 4) = "Kieåm tra xem moät soá thöïc coù trong input khoâng"
TB(8, 5) = "Lôùp Scanner khoâng coù phöông thöùc naøy"
DA(8) = 1

HB(9) = "Cho khoái leänh sau:if (yourSale >= target) bonus = 100 + 0.01 * (yourSale - target);else  & vbCrLf & bonus = 0 víi yourSale = 3000, target = 2500, vËy bonus = ?"
TB(9, 1) = "105"
TB(9, 2) = "150"
TB(9, 3) = "600"
TB(9, 4) = "0"
TB(9, 5) = "leänh sai"
DA(9) = 1

HB(10) = "Cho khoái leänh sau:if (yourSale >= target) bonus = 100 + 0.01 * (yourSale - target);else  & vbCrLf & bonus = 0 víi yourSale = 3000, target = 2500, vËy bonus = ?"
TB(10, 1) = "105"
TB(10, 2) = "150"
TB(10, 3) = "600"
TB(10, 4) = "0"
TB(10, 5) = "Leänh sai"
DA(10) = 1

HB(11) = "Cho khoái leänh  sau:if (yourSale >= target) bonus = 100 + 0.01 * (yourSale - target);else  & vbCrLf & bonus = 0 víi yourSale = 3000, target = 2500, vËy bonus = ?"
TB(11, 1) = "20.015"
TB(11, 2) = "15"
TB(11, 3) = "15.015"
TB(11, 4) = "20.035015"
TB(11, 5) = "Leänh sai"
DA(11) = 1

HB(12) = "Cho khoái leänh  sau:if (yourSale >= target) bonus = 100 + 0.01 * (yourSale - target);else  & vbCrLf & bonus = 0 víi yourSale = 3000, target = 2500, vËy bonus = ?"
TB(12, 1) = "20.015"
TB(12, 2) = "15"
TB(12, 3) = "15.015"
TB(12, 4) = "20.035015"
TB(12, 5) = "Leänh sai"
DA(12) = 1

HB(13) = "Cho khoái leänh sau:int years = 0;do {balance += payment;double interest = balance * 0.1 / 100;balance += interestyears ++;} while (balance < goal);Cho balance = 20, goal = 20, payment = 5.Sau khi chaïy khoái leänh, balance = ?"
TB(13, 1) = "20.015"
TB(13, 2) = "15"
TB(13, 3) = "15.015"
TB(13, 4) = "20.035015"
TB(13, 5) = "leänh sai"
DA(13) = 1

HB(14) = "Cho khoái leänh sau:int years = 0;do {balance += payment;double interest = balance * 0.1 / 100;balance += interestyears ++;} while (balance < goal);Cho balance = 20, goal = 20, payment = 5. Sau khi chaïy khoái leänh, years = ?"
TB(14, 1) = "20.015"
TB(14, 2) = "1"
TB(14, 3) = "15.015"
TB(14, 4) = "2"
TB(4, 5) = "leänh sai"
DA(14) = 1

HB(15) = "con meøo keâu theá naøo?"
TB(15, 1) = "meo meo"
TB(15, 2) = "gaâu gaâu"
TB(15, 3) = "eùc eùc"
TB(15, 4) = "quaëc quaëc"
TB(15, 5) = "taát caû ñeàu sai"
DA(15) = 1

HB(16) = "con choù keâu theá naøo?"
TB(16, 1) = "meo meo"
TB(16, 2) = "gaâu gaâu"
TB(16, 3) = "eùc eùc"
TB(16, 4) = "quaëc quaëc"
TB(16, 5) = "taát caû ñeàu sai"
DA(16) = 1

HB(17) = "con gaø keâu theá naøo?"
TB(17, 1) = "meo meo"
TB(17, 2) = "gaâu gaâu"
TB(17, 3) = "eùc eùc"
TB(17, 4) = "quaëc quaëc"
TB(17, 5) = "taát caû ñeàu sai"
DA(17) = 1

HB(18) = "con gaáu keâu theá naøo?"
TB(18, 1) = "meo meo"
TB(18, 2) = "gaâu gaâu"
TB(18, 3) = "eùc eùc"
TB(18, 4) = "quaëc quaëc"
TB(18, 5) = "taát caû ñeàu sai"
DA(18) = 1

HB(19) = "con chim tu huù keâu theá naøo?"
TB(19, 1) = "meo meo"
TB(19, 2) = "gaâu gaâu"
TB(19, 3) = "eùc eùc"
TB(19, 4) = "quaëc quaëc"
TB(19, 5) = "taát caû ñeàu sai"
DA(19) = 1

HB(20) = "con cuù keâu theá naøo?"
TB(20, 1) = "meo meo"
TB(20, 2) = "gaâu gaâu"
TB(20, 3) = "eùc eùc"
TB(20, 4) = "quaëc quaëc"
TB(20, 5) = "taát caû ñeàu sai"
DA(20) = 1

HB(21) = "con lôïn keâu theá naøo?"
TB(21, 1) = "meo meo"
TB(21, 2) = "gaâu gaâu"
TB(21, 3) = "eùc eùc"
TB(21, 4) = "quaëc quaëc"
TB(21, 5) = "taát caû ñeàu sai"
DA(21) = 1

HB(22) = "con vòt keâu theá naøo?"
TB(22, 1) = "meo meo"
TB(22, 2) = "gaâu gaâu"
TB(22, 3) = "eùc eùc"
TB(22, 4) = "quaëc quaëc"
TB(22, 5) = "taát caû ñeàu sai"
DA(22) = 1

HB(23) = "traùi ñaát coù hình gì?"
TB(23, 1) = "hình vuoâng"
TB(23, 2) = "hình troøn"
TB(23, 3) = "hình caàu"
TB(23, 4) = "hình elip"
TB(23, 5) = "khoâng coù ñaùp aùn ñuùng"
DA(23) = 1

HB(24) = "moät naêm coù bao nhieâu thaùng?"
TB(24, 1) = "12"
TB(24, 2) = "13"
TB(24, 3) = "14"
TB(24, 4) = "15"
TB(24, 5) = "khoâng coù ñaùp aùn ñuùng"
DA(24) = 1

HB(25) = "Java platform goàm maáy phaàn?"
TB(25, 1) = "1"
TB(25, 2) = "2"
TB(25, 3) = "3"
TB(25, 4) = "4"
TB(25, 5) = "khoâng coù ñaùp aùn ñuùng"
DA(25) = 1

HB(26) = "coù bao nhieâu caùch vieát chuù thích trong java?"
TB(26, 1) = "1"
TB(26, 2) = "2"
TB(26, 3) = "3"
TB(26, 4) = "4"
TB(26, 5) = "5"
DA(26) = 1

HB(27) = "khai baùo naøo döùoi ñaây laø ñuùng?"
TB(27, 1) = "public class default {}"
TB(27, 2) = "protected inner class engine {}"
TB(27, 3) = "final class outer {}"
TB(27, 4) = "taát caû ñeàu ñuùng"
TB(27, 5) = "taát caû ñeàu sai"
DA(27) = 1

HB(28) = "moät lôùp trong java coù bao nhieàu lôùp cha?"
TB(28, 1) = "1"
TB(28, 2) = "2"
TB(28, 3) = "3"
TB(28, 4) = "4"
TB(28, 5) = "5"
DA(28) = 1

HB(29) = "moät lôùp trong java coù bao nhieâu lôùp con?"
TB(29, 1) = "1"
TB(29, 2) = "2"
TB(29, 3) = "3"
TB(29, 4) = "4"
TB(29, 5) = "voâ soá"
DA(29) = 1

HB(30) = "moät chöông trình goàm 2 class seõ coù ba nhieâu phöông thöùc main?"
TB(30, 1) = "1"
TB(30, 2) = "2"
TB(30, 3) = "3"
TB(30, 4) = "4"
TB(30, 5) = "5"
DA(30) = 1

HB(31) = "coù bao nhieâu loaïi bieán trong java?"
TB(31, 1) = "1"
TB(31, 2) = "2"
TB(31, 3) = "3"
TB(31, 4) = "4"
TB(31, 5) = "5"
DA(31) = 1

HB(32) = "tröôøng döõ lieäu laø caùc bieán daïng naøo sau ñaây?"
TB(32, 1) = "bieán ñaïi dieän vaø tham bieán"
TB(32, 2) = "bieán ñaïi dieän vaø bieán lôùp"
TB(32, 3) = "bieán ñaïi dieän vaø bieán cuïc boä"
TB(32, 4) = "bieán lôùp vaø tham soá"
TB(32, 5) = "khoâng coù ñaùp aùn ñuùng"
DA(32) = 1


HB(33) = "coù bao nhieâu kieåu döõ lieäu cô sôû trong java?"
TB(33, 1) = "7"
TB(33, 2) = "8"
TB(33, 3) = "9"
TB(33, 4) = "10"
TB(33, 5) = "11"
DA(33) = 1

HB(34) = "coù bao nhieâu kieåu soá nguyeân trong java?"
TB(34, 1) = "1"
TB(34, 2) = "2"
TB(34, 3) = "3"
TB(34, 4) = "4"
TB(34, 5) = "5"
DA(34) = 1

HB(35) = "hình thang coù maáy caïch?"
TB(35, 1) = "1"
TB(35, 2) = "2"
TB(35, 3) = "3"
TB(35, 4) = "4"
TB(35, 5) = "5"
DA(35) = 1

HB(36) = "hình vuoâng coù maáy caïch?"
TB(36, 1) = "2"
TB(36, 2) = "3"
TB(36, 3) = "4"
TB(36, 4) = "5"
TB(36, 5) = "6"
DA(36) = 1

HB(37) = "caùi caây thöôøng coù maøu gì?"
TB(37, 1) = "maøu xanh"
TB(37, 2) = "maøu ñen"
TB(37, 3) = "maøu traéng"
TB(37, 4) = "trong suoát"
TB(37, 5) = "maøu hoàng"
DA(37) = 1

HB(38) = "hình tam giaùc coù maáy caïnh?"
TB(38, 1) = "1"
TB(38, 2) = "2"
TB(38, 3) = "3"
TB(38, 4) = "4"
TB(38, 5) = "5"
DA(38) = 1

HB(39) = "con choù coù maùy caùi chaân?"
TB(39, 1) = "1"
TB(39, 2) = "2"
TB(39, 3) = "3"
TB(39, 4) = "4"
TB(39, 5) = "5"
DA(39) = 1

HB(40) = "quaû tröùng coù hình gì?"
TB(40, 1) = "hình vuoâng"
TB(40, 2) = "hình troøn"
TB(40, 3) = "hình elip"
TB(40, 4) = "hình baàu duïc"
TB(40, 5) = "hình tam giaùc"
DA(40) = 1

HB(41) = "trong keá thöøa phöông thöùc cuûa lôùp con ñöôïc khai baùo gioáng phöông thöùc cuûa lôùp cha veà caû teân laãn tham soá?"
TB(41, 1) = "Overload"
TB(41, 2) = "Override"
TB(41, 3) = "synchronized"
TB(41, 4) = "Serializable"
TB(41, 5) = "taát caû ñeàu sai"
DA(41) = 1

HB(42) = "trong class caùc phöông thöùc truøng teân khaùc tham soá laø?"
TB(42, 1) = "Overload"
TB(42, 2) = " Override"
TB(42, 3) = "synchronized"
TB(42, 4) = "Serializable"
TB(42, 5) = "taát caû ñeàu sai"
DA(42) = 1

HB(43) = "Cho String str = univerity, leänh naøo döôùi ñaây laáy chuoãi univer vaø gaùn vaøo chuoãi  str1??"
TB(43, 1) = "String str1 = str.substring(0, 6);"
TB(43, 2) = "String str1 = str.substring(0, 5);"
TB(43, 3) = "String str1 = str.substring(1, 6);"
TB(43, 4) = ". String str1 = str.substring(5);"
TB(43, 5) = "taát caû ñeàu sai"
DA(43) = 1

HB(44) = "ñònh nghó interface naøo sau ñaây laø khoâng hôïp leä?"
TB(44, 1) = "public interface inout {}"
TB(44, 2) = "protected interface inout { int i = 12;}"
TB(44, 3) = ". interface inout { public final int MAX_INDEX = 100;}"
TB(44, 4) = "interface input { public void indl();}"
TB(44, 5) = "taât caû ñeàu hôïp leä"
DA(44) = 1

HB(45) = "leänh charAt(n) coù taùc duïng gì?"
TB(45, 1) = "tieàm kieám kí töï thöù n"
TB(45, 2) = "traû veà kí töï thöù n-1"
TB(45, 3) = "traû veà kí töï thöù n"
TB(45, 4) = "traû veà kí töï coù vò trí chæ muïc n"
TB(45, 5) = "taát caû ñeàu sai"
DA(45) = 1

HB(46) = "ñeå kieåm tra 2 chuoãi coù baèng nhau khoâng ta söû duïng phöông thöùc naøo?"
TB(46, 1) = "string1== string2"
TB(46, 2) = "string1 = string2"
TB(46, 3) = "string1.equals(string2)"
TB(46, 4) = ". string1.equal(string2)"
TB(46, 5) = "taât caû ñeàu sai"
DA(46) = 1

HB(47) = "ñeå ñaûo gaù trò cuûa 1 bieán boolean,ta duøng toaùn töû naøo?"
TB(47, 1) = "!"
TB(47, 2) = ">>"
TB(47, 3) = "<<"
TB(47, 4) = "=="
TB(47, 5) = "taát caû ñeàu sai"
DA(47) = 1

HB(48) = "leänh naøo ngöøng voøng laëp hieän thôøi vaø baét ñaàu voøng laëp tieáp theo?"
TB(48, 1) = "continue"
TB(48, 2) = "break"
TB(48, 3) = ". cease"
TB(48, 4) = "end"
TB(48, 5) = "taát caû ñeàu sai"
DA(48) = 1

HB(49) = "kieåu enum laø gì?"
TB(49, 1) = "laø kieåu döõ lieäp goàm caùc tröôøng chöùa moät taäp coá ñònh caùc haèng soá"
TB(49, 2) = "laø kieåu söõ lieäu thoáng keâ caùc bieán soá"
TB(49, 3) = "laø kieåu döõ lieäu trong java"
TB(49, 4) = "taát caû ñeàu ñuùng"
TB(49, 5) = "taát caû ñeàu sai"
DA(49) = 1

HB(50) = "coù bao nhieâu loaïi quyeàn truy caäp trong jav?"
TB(50, 1) = "1"
TB(50, 2) = "2"
TB(50, 3) = "3"
TB(50, 4) = "4"
TB(50, 5) = "5"
DA(50) = 1

End Sub

Sub daocauhoi()
Call nganhangcauhoi

    Randomize
    CH(1) = Int(Rnd() * 50 + 1)

    For I = 2 To 50
    Kt = False
        Do While Kt = False
            Randomize
            n = Int(Rnd() * 50 + 1)
            Kt = True
            For j = 1 To I - 1
                If n = CH(j) Then
                    Kt = False
                End If
            Next
        Loop
        CH(I) = n
    Next
    

 For nI = 1 To 50
        HBN(nI) = HB(CH(nI))
        dung = DA(CH(nI))
    Randomize
    TL(1) = Int(Rnd() * 5 + 1)
    For d = 2 To 5
            Kt = False
            Do While Kt = False
                Randomize
                n1 = Int(Rnd() * 5 + 1)
                Kt = True
                    For j = 1 To d - 1
                        If n1 = TL(j) Then
                            Kt = False
                        End If
                    Next
            Loop
        TL(d) = n1
    Next

    For M = 1 To 5
        
        If TB(CH(nI), TL(dung)) = TB(CH(nI), TL(M)) Then
            DAN(nI) = TL(M)
        End If
           TBN(nI, M) = TB(CH(nI), TL(M))
        Next
     
        
    Next
    

End Sub
