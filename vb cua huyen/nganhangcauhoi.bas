Attribute VB_Name = "Module1"
Sub nganhangcauhoi()
HB(1) = "Le�nh java co� ta�c du�ng g�?"
TB(1, 1) = "Bie�n d�ch file .java tha�nh file .class"
TB(1, 2) = "Cha�y ch��ng tr�nh java co� �uo�i .class"
TB(1, 3) = "Bie�n d�ch file .class tha�nh .java"
TB(1, 4) = "Du�ng le�nh debug ch��ng tr�nh"
TB(1, 5) = "Kho�ng co� le�nh na�y"
DA(1) = 1

HB(2) = "Le�nh khai bao :Scanner in = new Scanner(System.in);.co� ta�c du�ng g�?"
TB(2, 1) = "Khai ba�o �o�i�i t���ng te�n in, ke�t no�i v��i stream input chua�n"
TB(2, 2) = "Khai ba�o �o�i t���ng Scanner �e� in d�� lie�u"
TB(2, 3) = "Khai ba�o �o�i t���ng in, ke�t no�i v��i stream output chua�n"
TB(2, 4) = "Khai ba�o �o�i t���ng Scanner ke�t no�i t��i d�� lie�u"
TB(2, 5) = "Kho�ng co� le�nh tre�n"
DA(2) = 1

HB(3) = "Ph��ng th��c nextLine() trong l��p Scanner co� ta�c du�ng g�?"
TB(3, 1) = "�o�c mo�t chuo�i k� t�� t�� ba�n ph�m, ke� ca� da�u ca�ch"
TB(3, 2) = "�o�c mo�t chuo�i k� t�� t�� ba�n ph�m, kho�ng ke� da�u ca�ch"
TB(3, 3) = "�o�c mo�t gia� tr� so� nguye�n t�� ba�n ph�m"
TB(3, 4) = "�o�c mo�t gia� tr� so� th��c t�� ba�n ph�m"
TB(3, 5) = "l��p Scanner kho�ng co� ph��ng th��c na�y"
DA(3) = 1

HB(4) = "Ph��ng th��c next() trong l��p Scanner co� ta�c du�ng g�?"
TB(4, 1) = "�o�c mo�t chuo�i k� t�� t�� ba�n ph�m, ke� ca� da�u ca�ch"
TB(4, 2) = "�o�c mo�t chuo�i k� t�� t�� ba�n ph�m, kho�ng ke� da�u ca�ch"
TB(4, 3) = "�o�c mo�t gia� tr� so� nguye�n t�� ba�n ph�m"
TB(4, 4) = "�o�c mo�t gia� tr� so� th��c t�� ba�n ph�m"
TB(4, 5) = "l��p Scanner kho�ng co� ph��ng th��c na�y"
DA(4) = 1


HB(5) = "Ph��ng th��c nextInt() trong l��p Scanner  co� ta�c du�ng g�?"
TB(5, 1) = "�o�c mo�t chuo�i k� t�� t�� ba�n ph�m, ke� ca� da�u ca�ch"
TB(5, 2) = "�o�c mo�t chuo�i k� t�� t�� ba�n ph�m, kho�ng ke� da�u ca�ch"
TB(5, 3) = "�o�c mo�t gia� tr� so� nguye�n t�� ba�n ph�m"
TB(5, 4) = "�o�c mo�t gia� tr� so� th��c t�� ba�n ph�m"
TB(5, 5) = "l��p Scanner kho�ng co� ph��ng th��c na�y"
DA(5) = 1

HB(6) = "Ph��ng th��c hasNext() trong l��p Scanner co� ta�c du�ng g�?"
TB(6, 1) = "kie�m tra xem 1 chuo�i k� t�� na�m trong input kho�ng,ke� ca� da�u ca�ch"
TB(6, 2) = "kie�m tra xem mo�t chuo�i k� t�� co� na�m trong input kho�ng,kho�ng ke� da�u ca�ch"
TB(6, 3) = "kie�m tra xem mo�t so� nguye�n co� na�m trong input kho�ng"
TB(6, 4) = "Kie�m tra xem mo�t so� th��c co� trong input kho�ng"
TB(6, 5) = "L��p Scanner kho�ng co� ph��ng th��c na�y"
DA(6) = 1

HB(7) = "Ph��ng th��c hasNextLine() trong l��p Scanner co� ta�c du�ng g�?"
TB(7, 1) = "kie�m tra xem 1 chuo�i k� t�� na�m trong input kho�ng,ke� ca� da�u ca�ch"
TB(7, 2) = "kie�m tra xem mo�t chuo�i k� t�� co� na�m trong input kho�ng,kho�ng ke� da�u ca�ch"
TB(7, 3) = "kie�m tra xem mo�t so� nguye�n co� na�m trong input kho�ng"
TB(7, 4) = "Kie�m tra xem mo�t so� th��c co� trong input kho�ng"
TB(7, 5) = "L��p Scanner kho�ng co� ph��ng th��c na�y"
DA(7) = 1

HB(8) = "Ph��ng th��c hasNextInt() trong l��p Scanner co� ta�c du�ng g�?"
TB(8, 1) = "kie�m tra xem 1 chuo�i k� t�� na�m trong input kho�ng,ke� ca� da�u ca�ch"
TB(8, 2) = "kie�m tra xem mo�t chuo�i k� t�� co� na�m trong input kho�ng,kho�ng ke� da�u ca�ch"
TB(8, 3) = "kie�m tra xem mo�t so� nguye�n co� na�m trong input kho�ng"
TB(8, 4) = "Kie�m tra xem mo�t so� th��c co� trong input kho�ng"
TB(8, 5) = "L��p Scanner kho�ng co� ph��ng th��c na�y"
DA(8) = 1

HB(9) = "Cho kho�i le�nh sau:if (yourSale >= target) bonus = 100 + 0.01 * (yourSale - target);else  & vbCrLf & bonus = 0 v�i yourSale = 3000, target = 2500, v�y bonus = ?"
TB(9, 1) = "105"
TB(9, 2) = "150"
TB(9, 3) = "600"
TB(9, 4) = "0"
TB(9, 5) = "le�nh sai"
DA(9) = 1

HB(10) = "Cho kho�i le�nh sau:if (yourSale >= target) bonus = 100 + 0.01 * (yourSale - target);else  & vbCrLf & bonus = 0 v�i yourSale = 3000, target = 2500, v�y bonus = ?"
TB(10, 1) = "105"
TB(10, 2) = "150"
TB(10, 3) = "600"
TB(10, 4) = "0"
TB(10, 5) = "Le�nh sai"
DA(10) = 1

HB(11) = "Cho kho�i le�nh  sau:if (yourSale >= target) bonus = 100 + 0.01 * (yourSale - target);else  & vbCrLf & bonus = 0 v�i yourSale = 3000, target = 2500, v�y bonus = ?"
TB(11, 1) = "20.015"
TB(11, 2) = "15"
TB(11, 3) = "15.015"
TB(11, 4) = "20.035015"
TB(11, 5) = "Le�nh sai"
DA(11) = 1

HB(12) = "Cho kho�i le�nh  sau:if (yourSale >= target) bonus = 100 + 0.01 * (yourSale - target);else  & vbCrLf & bonus = 0 v�i yourSale = 3000, target = 2500, v�y bonus = ?"
TB(12, 1) = "20.015"
TB(12, 2) = "15"
TB(12, 3) = "15.015"
TB(12, 4) = "20.035015"
TB(12, 5) = "Le�nh sai"
DA(12) = 1

HB(13) = "Cho kho�i le�nh sau:int years = 0;do {balance += payment;double interest = balance * 0.1 / 100;balance += interestyears ++;} while (balance < goal);Cho balance = 20, goal = 20, payment = 5.Sau khi cha�y kho�i le�nh, balance = ?"
TB(13, 1) = "20.015"
TB(13, 2) = "15"
TB(13, 3) = "15.015"
TB(13, 4) = "20.035015"
TB(13, 5) = "le�nh sai"
DA(13) = 1

HB(14) = "Cho kho�i le�nh sau:int years = 0;do {balance += payment;double interest = balance * 0.1 / 100;balance += interestyears ++;} while (balance < goal);Cho balance = 20, goal = 20, payment = 5. Sau khi cha�y kho�i le�nh, years = ?"
TB(14, 1) = "20.015"
TB(14, 2) = "1"
TB(14, 3) = "15.015"
TB(14, 4) = "2"
TB(4, 5) = "le�nh sai"
DA(14) = 1

HB(15) = "con me�o ke�u the� na�o?"
TB(15, 1) = "meo meo"
TB(15, 2) = "ga�u ga�u"
TB(15, 3) = "e�c e�c"
TB(15, 4) = "qua�c qua�c"
TB(15, 5) = "ta�t ca� �e�u sai"
DA(15) = 1

HB(16) = "con cho� ke�u the� na�o?"
TB(16, 1) = "meo meo"
TB(16, 2) = "ga�u ga�u"
TB(16, 3) = "e�c e�c"
TB(16, 4) = "qua�c qua�c"
TB(16, 5) = "ta�t ca� �e�u sai"
DA(16) = 1

HB(17) = "con ga� ke�u the� na�o?"
TB(17, 1) = "meo meo"
TB(17, 2) = "ga�u ga�u"
TB(17, 3) = "e�c e�c"
TB(17, 4) = "qua�c qua�c"
TB(17, 5) = "ta�t ca� �e�u sai"
DA(17) = 1

HB(18) = "con ga�u ke�u the� na�o?"
TB(18, 1) = "meo meo"
TB(18, 2) = "ga�u ga�u"
TB(18, 3) = "e�c e�c"
TB(18, 4) = "qua�c qua�c"
TB(18, 5) = "ta�t ca� �e�u sai"
DA(18) = 1

HB(19) = "con chim tu hu� ke�u the� na�o?"
TB(19, 1) = "meo meo"
TB(19, 2) = "ga�u ga�u"
TB(19, 3) = "e�c e�c"
TB(19, 4) = "qua�c qua�c"
TB(19, 5) = "ta�t ca� �e�u sai"
DA(19) = 1

HB(20) = "con cu� ke�u the� na�o?"
TB(20, 1) = "meo meo"
TB(20, 2) = "ga�u ga�u"
TB(20, 3) = "e�c e�c"
TB(20, 4) = "qua�c qua�c"
TB(20, 5) = "ta�t ca� �e�u sai"
DA(20) = 1

HB(21) = "con l��n ke�u the� na�o?"
TB(21, 1) = "meo meo"
TB(21, 2) = "ga�u ga�u"
TB(21, 3) = "e�c e�c"
TB(21, 4) = "qua�c qua�c"
TB(21, 5) = "ta�t ca� �e�u sai"
DA(21) = 1

HB(22) = "con v�t ke�u the� na�o?"
TB(22, 1) = "meo meo"
TB(22, 2) = "ga�u ga�u"
TB(22, 3) = "e�c e�c"
TB(22, 4) = "qua�c qua�c"
TB(22, 5) = "ta�t ca� �e�u sai"
DA(22) = 1

HB(23) = "tra�i �a�t co� h�nh g�?"
TB(23, 1) = "h�nh vuo�ng"
TB(23, 2) = "h�nh tro�n"
TB(23, 3) = "h�nh ca�u"
TB(23, 4) = "h�nh elip"
TB(23, 5) = "kho�ng co� �a�p a�n �u�ng"
DA(23) = 1

HB(24) = "mo�t na�m co� bao nhie�u tha�ng?"
TB(24, 1) = "12"
TB(24, 2) = "13"
TB(24, 3) = "14"
TB(24, 4) = "15"
TB(24, 5) = "kho�ng co� �a�p a�n �u�ng"
DA(24) = 1

HB(25) = "Java platform go�m ma�y pha�n?"
TB(25, 1) = "1"
TB(25, 2) = "2"
TB(25, 3) = "3"
TB(25, 4) = "4"
TB(25, 5) = "kho�ng co� �a�p a�n �u�ng"
DA(25) = 1

HB(26) = "co� bao nhie�u ca�ch vie�t chu� th�ch trong java?"
TB(26, 1) = "1"
TB(26, 2) = "2"
TB(26, 3) = "3"
TB(26, 4) = "4"
TB(26, 5) = "5"
DA(26) = 1

HB(27) = "khai ba�o na�o d��oi �a�y la� �u�ng?"
TB(27, 1) = "public class default {}"
TB(27, 2) = "protected inner class engine {}"
TB(27, 3) = "final class outer {}"
TB(27, 4) = "ta�t ca� �e�u �u�ng"
TB(27, 5) = "ta�t ca� �e�u sai"
DA(27) = 1

HB(28) = "mo�t l��p trong java co� bao nhie�u l��p cha?"
TB(28, 1) = "1"
TB(28, 2) = "2"
TB(28, 3) = "3"
TB(28, 4) = "4"
TB(28, 5) = "5"
DA(28) = 1

HB(29) = "mo�t l��p trong java co� bao nhie�u l��p con?"
TB(29, 1) = "1"
TB(29, 2) = "2"
TB(29, 3) = "3"
TB(29, 4) = "4"
TB(29, 5) = "vo� so�"
DA(29) = 1

HB(30) = "mo�t ch��ng tr�nh go�m 2 class se� co� ba nhie�u ph��ng th��c main?"
TB(30, 1) = "1"
TB(30, 2) = "2"
TB(30, 3) = "3"
TB(30, 4) = "4"
TB(30, 5) = "5"
DA(30) = 1

HB(31) = "co� bao nhie�u loa�i bie�n trong java?"
TB(31, 1) = "1"
TB(31, 2) = "2"
TB(31, 3) = "3"
TB(31, 4) = "4"
TB(31, 5) = "5"
DA(31) = 1

HB(32) = "tr���ng d�� lie�u la� ca�c bie�n da�ng na�o sau �a�y?"
TB(32, 1) = "bie�n �a�i die�n va� tham bie�n"
TB(32, 2) = "bie�n �a�i die�n va� bie�n l��p"
TB(32, 3) = "bie�n �a�i die�n va� bie�n cu�c bo�"
TB(32, 4) = "bie�n l��p va� tham so�"
TB(32, 5) = "kho�ng co� �a�p a�n �u�ng"
DA(32) = 1


HB(33) = "co� bao nhie�u kie�u d�� lie�u c� s�� trong java?"
TB(33, 1) = "7"
TB(33, 2) = "8"
TB(33, 3) = "9"
TB(33, 4) = "10"
TB(33, 5) = "11"
DA(33) = 1

HB(34) = "co� bao nhie�u kie�u so� nguye�n trong java?"
TB(34, 1) = "1"
TB(34, 2) = "2"
TB(34, 3) = "3"
TB(34, 4) = "4"
TB(34, 5) = "5"
DA(34) = 1

HB(35) = "h�nh thang co� ma�y ca�ch?"
TB(35, 1) = "1"
TB(35, 2) = "2"
TB(35, 3) = "3"
TB(35, 4) = "4"
TB(35, 5) = "5"
DA(35) = 1

HB(36) = "h�nh vuo�ng co� ma�y ca�ch?"
TB(36, 1) = "2"
TB(36, 2) = "3"
TB(36, 3) = "4"
TB(36, 4) = "5"
TB(36, 5) = "6"
DA(36) = 1

HB(37) = "ca�i ca�y th���ng co� ma�u g�?"
TB(37, 1) = "ma�u xanh"
TB(37, 2) = "ma�u �en"
TB(37, 3) = "ma�u tra�ng"
TB(37, 4) = "trong suo�t"
TB(37, 5) = "ma�u ho�ng"
DA(37) = 1

HB(38) = "h�nh tam gia�c co� ma�y ca�nh?"
TB(38, 1) = "1"
TB(38, 2) = "2"
TB(38, 3) = "3"
TB(38, 4) = "4"
TB(38, 5) = "5"
DA(38) = 1

HB(39) = "con cho� co� ma�y ca�i cha�n?"
TB(39, 1) = "1"
TB(39, 2) = "2"
TB(39, 3) = "3"
TB(39, 4) = "4"
TB(39, 5) = "5"
DA(39) = 1

HB(40) = "qua� tr��ng co� h�nh g�?"
TB(40, 1) = "h�nh vuo�ng"
TB(40, 2) = "h�nh tro�n"
TB(40, 3) = "h�nh elip"
TB(40, 4) = "h�nh ba�u du�c"
TB(40, 5) = "h�nh tam gia�c"
DA(40) = 1

HB(41) = "trong ke� th��a ph��ng th��c cu�a l��p con ����c khai ba�o gio�ng ph��ng th��c cu�a l��p cha ve� ca� te�n la�n tham so�?"
TB(41, 1) = "Overload"
TB(41, 2) = "Override"
TB(41, 3) = "synchronized"
TB(41, 4) = "Serializable"
TB(41, 5) = "ta�t ca� �e�u sai"
DA(41) = 1

HB(42) = "trong class ca�c ph��ng th��c tru�ng te�n kha�c tham so� la�?"
TB(42, 1) = "Overload"
TB(42, 2) = " Override"
TB(42, 3) = "synchronized"
TB(42, 4) = "Serializable"
TB(42, 5) = "ta�t ca� �e�u sai"
DA(42) = 1

HB(43) = "Cho String str = univerity, le�nh na�o d���i �a�y la�y chuo�i univer va� ga�n va�o chuo�i  str1??"
TB(43, 1) = "String str1 = str.substring(0, 6);"
TB(43, 2) = "String str1 = str.substring(0, 5);"
TB(43, 3) = "String str1 = str.substring(1, 6);"
TB(43, 4) = ". String str1 = str.substring(5);"
TB(43, 5) = "ta�t ca� �e�u sai"
DA(43) = 1

HB(44) = "��nh ngh� interface na�o sau �a�y la� kho�ng h��p le�?"
TB(44, 1) = "public interface inout {}"
TB(44, 2) = "protected interface inout { int i = 12;}"
TB(44, 3) = ". interface inout { public final int MAX_INDEX = 100;}"
TB(44, 4) = "interface input { public void indl();}"
TB(44, 5) = "ta�t ca� �e�u h��p le�"
DA(44) = 1

HB(45) = "le�nh charAt(n) co� ta�c du�ng g�?"
TB(45, 1) = "tie�m kie�m k� t�� th�� n"
TB(45, 2) = "tra� ve� k� t�� th�� n-1"
TB(45, 3) = "tra� ve� k� t�� th�� n"
TB(45, 4) = "tra� ve� k� t�� co� v� tr� ch� mu�c n"
TB(45, 5) = "ta�t ca� �e�u sai"
DA(45) = 1

HB(46) = "�e� kie�m tra 2 chuo�i co� ba�ng nhau kho�ng ta s�� du�ng ph��ng th��c na�o?"
TB(46, 1) = "string1== string2"
TB(46, 2) = "string1 = string2"
TB(46, 3) = "string1.equals(string2)"
TB(46, 4) = ". string1.equal(string2)"
TB(46, 5) = "ta�t ca� �e�u sai"
DA(46) = 1

HB(47) = "�e� �a�o ga� tr� cu�a 1 bie�n boolean,ta du�ng toa�n t�� na�o?"
TB(47, 1) = "!"
TB(47, 2) = ">>"
TB(47, 3) = "<<"
TB(47, 4) = "=="
TB(47, 5) = "ta�t ca� �e�u sai"
DA(47) = 1

HB(48) = "le�nh na�o ng��ng vo�ng la�p hie�n th��i va� ba�t �a�u vo�ng la�p tie�p theo?"
TB(48, 1) = "continue"
TB(48, 2) = "break"
TB(48, 3) = ". cease"
TB(48, 4) = "end"
TB(48, 5) = "ta�t ca� �e�u sai"
DA(48) = 1

HB(49) = "kie�u enum la� g�?"
TB(49, 1) = "la� kie�u d�� lie�p go�m ca�c tr���ng ch��a mo�t ta�p co� ��nh ca�c ha�ng so�"
TB(49, 2) = "la� kie�u s�� lie�u tho�ng ke� ca�c bie�n so�"
TB(49, 3) = "la� kie�u d�� lie�u trong java"
TB(49, 4) = "ta�t ca� �e�u �u�ng"
TB(49, 5) = "ta�t ca� �e�u sai"
DA(49) = 1

HB(50) = "co� bao nhie�u loa�i quye�n truy ca�p trong jav?"
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
