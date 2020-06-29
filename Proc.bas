Attribute VB_Name = "Proc"
Private Winhttp As New Winhttp.WinHttpRequest

Public Function firstProc() As String
On Error GoTo Err:
    Dim ck As String
    
    ck = getCookieAll
    
    Winhttp.Open "GET", "https://m.worldjob.or.kr:444/index.do"
    
    Winhttp.SetRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
    Winhttp.SetRequestHeader "Accept-Language", "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
    Winhttp.SetRequestHeader "Connection", "keep-alive"
    Winhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    Winhttp.SetRequestHeader "Host", "m.worldjob.or.kr:444"
    Winhttp.SetRequestHeader "Origin", "https://m.worldjob.or.kr:444"
    Winhttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36"
    If ck <> "" Then Winhttp.SetRequestHeader "Cookie", ck
    
    Winhttp.Send
    
    setCookie Winhttp.GetAllResponseHeaders
    
    firstProc = Winhttp.ResponseText
Err:
End Function

Public Sub setBeacon(datex As String, typr As String)
On Error GoTo Err:
    Dim ck As String
    Dim contents_ As String
    Dim par1 As String
    Dim par2 As String
    
    par1 = inlineSptParameter(beaconAddr, "joCrtfcNo=")
    par2 = inlineSptParameter(beaconAddr, "indMbrid=")
    
    ck = getCookieAll
    
    '//attendScd ( I : ¿‘Ω«, O : ≈Ω«, R : ø‹√‚, K : ∫π±Õ
    contents_ = "indMbrid=" & par2 & "&joCrtfcNo=" & par1 & "&joCrtfcDsp=1&joCrtfcDspSn=1&attendScd=" & typr & "&attendOs=IOS&attendOsId=UID&prcerAtendDe=" & datex & "&localtimeGapHour=0"
    
    Winhttp.Open "POST", "https://m.worldjob.or.kr:444/indvdl/epmtSprtMng/ajaxInsertAtendBeacon.do"
    
    Winhttp.SetRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
    Winhttp.SetRequestHeader "Accept-Language", "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
    Winhttp.SetRequestHeader "Connection", "keep-alive"
    Winhttp.SetRequestHeader "Content-Length", Len(contents_)
    Winhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    Winhttp.SetRequestHeader "Host", "m.worldjob.or.kr:444"
    Winhttp.SetRequestHeader "Origin", "https://m.worldjob.or.kr:444"
    Winhttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36"
    Winhttp.SetRequestHeader "Referer", "https://m.worldjob.or.kr:444/indvdl/epmtSprtMng/beaconAtendenceAuto.do?joCrtfcNo=" & par1 & "&joCrtfcDsp=1&joCrtfcDspSn=1&prcerAtendDe=" & datex & "&indMbrid=" & par2
    Winhttp.SetRequestHeader "Sec-Fetch-Mode", "cors"
    Winhttp.SetRequestHeader "Sec-Fetch-Site", "same-origin"
    Winhttp.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    If ck <> "" Then Winhttp.SetRequestHeader "Cookie", ck
    
    Winhttp.Send contents_
    
    setCookie Winhttp.GetAllResponseHeaders
Err:
End Sub

Public Function Login(id As String, passwd As String) As String
On Error GoTo Err:
    Dim contents_ As String
    Dim ck As String
    
    ck = getCookieAll
    
    contents_ = "loginProcessType=2&memberType=PER&id=" & id & "&password=" & passwd
    
    Winhttp.Open "POST", "https://m.worldjob.or.kr:444/login/loginProcess.do"
    
    Winhttp.SetRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
    Winhttp.SetRequestHeader "Accept-Language", "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
    Winhttp.SetRequestHeader "Connection", "keep-alive"
    Winhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    Winhttp.SetRequestHeader "Host", "m.worldjob.or.kr:444"
    Winhttp.SetRequestHeader "Origin", "https://m.worldjob.or.kr:444"
    Winhttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36"
    Winhttp.SetRequestHeader "Referer", "https://m.worldjob.or.kr:444/login/login.do"
    Winhttp.SetRequestHeader "Content-Length", Len(contents_)
    Winhttp.SetRequestHeader "Sec-Fetch-Mode", "cors"
    Winhttp.SetRequestHeader "Sec-Fetch-Site", "same-origin"
    Winhttp.SetRequestHeader "X-Requested-With", "XMLHttpRequest"
    If ck <> "" Then Winhttp.SetRequestHeader "Cookie", ck
    
    Winhttp.Send contents_

    Login = Winhttp.GetAllResponseHeaders
Err:
End Function

Private Function inlineSptParameter(URL As String, parName As String) As String
On Error GoTo Err:
    
    Dim firstSpt As String
    
    firstSpt = Split(URL, parName)(1)
    firstSpt = Split(firstSpt, "&")(0)
    
Err:
    inlineSptParameter = firstSpt
End Function

Public Function getbeaconTime(day As String) As String
On Error GoTo Err:

    Dim ck As String
    Dim src As String
    Dim par1 As String
    Dim par2 As String
    Dim rst As String
    
    rst = "/"
    par1 = inlineSptParameter(beaconAddr, "joCrtfcNo=")
    par2 = inlineSptParameter(beaconAddr, "indMbrid=")

    ck = getCookieAll
    
    Winhttp.Open "GET", "https://m.worldjob.or.kr:444/indvdl/epmtSprtMng/beaconAtendenceAuto.do?joCrtfcNo=" & par1 & "&joCrtfcDsp=1&joCrtfcDspSn=1&prcerAtendDe=" & day & "&indMbrid=" & par2
    
    Winhttp.SetRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
    Winhttp.SetRequestHeader "Accept-Language", "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
    Winhttp.SetRequestHeader "Connection", "keep-alive"
    Winhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    Winhttp.SetRequestHeader "Host", "m.worldjob.or.kr:444"
    Winhttp.SetRequestHeader "Origin", "https://m.worldjob.or.kr:444"
    Winhttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36"
    If ck <> "" Then Winhttp.SetRequestHeader "Cookie", ck
    
    Winhttp.Send
    
    src = Winhttp.ResponseText
    
    rst = Split(Split(src, "<span id=""btnEnterTm"">")(1), "</span>")(0) & "/" & Split(Split(src, "<span id=""btnOutTm"">")(1), "</span>")(0)
Err:
    getbeaconTime = rst
End Function
