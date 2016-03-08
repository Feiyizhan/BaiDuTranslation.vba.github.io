Attribute VB_Name = "BaiDuTranslationModule"
Public Function translateZh_En(ByVal str As String) As String
   '调用百度翻译API，将指定内容由中文翻译为英文
    Dim par As String
    Dim data() As String
    par = generateReqStr(str, "zh", "en")
    data() = translateJson(getHttp(par)) '调用API,获取API 返回, '调用JSON转换函数，解析API返回的结果
    
'    Debug.Print data(0)
'    Debug.Print data(1)
'    Debug.Print data(2)
'    Debug.Print data(3)
'    Debug.Print data(4)
'    Debug.Print data(5)
    translateZh_En = UCase(data(3)) '获取API返回的翻译结果

End Function


Public Function translateZh_En_Batch(ByRef arr() As String) As String()
   '调用百度翻译API，将指定内容由中文翻译为英文
    Dim par As String
    Dim data() As String
    Dim returnData() As String
    Dim resp As String

    par = generateReqStrBatch(arr, "zh", "en")
    resp = getHttp(par)  '调用API,获取API 返回
    data() = translateJsonBatch(resp) '调用JSON转换函数，解析API返回的结果
    ReDim returnData(UBound(data) - 4) As String
    For i = 0 To UBound(data) - 4
        returnData(i) = UCase(Split(data(i), "|", -1, vbTextCompare)(1))
    Next
    
    translateZh_En_Batch = returnData() '获取API返回的翻译结果

End Function


Public Function translateEn_Zh(ByVal str As String) As String
    '调用百度翻译API，将指定内容由英文翻译为中文
    Dim par As String
    Dim data() As String
    par = generateReqStr(str, "en", "zh")
    data() = translateJson(getHttp(par)) '调用API,获取API 返回, '调用JSON转换函数，解析API返回的结果
    translateEn_Zh = data(3)

End Function

Public Function translateEn_Zh_Batch(ByRef arr() As String) As String()
   '调用百度翻译API，将指定内容由英文翻译为中文
    Dim par As String
    Dim data() As String
    Dim returnData() As String
    Dim resp As String
    par = generateReqStrBatch(arr, "en", "zh")
    resp = getHttp(par)  '调用API,获取API 返回
    data() = translateJsonBatch(resp) '调用JSON转换函数，解析API返回的结果
    ReDim returnData(UBound(data) - 4) As String
    For i = 0 To UBound(data) - 4
        returnData(i) = Split(data(i), "|", -1, vbTextCompare)(1)
    Next
    
    translateZh_En_Batch = returnData() '获取API返回的翻译结果

End Function


Public Function getHttpOld(str As String) As String
'调用API
Dim HttpReq As Object
Dim url As String
Set HttpReq = CreateObject("Microsoft.XMLHTTP") '创建XMLHTTP对象
url = "http://openapi.baidu.com/public/2.0/bmt/translate?client_id=iOMhRqTKNByhC80V9SbcQIpo&" & str
With HttpReq
        .Open "GET", url, False
        .setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .SEND
        'Debug.Print .responsetext
End With
'发送HTTL  Get请求， 百度API只允许GET，不允许POST

getHttp = HttpReq.responsetext

End Function

Public Function getHttp(str As String) As String
'调用API（new）
Dim HttpReq As Object
Dim url As String
Set HttpReq = CreateObject("Microsoft.XMLHTTP") '创建XMLHTTP对象
url = "http://api.fanyi.baidu.com/api/trans/vip/translate?" & str
Debug.Print (url)
With HttpReq
        .Open "GET", url, False
        .setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .SEND
        'Debug.Print .responsetext
End With
'发送HTTL  Get请求， 百度API只允许GET，不允许POST

getHttp = HttpReq.responsetext

End Function





Public Function generateReqStr(q As String, from_Str As String, to_Str As String) As String
'生成Request 字符串
    Dim appid As String
    Dim key As String
    Dim salt As Integer
    Dim sign As String
    Dim par As String
    
    Math.Randomize (Timer)
    salt = (Rnd * 1000000) Mod 20000
    key = "修改为你自己申请的API的Key"
    appid = "修改为你自己申请的API的appid"
    
    'Debug.Print (appid + q + CStr(salt) + key)
    sign = MD5_32(appid + q + CStr(salt) + key)  '转换为MD5
    'Debug.Print (sign)
    
    par = "q=" + encode(q) + "&from=" + from_Str + "&to=" + to_Str + "&appid=" + appid + "&salt=" + CStr(salt) + "&sign=" + sign  '调用urlencode 方法，将待翻译内容转换为urlencode,
    
    generateReqStr = par
    

End Function


Public Function generateReqStrBatch(q() As String, from_Str As String, to_Str As String) As String
'生成Request 字符串
    Dim appid As String
    Dim key As String
    Dim salt As Integer
    Dim sign As String
    Dim par As String
    Dim str1 As String
    'Dim ln As String
    
    'ln = ChrB(10) & ChrB(0)
    Math.Randomize (Timer)
    salt = (Rnd * 1000000) Mod 20000
    key = "修改为你自己申请的API的Key"
    appid = "修改为你自己申请的API的appid"
    
    For i = LBound(q) To UBound(q)
        str1 = str1 & q(i) & vbLf
    Next
    
    'Debug.Print (appid + str1 + CStr(salt) + key)
    sign = MD5_32(appid + str1 + CStr(salt) + key)  '转换为MD5
    
    'Debug.Print (sign)
    par = "q=" + encode(str1) + "&from=" + from_Str + "&to=" + to_Str + "&appid=" + appid + "&salt=" + CStr(salt) + "&sign=" + sign '调用urlencode 方法，将待翻译内容转换为urlencode,
    
    generateReqStrBatch = par
    

End Function


Public Function translateJson(str As String) As String()
'调用JScript 解析JSON
Dim js As Object
Dim objJSON As Object
Dim objJSON2 As Object
Dim strFunc As String
Dim returnData(6) As String

'创建Script对象
Set js = CreateObject("ScriptControl"): js.Language = "JScript"
'aa = "{""from"":""en"",""to"":""zh"",""trans_result"":[{""src"":""today"",""dst"":""\u4eca\u5929""}]}"
'获取第一层的数据内容的JavaScript函数代码
strFunc = "function getjson(s) { return eval('(' + s + ')'); }"
'获取第二层的数据内容JavaScript函数代码
strFunc2 = "function j(s) { return eval('(' + s + ').trans_result[0]'); }"
'将JavaScript函数代码加入到Script对象。
js.AddCode strFunc
js.AddCode strFunc2
Set objJSON = js.CodeObject.getjson(str) '执行函数方法 ,这是一种执行方法
On Error GoTo ErrorHandler1
Set objJSON2 = js.Run("j", str)   '执行函数方法 ,这是另一种执行方法

'获取第一层的结果
'Debug.Print objJSON.from
'Debug.Print objJSON.to
'Debug.Print objJSON.trans_result

returnData(0) = objJSON.from
'returnData(1) = objJSON.To

'获取第二层的结果
'Debug.Print CallByName(objJSON2, "src", VbGet)  '这是另一种获取属性的方法
'Debug.Print objJSON2.dst

returnData(2) = objJSON2.src
returnData(3) = objJSON2.dst

'如果API执行结果不正确，获取API的不正确的返回信息。
On Error GoTo ErrorHandler
returnData(4) = objJSON.error_code
returnData(5) = objJSON.error_msg

translateJson = returnData
Exit Function

ErrorHandler:
returnData(4) = ""
returnData(5) = ""
translateJson = returnData
Exit Function

ErrorHandler1:
returnData(0) = ""
returnData(1) = ""
returnData(2) = ""
returnData(3) = ""
returnData(4) = objJSON.error_code
returnData(5) = objJSON.error_msg
translateJson = returnData
Exit Function

End Function


Public Function translateJsonBatch(str As String) As String()
'调用JScript 解析JSON
Dim js As Object
Dim objJSON As Object
Dim objJSON2 As Object
Dim count As String
Dim count_i As Integer
Dim strFunc As String
Dim returnData(6) As String
Dim returnData2() As String

'创建Script对象
Set js = CreateObject("ScriptControl"): js.Language = "JScript"
'aa = "{""from"":""en"",""to"":""zh"",""trans_result"":[{""src"":""today"",""dst"":""\u4eca\u5929""}]}"
'获取第一层的数据内容的JavaScript函数代码
strFunc = "function getjson(s) { return eval('(' + s + ')'); }"
'获取第二层的数据个数JavaScript函数代码
strFunc1 = "function getjsonCount(s) { return eval('(' + s + ').trans_result.length'); }"
'获取第二层的数据内容JavaScript函数代码
strFunc2 = "function getjsonLevel(s,i) { return eval('(' + s + ').trans_result['+i+']'); }"
'将JavaScript函数代码加入到Script对象。
js.AddCode strFunc
js.AddCode strFunc1
js.AddCode strFunc2
On Error GoTo ErrorHandler2
Set objJSON = js.CodeObject.getjson(str) '执行函数方法 ,这是一种执行方法


On Error GoTo ErrorHandler1
count = js.CodeObject.getjsonCount(str) '获取数据个数
returnData(0) = objJSON.from
'returnData(1) = objJSON.To
count_i = Val(count)
If (count > 0) Then
    ReDim returnData2(count + 3) As String
    
'Set objJSON2 = js.Run("j", str)   '执行函数方法 ,这是另一种执行方法

returnData2(UBound(returnData2) - 3) = returnData(0)
returnData2(UBound(returnData2) - 2) = returnData(1)

    For i = 0 To count - 1  '获取返回的数组内容
        Set objJSON2 = js.CodeObject.getjsonLevel(str, i)
        returnData(2) = objJSON2.src
        returnData(3) = objJSON2.dst
        returnData2(i) = returnData(2) & "|" & returnData(3)
    Next

Else
   ReDim returnData2(4) As String  '无数据返回，则数据区域返回空
   returnData2(0) = " | "
End If
'如果API执行结果不正确，获取API的不正确的返回信息。
On Error GoTo ErrorHandler
'returnData(4) = objJSON.error_code
'returnData(5) = objJSON.error_msg
returnData2(UBound(returnData2) - 1) = returnData(4)
returnData2(UBound(returnData2)) = returnData(5)
translateJsonBatch = returnData2

Exit Function

ErrorHandler:
'获取错误信息失败，则设置错误信息为空
returnData(4) = ""
returnData(5) = ""
returnData2(UBound(returnData2) - 1) = returnData(4)
returnData2(UBound(returnData2)) = returnData(5)
translateJsonBatch = returnData2
Exit Function

ErrorHandler1: '获取第一层数据失败，则返回错误信息
returnData(0) = ""
returnData(1) = ""
returnData(2) = ""
returnData(3) = ""
returnData(4) = objJSON.error_code
returnData(5) = objJSON.error_msg
ReDim returnData2(4) As String  '无数据返回，则数据区域返回空
returnData2(UBound(returnData2) - 3) = returnData(0)
returnData2(UBound(returnData2) - 2) = returnData(1)
returnData2(UBound(returnData2) - 1) = returnData(4)
returnData2(UBound(returnData2)) = returnData(5)
returnData2(0) = " | "

translateJsonBatch = returnData2
Exit Function

ErrorHandler2: '获取第一层数据失败，则返回错误信息
   ReDim returnData2(4) As String  '无数据返回，则数据区域返回空
   returnData2(0) = " | "

translateJsonBatch = returnData2
Exit Function

End Function



Public Function encode(ByVal str As String) As String
'调用JavaScript的encodeURIComponent方法进行urlencode 编码
   Dim js As Object
   Dim strFun As String
   Dim data As String
  
   Set js = CreateObject("ScriptControl"): js.Language = "JScript"
     'aa = "{""from"":""en"",""to"":""zh"",""trans_result"":[{""src"":""today"",""dst"":""\u4eca\u5929""}]}"
    strFunc = "function getjson(s) { return eval('encodeURIComponent(\""'+s+'\"")'); }"
    js.AddCode strFunc
    str = Replace(str, vbCrLf, "\r\n")  '转换回车换行符为URL回车换行符
    str = Replace(str, vbLf, "\n")   '转换换行符为URL换行符
   data = js.CodeObject.getjson(str)
   'Debug.Print data
   encode = data
End Function


Sub testTranslate()
    Debug.Print translateZh_En("陕西双和劳动防护装有限公司")
    Dim testData() As String
    ReDim testData(5)
    testData(0) = "苹果"
    testData(1) = "香蕉"
    testData(2) = "不作就不会死"
    testData(3) = "红色"
    testData(4) = "绿色"
    testData(5) = "黄色"
    
    Dim returnData() As String
    returnData = translateZh_En_Batch(testData)
    For i = 0 To UBound(returnData)
        Debug.Print returnData(i)
    Next
    
    
End Sub


