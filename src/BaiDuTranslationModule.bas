Attribute VB_Name = "BaiDuTranslationModule"
Public Function translateZh_En(ByVal str As String) As String
   '���ðٶȷ���API����ָ�����������ķ���ΪӢ��
    Dim par As String
    Dim data() As String
    par = generateReqStr(str, "zh", "en")
    data() = translateJson(getHttp(par)) '����API,��ȡAPI ����, '����JSONת������������API���صĽ��
    
'    Debug.Print data(0)
'    Debug.Print data(1)
'    Debug.Print data(2)
'    Debug.Print data(3)
'    Debug.Print data(4)
'    Debug.Print data(5)
    translateZh_En = UCase(data(3)) '��ȡAPI���صķ�����

End Function


Public Function translateZh_En_Batch(ByRef arr() As String) As String()
   '���ðٶȷ���API����ָ�����������ķ���ΪӢ��
    Dim par As String
    Dim data() As String
    Dim returnData() As String
    Dim resp As String

    par = generateReqStrBatch(arr, "zh", "en")
    resp = getHttp(par)  '����API,��ȡAPI ����
    data() = translateJsonBatch(resp) '����JSONת������������API���صĽ��
    ReDim returnData(UBound(data) - 4) As String
    For i = 0 To UBound(data) - 4
        returnData(i) = UCase(Split(data(i), "|", -1, vbTextCompare)(1))
    Next
    
    translateZh_En_Batch = returnData() '��ȡAPI���صķ�����

End Function


Public Function translateEn_Zh(ByVal str As String) As String
    '���ðٶȷ���API����ָ��������Ӣ�ķ���Ϊ����
    Dim par As String
    Dim data() As String
    par = generateReqStr(str, "en", "zh")
    data() = translateJson(getHttp(par)) '����API,��ȡAPI ����, '����JSONת������������API���صĽ��
    translateEn_Zh = data(3)

End Function

Public Function translateEn_Zh_Batch(ByRef arr() As String) As String()
   '���ðٶȷ���API����ָ��������Ӣ�ķ���Ϊ����
    Dim par As String
    Dim data() As String
    Dim returnData() As String
    Dim resp As String
    par = generateReqStrBatch(arr, "en", "zh")
    resp = getHttp(par)  '����API,��ȡAPI ����
    data() = translateJsonBatch(resp) '����JSONת������������API���صĽ��
    ReDim returnData(UBound(data) - 4) As String
    For i = 0 To UBound(data) - 4
        returnData(i) = Split(data(i), "|", -1, vbTextCompare)(1)
    Next
    
    translateZh_En_Batch = returnData() '��ȡAPI���صķ�����

End Function


Public Function getHttpOld(str As String) As String
'����API
Dim HttpReq As Object
Dim url As String
Set HttpReq = CreateObject("Microsoft.XMLHTTP") '����XMLHTTP����
url = "http://openapi.baidu.com/public/2.0/bmt/translate?client_id=iOMhRqTKNByhC80V9SbcQIpo&" & str
With HttpReq
        .Open "GET", url, False
        .setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .SEND
        'Debug.Print .responsetext
End With
'����HTTL  Get���� �ٶ�APIֻ����GET��������POST

getHttp = HttpReq.responsetext

End Function

Public Function getHttp(str As String) As String
'����API��new��
Dim HttpReq As Object
Dim url As String
Set HttpReq = CreateObject("Microsoft.XMLHTTP") '����XMLHTTP����
url = "http://api.fanyi.baidu.com/api/trans/vip/translate?" & str
Debug.Print (url)
With HttpReq
        .Open "GET", url, False
        .setRequestHeader "content-type", "application/x-www-form-urlencoded"
        .SEND
        'Debug.Print .responsetext
End With
'����HTTL  Get���� �ٶ�APIֻ����GET��������POST

getHttp = HttpReq.responsetext

End Function





Public Function generateReqStr(q As String, from_Str As String, to_Str As String) As String
'����Request �ַ���
    Dim appid As String
    Dim key As String
    Dim salt As Integer
    Dim sign As String
    Dim par As String
    
    Math.Randomize (Timer)
    salt = (Rnd * 1000000) Mod 20000
    key = "�޸�Ϊ���Լ������API��Key"
    appid = "�޸�Ϊ���Լ������API��appid"
    
    'Debug.Print (appid + q + CStr(salt) + key)
    sign = MD5_32(appid + q + CStr(salt) + key)  'ת��ΪMD5
    'Debug.Print (sign)
    
    par = "q=" + encode(q) + "&from=" + from_Str + "&to=" + to_Str + "&appid=" + appid + "&salt=" + CStr(salt) + "&sign=" + sign  '����urlencode ������������������ת��Ϊurlencode,
    
    generateReqStr = par
    

End Function


Public Function generateReqStrBatch(q() As String, from_Str As String, to_Str As String) As String
'����Request �ַ���
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
    key = "�޸�Ϊ���Լ������API��Key"
    appid = "�޸�Ϊ���Լ������API��appid"
    
    For i = LBound(q) To UBound(q)
        str1 = str1 & q(i) & vbLf
    Next
    
    'Debug.Print (appid + str1 + CStr(salt) + key)
    sign = MD5_32(appid + str1 + CStr(salt) + key)  'ת��ΪMD5
    
    'Debug.Print (sign)
    par = "q=" + encode(str1) + "&from=" + from_Str + "&to=" + to_Str + "&appid=" + appid + "&salt=" + CStr(salt) + "&sign=" + sign '����urlencode ������������������ת��Ϊurlencode,
    
    generateReqStrBatch = par
    

End Function


Public Function translateJson(str As String) As String()
'����JScript ����JSON
Dim js As Object
Dim objJSON As Object
Dim objJSON2 As Object
Dim strFunc As String
Dim returnData(6) As String

'����Script����
Set js = CreateObject("ScriptControl"): js.Language = "JScript"
'aa = "{""from"":""en"",""to"":""zh"",""trans_result"":[{""src"":""today"",""dst"":""\u4eca\u5929""}]}"
'��ȡ��һ����������ݵ�JavaScript��������
strFunc = "function getjson(s) { return eval('(' + s + ')'); }"
'��ȡ�ڶ������������JavaScript��������
strFunc2 = "function j(s) { return eval('(' + s + ').trans_result[0]'); }"
'��JavaScript����������뵽Script����
js.AddCode strFunc
js.AddCode strFunc2
Set objJSON = js.CodeObject.getjson(str) 'ִ�к������� ,����һ��ִ�з���
On Error GoTo ErrorHandler1
Set objJSON2 = js.Run("j", str)   'ִ�к������� ,������һ��ִ�з���

'��ȡ��һ��Ľ��
'Debug.Print objJSON.from
'Debug.Print objJSON.to
'Debug.Print objJSON.trans_result

returnData(0) = objJSON.from
'returnData(1) = objJSON.To

'��ȡ�ڶ���Ľ��
'Debug.Print CallByName(objJSON2, "src", VbGet)  '������һ�ֻ�ȡ���Եķ���
'Debug.Print objJSON2.dst

returnData(2) = objJSON2.src
returnData(3) = objJSON2.dst

'���APIִ�н������ȷ����ȡAPI�Ĳ���ȷ�ķ�����Ϣ��
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
'����JScript ����JSON
Dim js As Object
Dim objJSON As Object
Dim objJSON2 As Object
Dim count As String
Dim count_i As Integer
Dim strFunc As String
Dim returnData(6) As String
Dim returnData2() As String

'����Script����
Set js = CreateObject("ScriptControl"): js.Language = "JScript"
'aa = "{""from"":""en"",""to"":""zh"",""trans_result"":[{""src"":""today"",""dst"":""\u4eca\u5929""}]}"
'��ȡ��һ����������ݵ�JavaScript��������
strFunc = "function getjson(s) { return eval('(' + s + ')'); }"
'��ȡ�ڶ�������ݸ���JavaScript��������
strFunc1 = "function getjsonCount(s) { return eval('(' + s + ').trans_result.length'); }"
'��ȡ�ڶ������������JavaScript��������
strFunc2 = "function getjsonLevel(s,i) { return eval('(' + s + ').trans_result['+i+']'); }"
'��JavaScript����������뵽Script����
js.AddCode strFunc
js.AddCode strFunc1
js.AddCode strFunc2
On Error GoTo ErrorHandler2
Set objJSON = js.CodeObject.getjson(str) 'ִ�к������� ,����һ��ִ�з���


On Error GoTo ErrorHandler1
count = js.CodeObject.getjsonCount(str) '��ȡ���ݸ���
returnData(0) = objJSON.from
'returnData(1) = objJSON.To
count_i = Val(count)
If (count > 0) Then
    ReDim returnData2(count + 3) As String
    
'Set objJSON2 = js.Run("j", str)   'ִ�к������� ,������һ��ִ�з���

returnData2(UBound(returnData2) - 3) = returnData(0)
returnData2(UBound(returnData2) - 2) = returnData(1)

    For i = 0 To count - 1  '��ȡ���ص���������
        Set objJSON2 = js.CodeObject.getjsonLevel(str, i)
        returnData(2) = objJSON2.src
        returnData(3) = objJSON2.dst
        returnData2(i) = returnData(2) & "|" & returnData(3)
    Next

Else
   ReDim returnData2(4) As String  '�����ݷ��أ����������򷵻ؿ�
   returnData2(0) = " | "
End If
'���APIִ�н������ȷ����ȡAPI�Ĳ���ȷ�ķ�����Ϣ��
On Error GoTo ErrorHandler
'returnData(4) = objJSON.error_code
'returnData(5) = objJSON.error_msg
returnData2(UBound(returnData2) - 1) = returnData(4)
returnData2(UBound(returnData2)) = returnData(5)
translateJsonBatch = returnData2

Exit Function

ErrorHandler:
'��ȡ������Ϣʧ�ܣ������ô�����ϢΪ��
returnData(4) = ""
returnData(5) = ""
returnData2(UBound(returnData2) - 1) = returnData(4)
returnData2(UBound(returnData2)) = returnData(5)
translateJsonBatch = returnData2
Exit Function

ErrorHandler1: '��ȡ��һ������ʧ�ܣ��򷵻ش�����Ϣ
returnData(0) = ""
returnData(1) = ""
returnData(2) = ""
returnData(3) = ""
returnData(4) = objJSON.error_code
returnData(5) = objJSON.error_msg
ReDim returnData2(4) As String  '�����ݷ��أ����������򷵻ؿ�
returnData2(UBound(returnData2) - 3) = returnData(0)
returnData2(UBound(returnData2) - 2) = returnData(1)
returnData2(UBound(returnData2) - 1) = returnData(4)
returnData2(UBound(returnData2)) = returnData(5)
returnData2(0) = " | "

translateJsonBatch = returnData2
Exit Function

ErrorHandler2: '��ȡ��һ������ʧ�ܣ��򷵻ش�����Ϣ
   ReDim returnData2(4) As String  '�����ݷ��أ����������򷵻ؿ�
   returnData2(0) = " | "

translateJsonBatch = returnData2
Exit Function

End Function



Public Function encode(ByVal str As String) As String
'����JavaScript��encodeURIComponent��������urlencode ����
   Dim js As Object
   Dim strFun As String
   Dim data As String
  
   Set js = CreateObject("ScriptControl"): js.Language = "JScript"
     'aa = "{""from"":""en"",""to"":""zh"",""trans_result"":[{""src"":""today"",""dst"":""\u4eca\u5929""}]}"
    strFunc = "function getjson(s) { return eval('encodeURIComponent(\""'+s+'\"")'); }"
    js.AddCode strFunc
    str = Replace(str, vbCrLf, "\r\n")  'ת���س����з�ΪURL�س����з�
    str = Replace(str, vbLf, "\n")   'ת�����з�ΪURL���з�
   data = js.CodeObject.getjson(str)
   'Debug.Print data
   encode = data
End Function


Sub testTranslate()
    Debug.Print translateZh_En("����˫���Ͷ�����װ���޹�˾")
    Dim testData() As String
    ReDim testData(5)
    testData(0) = "ƻ��"
    testData(1) = "�㽶"
    testData(2) = "�����Ͳ�����"
    testData(3) = "��ɫ"
    testData(4) = "��ɫ"
    testData(5) = "��ɫ"
    
    Dim returnData() As String
    returnData = translateZh_En_Batch(testData)
    For i = 0 To UBound(returnData)
        Debug.Print returnData(i)
    Next
    
    
End Sub


