<h1>Lotus Script 程式 呼叫 Line Bot Messape API 的方法</h1>

首先先申請 Line 的官方帳號，在建立 Message API之後，在該服務的頁面切換到 Message API 的頁籤，

Line 官方帳號申請網址 :  https://developers.line.biz/en/

如下圖中，找到 Channel access token

## Line Bot develope 管理頁面(Message API)
>![](https://github.com/derricktsai0904/LotusNotes/blob/master/LineBot_MessageAPI/LineBotToken.jpg?raw=true)

## 程式說明

[以下程式來源 linebot_with_lotusscript.txt ]:https://github.com/derricktsai0904/LotusNotes/blob/master/LineBot_MessageAPI/linebot_with_lotusscript.txt "linebot_with_lotusscript.txt"
[以下程式來源 linebot_with_lotusscript.txt ]
``` Lotus Script

Sub Initialize
	Set xml = CreateObject("MSXML2.ServerXMLHTTP")    '// 宣告 http post xml 物件
	
	url = "https://api.line.me/v2/bot/message/push"    '// Line Bot Message API 的 Post 網址

	TOKEN_KEY = "111111111111222222222222222233333333333333aaaaaaaaaaaaaaaaabbbbbbbbbbbbbbbbbccccccc"
        '// 請參閱 Line 官方帳號的 Channel access token
	
	
	Dim jsonBody As String  '// 需要拋送的 Line 訊息 的 Post 字串
	
	GID = "CCCC888CCC1111BBBBB34567"   '// 要拋送 Line 群組的 Group ID
	strMessage = "這是從 Notes 發的 LINE 訊息"     '// 要拋送Line訊息的字串
	
	jsonBody = |{"to":"| & GID & |","messages":[{"type":"text","text":"| & strMessage &  |"}]}|   '// Post 訊息 必須符合 json 格式

  '// 帶入相關參數發送 POST 訊息
	With xml
		.Open "POST", url, False
		.setRequestHeader "Authorization", "Bearer " & TOKEN_KEY
		.setRequestHeader "Content-Type", "application/json"
		.send jsonBody
	End With
	
	tmp = xml.responseText
	
	Msgbox tmp
	
End Sub


```




