<h1>Lotus Script 程式 呼叫 Line Bot Messape API 的方法</h1>

首先先申請 Line 的官方帳號，在建立 Message API之後，在該服務的頁面切換到 Message API 的頁籤，

Line 官方帳號申請網址 :  https://developers.line.biz/en/

如下圖中，找到 Channel token ID，


## 程式說明

[以下程式來源 LED_Control.txt ]:[https://github.com/derricktsai0904/Arduino/blob/master/04%20NodeMCU/LEDControl/LED_Control.ino](https://github.com/derricktsai0904/Course/blob/main/2024.09%E6%84%9F%E6%B8%AC%E5%85%83%E4%BB%B6/Arduino%20LED%E9%9C%B9%E9%9D%82%E7%87%88/LED_Control.ino) "LED_Control.ino"
[以下程式來源 LED_Control.txt ]
``` Lotus Script

Sub Initialize
	Set xml = CreateObject("MSXML2.ServerXMLHTTP")    '// 宣告 http post xml 物件
	
	url = "https://api.line.me/v2/bot/message/push"    '// Line Bot Message API 的 Post 網址

	TOKEN_KEY = "111111111111222222222222222233333333333333aaaaaaaaaaaaaaaaabbbbbbbbbbbbbbbbbccccccc"  '// 請參閱 Line 官方帳號的 Channel token ID
	
	
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




