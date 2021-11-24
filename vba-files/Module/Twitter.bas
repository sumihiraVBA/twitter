Attribute VB_Name = "Twitter"
' Twitter VBA
' (c) すみひら
'    - https://sumihira.site/
'    - https://github.com/sumihiraVBA
'
' TwitterAPIを利用をまとめたモジュールとなります。
' VBA-JSON v2.3.1を利用して作成していますので、
' 下記から標準モジュールとして「JsonConverter」をインポートしてからご利用ください。
'    - https://github.com/VBA-tools/VBA-JSON
'
' 実行時にエラーが出た場合は参照設定が足りていない可能性があるので、
' HPをご確認いただき、設定を導入してください。
'
' ご自身のTwitterAPIキーを登録してください。
Private Const oauth_consumer_key As String = "oauth_consumer_key"
Private Const oauth_consumer_secret As String = "oauth_consumer_secret"

Private Const bearer_token As String = "bearer_token"

Private Const oauth_token As String = "oauth_token"
Private Const oauth_token_secret As String = "oauth_token_secret"


Dim userId As String

Sub test()
    'ユーザ名を指定する
    Call outputTweetToExcel("test")

End Sub

Public Sub outputTweetToExcel(ByVal userName As String)
    
    Dim i As Integer

    Dim tweets As Collection
    Dim tweet As Dictionary
    
    Set tweets = getTweets(userName)
    
    i = 1
    
    For Each tweet In tweets
        
        i = i + 1
        
        '出力先を変更するのであればActiveSheetを変更する
        With ActiveSheet
            .Cells(i, 1).Value = tweet("id")
            .Cells(i, 2).Value = tweet("text")
            .Cells(i, 3).Value = tweet("created_at")
        End With
        
    Next
    
End Sub

Public Function getTweets(ByVal userName As String) As Collection
    
    Dim url As String
    
    url = "https://api.twitter.com/2/users/" & getUserId(userName) & "/tweets"
    
    '必要な情報はtweet.fieldsにパラメータ付与で拡張
    Set getTweets = responseJson(url, queryParameter:="expansions=referenced_tweets.id&tweet.fields=created_at&max_results=100")
    
End Function

Private Function getUserId(ByVal userName As String, Optional ByVal resetFlag As Boolean = False) As String
    
    '過去投稿したtweet全件に対して処理を掛けるといった
    'コード実装時に何度も同じ処理が実施されないように
    
    If resetFlag = True Or userId = "" Then
            
        Dim json As Dictionary
        Dim url As String
        
        url = "https://api.twitter.com/2/users/by/username/" & userName
        
        userId = responseJson(url)("id")
        
    End If
    
    getUserId = userId

End Function

Private Function addVerticalLine(ByVal originString As String) As String

    'JsonConverterの仕様では"100000000"という文字列型をdictionaryに格納しても、
    '100000000という数値型になってしまう。この状態でPOSTした場合型不一致になってしまうので、
    'dictionaryに格納前にバーティカルラインを追加し、POST前に置換して消すことで文字列型としてJSONを利用できる。
    addVerticalLine = "|" & originString & "|"
    
End Function

Public Function postFollowing(ByVal userName As String, ByVal followUserName As String) As Boolean

    Dim url As String
    Dim target_user_id As String
    Dim user As New Dictionary
    
    target_user_id = addVerticalLine(getUserId(followUserName, True))
    user.Add "target_user_id", target_user_id
    
    url = "https://api.twitter.com/2/users/" & getUserId(userName, True) & "/following"
    
    postFollowing = responseJson(url, "POST", user)("following")
    
End Function

Public Function deleteFollowing(ByVal userName As String, ByVal followUserName As String) As Boolean
    
    Dim url As String
    Dim target_user_id As String
    
    target_user_id = getUserId(followUserName)
    userId = ""
    
    url = "https://api.twitter.com/2/users/" & getUserId(userName, True) & "/following/" & target_user_id
    
    deleteFollowing = responseJson(url, "DELETE")("following")
    
End Function

Public Function getFollowers(ByVal userName As String) As Collection
    
    Dim url As String
    
    url = "https://api.twitter.com/2/users/" & getUserId(userName, True) & "/followers"
    Set getFollowers = responseJson(url)("following")
    
End Function

Public Function getFollowing(ByVal userName As String) As Collection
    
    Dim url As String
    
    url = "https://api.twitter.com/2/users/" & getUserId(userName) & "/following"
    Set getFollowing = responseJson(url)
    
End Function

Public Function postTweet(ByVal tweetText As String) As String

    Dim url As String
    Dim tweet As New Dictionary
    
    tweet.Add "text", tweetText

    url = "https://api.twitter.com/2/tweets"
    postTweet = responseJson(url, "POST", tweet)("id")
    
End Function

Public Function deleteTweet(ByVal tweetId As String) As Boolean
    
    Dim url As String
    
    url = "https://api.twitter.com/2/tweets/" & tweetId
    deleteTweet = responseJson(url, "DELETE")("deleted")
    
End Function

Public Sub deleteTweets()
    
    Dim tweets As Collection
    Dim tweet As Dictionary
    Dim url As String
    
    Set tweets = getTweets("7dk1s20h")
    
    For Each tweet In tweets
        
        url = "https://api.twitter.com/2/tweets/" & tweet("id")
        
        '削除結果を表示する
        '結果をシートに出力したい場合はここを変更する
        Debug.Print responseJson(url, "DELETE")("deleted")
    
    Next
    
End Sub

Private Function signinKey() As String
    
    signinKey = oauth_consumer_secret & "&" & oauth_token_secret
    
End Function

Private Function sourceEncodeUrl(ByVal method As String, ByVal url As String, ByVal oauth_timestamp As String, ByVal oauth_nonce As String) As String
    
    sourceEncodeUrl = method & "&" & WorksheetFunction.EncodeURL(url) & "&"
    
    sourceEncodeUrl = sourceEncodeUrl & _
                     "oauth_consumer_key%3D" & oauth_consumer_key & _
                  "%26oauth_nonce%3D" & oauth_nonce & _
                  "%26oauth_signature_method%3DHMAC-SHA1" & _
                  "%26oauth_timestamp%3D" & oauth_timestamp & _
                  "%26oauth_token%3D" & oauth_token & _
                  "%26oauth_version%3D1.0"
    
End Function

Private Function authorization(ByVal method As String, ByVal url As String) As String
    
    Dim timestamp As Long
    Dim oauth_timestamp As String
    Dim oauth_nonce As String
    
    timestamp = DateDiff("s", #1/1/1970#, Now)
    
    oauth_timestamp = CStr(timestamp)
    oauth_nonce = CStr(timestamp + 1)
    
    authorization = "OAuth " & _
                    "oauth_consumer_key=""" & oauth_consumer_key & """, " & _
                    "oauth_nonce=""" & oauth_nonce & """," & _
                    "oauth_signature=""" & HMAC_SHA1(sourceEncodeUrl(method, url, oauth_timestamp, oauth_nonce)) & """," & _
                    "oauth_signature_method=""HMAC-SHA1""," & _
                    "oauth_timestamp=""" & oauth_timestamp & """," & _
                    "oauth_token=""" & oauth_token & """, " & _
                    "oauth_version = ""1.0"""

End Function

Private Function HMAC_SHA1(ByVal sourceEncodeUrl As String) As String
                            
    Dim HMACSHA1 As New HMACSHA1
    Dim msxml2Doc As New MSXML2.DOMDocument60
    Dim msxml2Ele As MSXML2.IXMLDOMElement
    
    Dim keys() As Byte
    Dim bytes() As Byte
    Dim hash_sha1() As Byte
    
    keys = StrConv(signinKey, vbFromUnicode)
    HMACSHA1.Key = keys
    
    bytes = StrConv(sourceEncodeUrl, vbFromUnicode)
    hash_sha1 = HMACSHA1.ComputeHash_2(bytes)
    
    Set msxml2Ele = msxml2Doc.createElement("b64")
    
    msxml2Ele.DataType = "bin.base64"
    msxml2Ele.nodeTypedValue = hash_sha1
    
    HMAC_SHA1 = WorksheetFunction.EncodeURL(msxml2Ele.text)
    
End Function

Private Function responseJson(ByVal url As String, Optional ByVal method As String = "GET", Optional ByVal requestJson As Dictionary, Optional ByVal queryParameter As String = "") As Object

    Dim xmlHttp As New MSXML2.XMLHTTP60
    
On Error GoTo sendError
    
    With xmlHttp
        
        'ObjectやVariant型にすると何故か初期値がセットされないので、されなかった場合の考慮
        If method = "" Then method = "GET"
        
        If queryParameter = "" Then
            .Open method, url, False
        Else
            '.Open method, url & "?" & WorksheetFunction.EncodeURL(queryParameter), False
            .Open method, url & "?" & queryParameter, False
            Debug.Print url & "?" & queryParameter
        End If
        
        If method = "POST" Then
            .setRequestHeader "Content-Type", "application/json"
        Else
            .setRequestHeader "Content-Type", "text/json"
        End If
        
        If queryParameter = "" Then
            .setRequestHeader "Authorization", authorization(method, url)
        Else
            .setRequestHeader "Authorization", "Bearer " & bearer_token
        End If
        
        If requestJson Is Nothing Then
            .send
        Else
            .send Replace(JsonConverter.ConvertToJson(requestJson), "|", "")
        End If
        
        Set responseJson = JsonConverter.ParseJson(.responseText)("data")
        Debug.Print .responseText
    End With
    
    Exit Function

sendError:

    MsgBox xmlHttp.responseText, vbExclamation
    End
    
End Function

