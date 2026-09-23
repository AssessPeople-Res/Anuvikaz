<!-- #include file ="CaptchaClass.asp" -->

<%

' basic request validation
Dim command : command = LBD_Normalize(Request("get"))
If command = "" Or Len(command) < 4 Then
  Call BadRequest("command " & Len(command))
End If

Dim captchaId : captchaId = LBD_Normalize(Request("c"))
If captchaId = "" Or Len(captchaId) = 0 Then
  Call BadRequest("c " & Len(captchaId))
End If

Dim instanceId : instanceId = LBD_Normalize(Request("t"))
If instanceId = "" Or Len(instanceId) <> 32 Then
  Call BadRequest("t " & Len(instanceId))
End If

'GbPlugin workaround, see http://captcha.com/doc/asp/known-issues.html#gbplugin
Dim userAgent : userAgent = Request.ServerVariables("HTTP_USER_AGENT")
Dim gbpluginRegExp : Set gbpluginRegExp = New RegExp
With gbpluginRegExp
  .Global = True
  .IgnoreCase = True
  .Pattern = "(^GbPlugin$)"
End With
If (gbpluginRegExp.Test(userAgent) = True) Then
  Call IgnoreRequest
End If
Set gbpluginRegExp = Nothing


'Captcha object
Dim captchaInstance : Set captchaInstance = (New Captcha)(Array(captchaId, instanceId))
Dim binaryContent



If (command="image") Then
'Captcha image generation

  binaryContent = captchaInstance.GenerateImage
  Response.ContentType = captchaInstance.ImageMimeType
  Response.Buffer = True
  Response.CacheControl = "no-cache, no-store, must-revalidate"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  Response.AddHeader "Connection", "Close"
  Response.AddHeader "Accept-Ranges", "none"
  Response.AddHeader "Server", "Apache/2.2.17 (Win32) mod_ssl/2.2.17 OpenSSL/0.9.8o PHP/5.3.4 mod_perl/2.0.4 Perl/v5.10.1"
  Response.AddHeader "X-Robots-Tag", "noindex, nofollow, noarchive, nosnippet"
  Response.BinaryWrite binaryContent
  Response.End
'end Captcha image generation



ElseIf (command="sound") Then
'audio Captcha generation

  ' iPhone/iPad sound issues workaround:
  ' we need a completely different Captcha sound
  ' Http workflow
  If (Detect_iOS_ChunkedRequest()) Then
  
    ' detect the first request after a new sound icon click, and clear any stored data
    ' to avoid reusing the same sound endlessly within a session
    ' javascript player adds a timestamp querystring param ("&d=..."), so it can be detected by it
    If (Request("d") <> "") Then
      ' when javascript is enabled, we can detect the first request because the timestamp changed
      Dim soundClickId : soundClickId = LBD_Normalize(Request("d"))
      Dim prevSoundClickId : prevSoundClickId = LBD_Persistence("prevSoundClickId")
      If (0 <> StrComp(soundClickId, prevSoundClickId, 1)) Then
        Call Clear_iOS_SoundData()
        LBD_Persistence("prevSoundClickId") = soundClickId ' on first request, save for future checks
      End If
    End If 

    ' chunked requests must include the desired byte range
    Dim byteRange : byteRange = GetSoundByteRange()
    Dim rangeStart : rangeStart = byteRange(0)
    Dim rangeEnd : rangeEnd = byteRange(1)
    Dim rangeSize : rangeSize = rangeEnd - rangeStart + 1
    
    Dim soundBytes : soundBytes = Get_iOS_SoundData(captchaInstance)
    Dim totalSize : totalSize = lenB(Cstr(soundBytes))

    ' initial iOS 6.0.1 testing; leaving as fallback since we can't be sure it won't happen again:
    ' we depend on observed behavior of invalid range requests to detect
    ' end of sound playback, cleanup and tell AppleCoreMedia to stop requesting
    ' invalid "bytes=rangeEnd-rangeEnd" ranges in an infinite(?) loop
    If (rangeEnd <= rangeStart Or rangeEnd > totalSize) Then
      Call Clear_iOS_SoundData()
      Call BadRequest("invalid byte range " & rangeStart &"-" & rangeEnd & "/" & totalSize)
    End If
    
    Dim rangeBytes : rangeBytes = MidB(Cstr(soundBytes), rangeStart + 1, rangeSize)
    
    ' partial content response with the requested byte range
    Response.Status = "206 Partial Content"
    Response.ContentType = captchaInstance.SoundMimeType
    If (Request("d") = "") Then
      Response.AddHeader "content-disposition", _
        "attachment; filename=" & captchaInstance.SoundFilename
    End If
    Response.Buffer = True
    Response.AddHeader "Content-Transfer-Encoding", "binary"
    Response.AddHeader "X-Robots-Tag", "noindex, nofollow, noarchive, nosnippet"
    Response.AddHeader "Accept-Ranges", "bytes"
    'Response.AddHeader "Content-Length", rangeSize
    Response.AddHeader "Content-Range", "bytes " & rangeStart &"-" & rangeEnd & "/" & totalSize
    If ((Request.ServerVariables("HTTPS")="off") Or Request("e")="") Then
      Response.CacheControl = "no-cache, no-store, must-revalidate"
      Response.AddHeader "Pragma", "no-cache"
      Response.Expires = -1
    End If
    Response.BinaryWrite rangeBytes
    Response.End
  Else
    ' when javascript is disabled in the iOS browser, it won't be
    ' detected using the above check; so this is the only point where
    ' we can detect the first request after the sound icon is clicked
    ' in the iOS browser with javascript disabled
    Call Clear_iOS_SoundData()
  
    ' normal sound generation
    binaryContent = captchaInstance.GenerateSound

    If Not IsEmpty(binaryContent) Then
      Response.ContentType = captchaInstance.SoundMimeType
      If (Request("d") = "") Then
        Response.AddHeader "content-disposition", _
          "attachment; filename=" & captchaInstance.SoundFilename
      End If
      Response.Buffer = True
      Response.AddHeader "Content-Transfer-Encoding", "binary"
      Response.AddHeader "Accept-Ranges", "none"
      Response.AddHeader "X-Robots-Tag", "noindex, nofollow, noarchive, nosnippet"
      If ((Request.ServerVariables("HTTPS")="off") Or Request("e")="") Then
        Response.CacheControl = "no-cache, no-store, must-revalidate"
        Response.AddHeader "Pragma", "no-cache"
        Response.Expires = -1
      End If
      Response.BinaryWrite binaryContent
      Response.End
    Else
      Call BadRequest("no code")
    End If
  End If
'end audio Captcha generation



ElseIf (command="validationresult") Then
'Ajax Captcha validation

  'BotDetect built-in Ajax Captcha validation
  Dim userInput : userInput = LBD_Normalize(Request("i"))
  
  If userInput = "" Or Len(userInput) < 1 Then
    'jQuery validation support, the input key may be just about anything,
    'so we have to loop through fields and take the first unrecognized one
    For Each field in Request.Querystring
      If (field <> "get" And field <> "c" And field <> "t") Then
        userInput = LBD_Normalize(Request(field))
      End If
    Next
  End If
  
  Dim result
  If (userInput = "" Or Len(userInput) < 1) Then
    result = False
  Else
    result = captchaInstance.AjaxValidate(instanceId, userInput)
  End If
  
  'Http response headers
  Response.Buffer = True
  Response.ContentType = "application/json"
  Response.CacheControl = "no-cache, no-store, must-revalidate"
  Response.AddHeader "Pragma", "no-cache"
  Response.Expires = -1
  Response.AddHeader "Connection", "Close"
  Response.AddHeader "X-Robots-Tag", "noindex, nofollow, noarchive, nosnippet"

  'send the JSON validation result to the client
  Response.Write LCase(CStr(result))
  Response.End

'end Ajax Captcha validation
End If

'If neither of the above conditions was met
Call BadRequest("no command")



' Instead of relying on unreliable user agent checks, we detect the iOS sound
' requests by the Http headers they will always contain
Function Detect_iOS_ChunkedRequest()
  Detect_iOS_ChunkedRequest = False
  Dim playbackSessionId : playbackSessionId = Request.ServerVariables("HTTP_X_PLAYBACK_SESSION_ID")
  Dim byteRangeStr : byteRangeStr = Request.ServerVariables("HTTP_RANGE")
  If (playbackSessionId <> "" And Len(playbackSessionId) > 4 And byteRangeStr <> "" And Len(byteRangeStr) > 4)Then
    Detect_iOS_ChunkedRequest = True
  End If
End Function


' chunked requests must include the desired byte range
Function GetSoundByteRange()
  ' chunked requests must include the desired byte range
  Dim rangeStr : rangeStr = Request.ServerVariables("HTTP_RANGE")
  If rangeStr = "" Or Len(rangeStr) < 4 Then
    Call BadRequest("rangeStr")
  End If
  
  ' sound byte subset
  Dim matches
  Dim rangeRegExp : Set rangeRegExp = New RegExp
  rangeRegExp.Global = False
  rangeRegExp.IgnoreCase = True
  rangeRegExp.Pattern = "bytes=([0-9]+)-([0-9]+)"
  Set matches = rangeRegExp.Execute(rangeStr)
  
  Dim count : count = 0
  Dim count2 : count2 = 0
  For Each match in matches
      For Each submatch in match.SubMatches
        If (0 = count And 0 = count2) Then
          rangeStart = CLng(submatch)
        End If
        If (0 = count And 1 = count2) Then
          rangeEnd = CLng(submatch)
        End If
        count2 = count2 + 1
      Next
      count = count + 1
  Next
  
  Dim rangeResult(2)
  rangeResult(0) = rangeStart
  rangeResult(1) = rangeEnd
  
  Set rangeRegExp = Nothing
  Set matches = Nothing
  
  GetSoundByteRange = rangeResult
End Function


Function Get_iOS_SoundData(captchaInstance)
  ' since we need to keep the same Captcha sound across all chunked
  ' requests, it is only generated when a new (non-duplicate) playback 
  ' session is started in a request
  Dim soundBytes : soundBytes = Load_iOS_SoundData()
  If (Not IsArray(soundBytes)) Then
    soundBytes = captchaInstance.GenerateSound
    Call Save_iOS_SoundData(soundBytes)
  End If

  Get_iOS_SoundData = soundBytes
End Function



' we persist the chunked sound across requests in a Session variable
Sub Save_iOS_SoundData(soundBytes)
  LBD_Persistence("Cached_iOS_Sound") = soundBytes
End Sub

Function Load_iOS_SoundData()
  Load_iOS_SoundData = LBD_Persistence("Cached_iOS_Sound")
End Function

Sub Clear_iOS_SoundData()
  LBD_Persistence.Contents.Remove("Cached_iOS_Sound")
End Sub


' Called when Captcha request validation fails
Sub BadRequest(message)
  Response.Status = "200 OK"
  Response.ContentType = "text/plain"
  Response.Write message & vbCrlf
  Response.AddHeader "X-Robots-Tag", "noindex, nofollow, noarchive, nosnippet"
  Response.Flush
  Response.End
End Sub

' Called when Captcha requests should be ignored
Sub IgnoreRequest()
  Response.Status = "200 OK"
  Response.ContentType = "text/plain"
  Response.Write "OK" & vbCrlf
  Response.AddHeader "X-Robots-Tag", "noindex, nofollow, noarchive, nosnippet"
  Response.End
End Sub
%>