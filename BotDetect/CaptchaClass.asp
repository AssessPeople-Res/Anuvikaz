<!-- #include file ="Utils.asp" -->
<!-- #include file ="CaptchaConfig.asp" -->

<%

' Instances of the Captcha type can be created on your ASP forms and will take
'care of Captcha display and validation.
Class Captcha

  'Unique identifier of the Captcha object within the application (for example,
  'if you placed one Captcha object on the Registration page and another on
  'the Contact Us page, they would have distinct CaptchaId values).
  Private m_CaptchaID
  Public Property Get CaptchaID()
    CaptchaID = m_CaptchaID
  End Property

  Private m_UserSpecifiedCaptchaID
  
  'Globally unique identifier of the current Captcha object instance, used
  'to ensure each page load keeps separate Captcha codes, for example
  'when opening the same form in multiple browser tabs.
  Private m_InstanceID
  Public Property Get InstanceID()
    InstanceID = m_InstanceID
  End Property


  ' constructor
  Private Sub Class_Initialize()
    'create the Captcha component instance
    Set m_ComCaptcha = CreateObject("BotDetect.Captcha.3")
  End Sub

  'Instead of using the constructor, Captcha objects should always be created
  'using the default Init function, which takes the Captcha identifier (a
  'string) as parameter and automatically initializes the instance identifier.
  '
  'The same function can also be used to create a Captcha VbScript object
  'instance with the given Captcha identifier and instance identifier (sent as
  'an Array). This is used to reference data for a previous object instance,
  'for example when generating Captcha images and sounds in the
  'BotDetect\CaptchaHandler.asp source.
  '
  'This "overload" of the Init function is not used directly when implementing
  'the default Captcha library behavior, only when making significant
  'customizations in the BotDetect source.
  Public Default Function Init(p_Value)

    If IsArray(p_Value) Then
      m_CaptchaID = LCase(LBD_Normalize(CStr(p_Value(0))))
      m_UserSpecifiedCaptchaID = CStr(p_Value(0))
      m_InstanceID = LCase(LBD_Normalize(CStr(p_Value(1))))
    Else
      m_CaptchaID = LCase(LBD_Normalize(CStr(p_Value)))
      m_UserSpecifiedCaptchaID = CStr(p_Value)
      m_InstanceID = LCase(LBD_CreateGuid())
    End If

    m_ImageTooltip = LBD_Configuration_ImageTooltip

    m_SoundEnabled = LBD_Configuration_SoundEnabled
    m_SoundTooltip = LBD_Configuration_SoundTooltip
    m_SoundIconUrl = LBD_Configuration_SoundIconUrl

    m_ReloadEnabled = LBD_Configuration_ReloadEnabled
    m_ReloadTooltip = LBD_Configuration_ReloadTooltip
    m_ReloadIconUrl = LBD_Configuration_ReloadIconUrl

    ' persistence keys
    m_LocaleKey = "LBD_Locale_" & m_CaptchaID

    m_CodeLengthKey = "LBD_CodeLength_" & m_CaptchaID
    m_CodeStyleKey = "LBD_CodeStyle_" & m_CaptchaID
    m_CodeTimeoutKey = "LBD_CodeTimeout_" & m_CaptchaID

    m_ImageStyleKey = "LBD_ImageStyle_" & m_CaptchaID
    m_ImageWidthKey = "LBD_ImageWidth_" & m_CaptchaID
    m_ImageHeightKey = "LBD_ImageHeight_" & m_CaptchaID
    m_ImageFormatKey = "LBD_ImageFormat_" & m_CaptchaID
    m_CustomDarkColorKey = "LBD_CustomDarkColor_" & m_CaptchaID
    m_CustomLightColorKey = "LBD_CustomLightColor_" & m_CaptchaID

    m_SoundStyleKey = "LBD_SoundStyle_" & m_CaptchaID
    m_SoundFormatKey = "LBD_SoundFormat_" & m_CaptchaID

    m_IsSolvedKey = "LBD_IsSolved_" & m_CaptchaID

    m_HiddenFieldID = "LBD_VCID_" & m_CaptchaID

    m_TabIndexStart = -255
    
    m_HelpLinkEnabled = LBD_Configuration_HelpLinkEnabled
    m_HelpLinkMode = LBD_Configuration_HelpLinkMode
    m_HelpLinkUrl = LBD_Configuration_HelpLinkUrl
    m_HelpLinkText = LBD_Configuration_HelpLinkText
    
    Me.Load

    Set Init = Me
  End Function

  
  ' destructor
  Private Sub Class_Terminate()
    ' dispose of the COM instance
    Set m_ComCaptcha = Nothing
  End Sub



  'Locale string, affects the character set used for Captcha code generation
  'and the pronunciation language used for Captcha sound generation
  Private m_Locale
  Public Property Get Locale()
    Locale = m_Locale
  End Property
  Public Property Let Locale(p_Value)
    m_Locale = CStr(p_Value)
    m_IsLocaleModified = True
  End Property
  Private m_LocaleKey

  'Validates the given Locale string, checking it's not empty or too short to
  'be a valid locale specification.
  'Since the locale strings supported by BotDetect will change depending on
  'the underlying COM component version and sound package files available,
  'this function doesn't perform any comprehensive locale checks.
  Public Function IsLocaleValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = Trim(CStr(p_Value))
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (Len(value) > 1) Then
      result = True
    End If
    IsLocaleValid = result
  End Function

  Private m_IsLocaleModified
  Private m_IsLocaleRestored
  Private m_IsLocaleRemembered
  Private m_IsLocaleReset

  Public Sub LoadLocale
    m_Locale = LBD_Configuration_Locale

    m_IsLocaleModified = False
    m_IsLocaleRestored = False
    m_IsLocaleRemembered = False
    m_IsLocaleReset = False

    Dim savedLocale : savedLocale = LBD_Persistence(m_LocaleKey)

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback And IsEmpty(savedLocale) Then
      savedLocale = Application(m_LocaleKey)
    End If

    If Not IsEmpty(savedLocale) And IsLocaleValid(savedLocale) Then
      m_Locale = savedLocale
      m_IsLocaleRestored = True
    End If
  End Sub

  Public Sub SaveLocale
    If m_IsLocaleModified Then
      If IsLocaleValid(m_Locale) Then
        Me.RememberLocale
      Else
        Me.ResetLocale
      End If
    End If
  End Sub

  Public Sub RememberLocale
    LBD_Persistence(m_LocaleKey) = m_Locale
    m_IsLocaleRemembered = True

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback Then
      Application.Lock
      Application(m_LocaleKey) = m_Locale
      Application.Unlock
    End If
  End Sub

  Public Sub ResetLocale
    LBD_Persistence(m_LocaleKey) = Empty
    m_Locale = LBD_Configuration_Locale
    m_IsLocaleReset = True

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback Then
      Application.Lock
      Application(m_LocaleKey) = Empty
      Application.Unlock
    End If
  End Sub


  'Length (number of characters) of the Captcha code rendered; the default
  'value is 5.
  Private m_CodeLength
  Public Property Get CodeLength()
    CodeLength = m_CodeLength
  End Property
  Public Property Let CodeLength(p_Value)
    m_CodeLength = CLng(p_Value)
    m_IsCodeLengthModified = True
  End Property
  Private m_CodeLengthKey

  'Validates the given code length, checking that it's a number between 1 and 15.
  Public Function IsCodeLengthValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = CInt(p_Value)
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (value >= 1 And value <= 15) Then
      result = True
    End If
    IsCodeLengthValid = result
  End Function

  Private m_IsCodeLengthModified
  Private m_IsCodeLengthRestored
  Private m_IsCodeLengthRemembered
  Private m_IsCodeLengthReset

  Public Sub LoadCodeLength
    m_CodeLength = LBD_Configuration_CodeLength

    m_IsCodeLengthModified = False
    m_IsCodeLengthRestored = False
    m_IsCodeLengthRemembered = False
    m_IsCodeLengthReset = False

    Dim savedCodeLength : savedCodeLength = LBD_Persistence(m_CodeLengthKey)
    If Not IsEmpty(savedCodeLength) And IsCodeLengthValid(savedCodeLength) Then
      m_CodeLength = savedCodeLength
      m_IsCodeLengthRestored = True
    End If
  End Sub

  Public Sub SaveCodeLength
    If m_IsCodeLengthModified Then
      If IsCodeLengthValid(m_CodeLength) Then
        Me.RememberCodeLength
      Else
        Me.ResetCodeLength
      End If
    End If
  End Sub

  Public Sub RememberCodeLength
    LBD_Persistence(m_CodeLengthKey) = m_CodeLength
    m_IsCodeLengthRemembered = True
  End Sub

  Public Sub ResetCodeLength
    LBD_Persistence(m_CodeLengthKey) = Empty
    m_CodeLength = LBD_Configuration_CodeLength
    m_IsCodeLengthReset = True
  End Sub


  'Numeric representation of the Captcha code style, i.e. the algorithm used
  'to generate Captcha codes from default pre-defined locale character sets;
  'the default value is 0.
  'This value is persisted if needed and directly propagated to the COM
  'Captcha component interface.
  Private m_CodeStyle
  Public Property Get CodeStyle()
    CodeStyle = m_CodeStyle
  End Property
  Public Property Let CodeStyle(p_Value)
    m_CodeStyle = CLng(p_Value)
    m_IsCodeStyleModified = True
  End Property

  'String representation of the Captcha code style, i.e. the algorithm used
  'to generate Captcha codes from default pre-defined locale character sets;
  'the default value is "Alphanumeric".
  'This value is converted to its numerical representation before it's actually
  'used, so the string property is just a helper to make your ASP code
  'more readable.
  Public Property Get CodeStyleName()
    CodeStyleName = LBD_CodeStyleNames(m_CodeStyle)
  End Property
  Public Property Let CodeStyleName(p_Value)
    m_CodeStyle = CLng(LBD_CodeStyles(LCase(p_Value)))
    m_IsCodeStyleModified = True
  End Property
  Private m_CodeStyleKey

  'Validates the given numeric code style value, checking that it's a number
  'between 0 and 2.
  Public Function IsCodeStyleValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = CInt(p_Value)
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (value >= 0 And value <= 2) Then
      result = True
    End If
    IsCodeStyleValid = result
  End Function

  Private m_IsCodeStyleModified
  Private m_IsCodeStyleRestored
  Private m_IsCodeStyleRemembered
  Private m_IsCodeStyleReset

  Public Sub LoadCodeStyle
    m_CodeStyle = LBD_Configuration_CodeStyle

    m_IsCodeStyleModified = False
    m_IsCodeStyleRestored = False
    m_IsCodeStyleRemembered = False
    m_IsCodeStyleReset = False

    Dim savedCodeStyle : savedCodeStyle = LBD_Persistence(m_CodeStyleKey)
    If Not IsEmpty(savedCodeStyle) And IsCodeStyleValid(savedCodeStyle) Then
      m_CodeStyle = savedCodeStyle
      m_IsCodeStyleRestored = True
    End If
  End Sub

  Public Sub SaveCodeStyle
    If m_IsCodeStyleModified Then
      If IsCodeStyleValid(m_CodeStyle) Then
        Me.RememberCodeStyle
      Else
        Me.ResetCodeStyle
      End If
    End If
  End Sub

  Public Sub RememberCodeStyle
    LBD_Persistence(m_CodeStyleKey) = m_CodeStyle
    m_IsCodeStyleRemembered = True
  End Sub

  Public Sub ResetCodeStyle
    LBD_Persistence(m_CodeStyleKey) = Empty
    m_CodeStyle = LBD_Configuration_CodeStyle
    m_IsCodeStyleReset = True
  End Sub


  'Period (in seconds) during which each generated Captcha code can be
  'validated after generation; the default value is 1200 seconds (20 minutes).
  Private m_CodeTimeout
  Public Property Get CodeTimeout()
    CodeTimeout = m_CodeTimeout
  End Property
  Public Property Let CodeTimeout(p_Value)
    m_CodeTimeout = CLng(p_Value)
    m_IsCodeTimeoutModified = True
  End Property
  Private m_CodeTimeoutKey

  'Validates the given code timeout, checking that it's a number between 30 and
  '1200.
  'Code timeouts longer than 20 minutes are considered invalid since the default
  'Classic ASP Session timeout is 20 minutes, and setting a longer code timeout
  'wouldn't work without increasing the Session timeout first.
  'If you increase your server's ASP Session state timeout, you can also change
  'the validation function upper bound.
  Public Function IsCodeTimeoutValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = CInt(p_Value)
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (value >= 30 And value <= 1200) Then
      result = True
    End If
    IsCodeTimeoutValid = result
  End Function

  Private m_IsCodeTimeoutModified
  Private m_IsCodeTimeoutRestored
  Private m_IsCodeTimeoutRemembered
  Private m_IsCodeTimeoutReset

  Public Sub LoadCodeTimeout
    m_CodeTimeout = LBD_Configuration_CodeTimeout

    m_IsCodeTimeoutModified = False
    m_IsCodeTimeoutRestored = False
    m_IsCodeTimeoutRemembered = False
    m_IsCodeTimeoutReset = False

    Dim savedCodeTimeout : savedCodeTimeout = LBD_Persistence(m_CodeTimeoutKey)
    If Not IsEmpty(savedCodeTimeout) And IsCodeTimeoutValid(savedCodeTimeout) Then
      m_CodeTimeout = savedCodeTimeout
      m_IsCodeTimeoutRestored = True
    End If
  End Sub

  Public Sub SaveCodeTimeout
    If m_IsCodeTimeoutModified Then
      If IsCodeTimeoutValid(m_CodeTimeout) Then
        Me.RememberCodeTimeout
      Else
        Me.ResetCodeTimeout
      End If
    End If
  End Sub

  Public Sub RememberCodeTimeout
    LBD_Persistence(m_CodeTimeoutKey) = m_CodeTimeout
    m_IsCodeTimeoutRemembered = True
  End Sub

  Public Sub ResetCodeTimeout
    LBD_Persistence(m_CodeTimeoutKey) = Empty
    m_CodeTimeout = LBD_Configuration_CodeTimeout
    m_IsCodeTimeoutReset = True
  End Sub


  'Numeric representation of the Captcha image style, i.e. the algorithm
  'used to render Captcha codes in images; the default image style is 31.
  'This value is persisted if needed and directly propagated to the COM
  'Captcha component interface.
  Private m_ImageStyle
  Public Property Get ImageStyle()
    ImageStyle = m_ImageStyle
  End Property
  Public Property Let ImageStyle(p_Value)
    m_ImageStyle = CLng(p_Value)
    m_IsImageStyleModified = True
  End Property

  'String representation of the Captcha image style, i.e. the algorithm used
  'to render Captcha codes in images; the default image style is "Chalkboard"
  'This value is converted to its numerical representation before it's actually
  'used, so the string property is just a helper to make your ASP code more
  'readable.
  Public Property Get ImageStyleName()
    ImageStyleName = LBD_ImageStyleNames(m_ImageStyle)
  End Property
  Public Property Let ImageStyleName(p_Value)
    m_ImageStyle = LBD_ImageStyles(LCase(p_Value))
    m_IsImageStyleModified = True
  End Property
  Private m_ImageStyleKey

  'Validates the given numeric image style value, checking that it's a number between 0 and 49.
  Public Function IsImageStyleValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = CInt(p_Value)
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (value >= 0 And value <= 59) Then
      result = True
    End If
    IsImageStyleValid = result
  End Function

  Private m_IsImageStyleModified
  Private m_IsImageStyleRestored
  Private m_IsImageStyleRemembered
  Private m_IsImageStyleReset

  Public Sub LoadImageStyle
    m_ImageStyle = LBD_Configuration_ImageStyle

    m_IsImageStyleModified = False
    m_IsImageStyleRestored = False
    m_IsImageStyleRemembered = False
    m_IsImageStyleReset = False

    Dim savedImageStyle : savedImageStyle = LBD_Persistence(m_ImageStyleKey)
    If Not IsEmpty(savedImageStyle) And IsImageStyleValid(savedImageStyle) Then
      m_ImageStyle = savedImageStyle
      m_IsImageStyleRestored = True
    End If
  End Sub

  Public Sub SaveImageStyle
    If m_IsImageStyleModified Then
      If IsImageStyleValid(m_ImageStyle) Then
        Me.RememberImageStyle
      Else
        Me.ResetImageStyle
      End If
    End If
  End Sub

  Public Sub RememberImageStyle
    LBD_Persistence(m_ImageStyleKey) = m_ImageStyle
    m_IsImageStyleRemembered = True
  End Sub

  Public Sub ResetImageStyle
    LBD_Persistence(m_ImageStyleKey) = Empty
    m_ImageStyle = LBD_Configuration_ImageStyle
    m_IsImageStyleReset = True
  End Sub


  ' Width of the generate Captcha image, in pixels; the default Captcha image
  'width is 250 px.
  Private m_ImageWidth
  Public Property Get ImageWidth()
    ImageWidth = m_ImageWidth
  End Property
  Public Property Let ImageWidth(p_Value)
    m_ImageWidth = CLng(p_Value)
    m_IsImageWidthModified = True
  End Property
  Private m_ImageWidthKey

  'Validates the given image width value, checking that it's a number between
  '20 and 1000.
  'Captcha images smaller than 20 x 20 pixels are simply not supported, and
  'the larger the image, the longer it takes to generate it and transfer it to
  'the client.
  Public Function IsImageWidthValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = CInt(p_Value)
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (value >= 20 And value <= 1000) Then
      result = True
    End If
    IsImageWidthValid = result
  End Function

  Private m_IsImageWidthModified
  Private m_IsImageWidthRestored
  Private m_IsImageWidthRemembered
  Private m_IsImageWidthReset

  Public Sub LoadImageWidth
    m_ImageWidth = LBD_Configuration_ImageWidth

    m_IsImageWidthModified = False
    m_IsImageWidthRestored = False
    m_IsImageWidthRemembered = False
    m_IsImageWidthReset = False

    Dim savedImageWidth : savedImageWidth = LBD_Persistence(m_ImageWidthKey)
    If Not IsEmpty(savedImageWidth) And IsImageWidthValid(savedImageWidth) Then
      m_ImageWidth = savedImageWidth
      m_IsImageWidthRestored = True
    End If
  End Sub

  Public Sub SaveImageWidth
    If m_IsImageWidthModified Then
      If IsImageWidthValid(m_ImageWidth) Then
        Me.RememberImageWidth
      Else
        Me.ResetImageWidth
      End If
    End If
  End Sub

  Public Sub RememberImageWidth
    LBD_Persistence(m_ImageWidthKey) = m_ImageWidth
    m_IsImageWidthRemembered = True
  End Sub

  Public Sub ResetImageWidth
    LBD_Persistence(m_ImageWidthKey) = Empty
    m_ImageWidth = LBD_Configuration_ImageWidth
    m_IsImageWidthReset = True
  End Sub


  'Height of the generate Captcha image, in pixels; the default Captcha image
  'height is 50 px.
  Private m_ImageHeight
  Public Property Get ImageHeight()
    ImageHeight = m_ImageHeight
  End Property
  Public Property Let ImageHeight(p_Value)
    m_ImageHeight = CLng(p_Value)
    m_IsImageHeightModified = True
  End Property
  
  
  Private m_ImageHeightKey

  'Validates the given image height value, checking that it's a number between
  '20 and 1000.
  'Captcha images smaller than 20 x 20 pixels are simply not supported, and
  'the larger the image, the longer it takes to generate it and transfer it
  'to the client.
  Public Function IsImageHeightValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = CInt(p_Value)
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (value >= 20 And value <= 1000) Then
      result = True
    End If
    IsImageHeightValid = result
  End Function

  Private m_IsImageHeightModified
  Private m_IsImageHeightRestored
  Private m_IsImageHeightRemembered
  Private m_IsImageHeightReset

  Public Sub LoadImageHeight
    m_ImageHeight = LBD_Configuration_ImageHeight

    m_IsImageHeightModified = False
    m_IsImageHeightRestored = False
    m_IsImageHeightRemembered = False
    m_IsImageHeightReset = False

    Dim savedImageHeight : savedImageHeight = LBD_Persistence(m_ImageHeightKey)
    If Not IsEmpty(savedImageHeight) And IsImageHeightValid(savedImageHeight) Then
      m_ImageHeight = savedImageHeight
      m_IsImageHeightRestored = True
    End If
  End Sub

  Public Sub SaveImageHeight
    If m_IsImageHeightModified Then
      If IsImageHeightValid(m_ImageHeight) Then
        Me.RememberImageHeight
      Else
        Me.ResetImageHeight
      End If
    End If
  End Sub

  Public Sub RememberImageHeight
    LBD_Persistence(m_ImageHeightKey) = m_ImageHeight
    m_IsImageHeightRemembered = True
  End Sub

  Public Sub ResetImageHeight
    LBD_Persistence(m_ImageHeightKey) = Empty
    m_ImageHeight = LBD_Configuration_ImageHeight
    m_IsImageHeightReset = True
  End Sub


  'Numeric representation of the image format in which the Captcha image will
  'be generated; the default format is 0.
  'This value is persisted if needed and directly propagated to the COM Captcha
  'component interface.
  Private m_ImageFormat
  Public Property Get ImageFormat()
    ImageFormat = m_ImageFormat
  End Property
  Public Property Let ImageFormat(p_Value)
    m_ImageFormat = CLng(p_Value)
    m_IsImageFormatModified = True
  End Property

  'String representation of the image format in which the Captcha image will
  'be generated; the default format is "JPEG"
  'This value is converted to its numerical representation before it's actually
  'used, so the string property is just a helper to make your ASP code more
  'readable.
  Public Property Get ImageFormatName()
    ImageFormatName = LBD_ImageFormatNames(m_ImageFormat)
  End Property
  Public Property Let ImageFormatName(p_Value)
    If LCase(p_Value) = "jpg" Then
      p_Value = "jpeg"
    End If
    m_ImageFormat = CLng(LBD_ImageFormats(LCase(p_Value)))
    m_IsImageFormatModified = True
  End Property
  Private m_ImageFormatKey

  'Validates the given numeric code style value, checking that it's a number
  'between 0 and 3.
  Public Function IsImageFormatValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = CInt(p_Value)
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (value >= 0 And value <= 3) Then
      result = True
    End If
    IsImageFormatValid = result
  End Function

  Private m_IsImageFormatModified
  Private m_IsImageFormatRestored
  Private m_IsImageFormatRemembered
  Private m_IsImageFormatReset

  Public Sub LoadImageFormat
    m_ImageFormat = LBD_Configuration_ImageFormat

    m_IsImageFormatModified = False
    m_IsImageFormatRestored = False
    m_IsImageFormatRemembered = False
    m_IsImageFormatReset = False

    Dim savedImageFormat : savedImageFormat = LBD_Persistence(m_ImageFormatKey)
    If Not IsEmpty(savedImageFormat) And IsImageFormatValid(savedImageFormat) Then
      m_ImageFormat = savedImageFormat
      m_IsImageFormatRestored = True
    End If
  End Sub

  Public Sub SaveImageFormat
    If m_IsImageFormatModified Then
      If IsImageFormatValid(m_ImageFormat) Then
        Me.RememberImageFormat
      Else
        Me.ResetImageFormat
      End If
    End If
  End Sub

  Public Sub RememberImageFormat
    LBD_Persistence(m_ImageFormatKey) = m_ImageFormat
    m_IsImageFormatRemembered = True
  End Sub

  Public Sub ResetImageFormat
    LBD_Persistence(m_ImageFormatKey) = Empty
    m_ImageFormat = LBD_Configuration_ImageFormat
    m_IsImageFormatReset = True
  End Sub


  'Optional custom dark color point, modifies the color palette used for
  'Captcha image drawing. Should be a Html color, i.e. specified either by
  'name ("Red") or hex value ("#f00" or "#ff0000")
  Private m_CustomDarkColor
  Public Property Get CustomDarkColor()
    If (m_CustomDarkColor <> "") Then
      CustomDarkColor = m_CustomDarkColor
    Else
      CustomDarkColor = "Not Set"
    End If
  End Property
  Public Property Let CustomDarkColor(p_Value)
    m_CustomDarkColor = CStr(p_Value)
    m_IsCustomDarkColorModified = True
  End Property
  Private m_CustomDarkColorKey

  'Validates the given string custom dark color value, checking that it's not
  'an empty string.
  'The Classic ASP Captcha library doesn't check that the given value is a
  'valid Html color value, but the underlying COM implementation will throw
  'an error if an invalid value is specified.
  Public Function IsCustomDarkColorValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = Trim(CStr(p_Value))
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (Len(value) > 1) Then
      result = True
    End If
    IsCustomDarkColorValid = result
  End Function

  Private m_IsCustomDarkColorModified
  Private m_IsCustomDarkColorRestored
  Private m_IsCustomDarkColorRemembered
  Private m_IsCustomDarkColorReset

  Public Sub LoadCustomDarkColor
    m_CustomDarkColor = LBD_Configuration_CustomDarkColor

    m_IsCustomDarkColorModified = False
    m_IsCustomDarkColorRestored = False
    m_IsCustomDarkColorRemembered = False
    m_IsCustomDarkColorReset = False

    Dim savedCustomDarkColor : savedCustomDarkColor = LBD_Persistence(m_CustomDarkColorKey)
    If Not IsEmpty(savedCustomDarkColor) And IsCustomDarkColorValid(savedCustomDarkColor) Then
      m_CustomDarkColor = savedCustomDarkColor
      m_IsCustomDarkColorRestored = True
    End If
  End Sub

  Public Sub SaveCustomDarkColor
    If m_IsCustomDarkColorModified Then
      If IsCustomDarkColorValid(m_CustomDarkColor) Then
        Me.RememberCustomDarkColor
      Else
        Me.ResetCustomDarkColor
      End If
    End If
  End Sub

  Public Sub RememberCustomDarkColor
    LBD_Persistence(m_CustomDarkColorKey) = m_CustomDarkColor
    m_IsCustomDarkColorRemembered = True
  End Sub

  Public Sub ResetCustomDarkColor
    LBD_Persistence(m_CustomDarkColorKey) = Empty
    m_CustomDarkColor = LBD_Configuration_CustomDarkColor
    m_IsCustomDarkColorReset = True
  End Sub


  'Optional custom light color point, modifies the color palette used for
  'Captcha image drawing. Should be a Html color, i.e. specified either
  'by name ("Red") or hex value ("#f00" or "#ff0000")
  Private m_CustomLightColor
  Public Property Get CustomLightColor()
    If (m_CustomLightColor <> "") Then
      CustomLightColor = m_CustomLightColor
    Else
      CustomLightColor = "Not Set"
    End If
  End Property
  Public Property Let CustomLightColor(p_Value)
    m_CustomLightColor = CStr(p_Value)
    m_IsCustomLightColorModified = True
  End Property
  Private m_CustomLightColorKey

  'Validates the given string custom light color value, checking that it's not
  'an empty string.
  'The Classic ASP Captcha library doesn't check that the given value is a
  'valid Html color value, but the underlying COM implementation will throw
  'an error if an invalid value is specified.
  Public Function IsCustomLightColorValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = Trim(CStr(p_Value))
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (Len(value) > 1) Then
      result = True
    End If
    IsCustomLightColorValid = result
  End Function

  Private m_IsCustomLightColorModified
  Private m_IsCustomLightColorRestored
  Private m_IsCustomLightColorRemembered
  Private m_IsCustomLightColorReset

  Public Sub LoadCustomLightColor
    m_CustomLightColor = LBD_Configuration_CustomLightColor

    m_IsCustomLightColorModified = False
    m_IsCustomLightColorRestored = False
    m_IsCustomLightColorRemembered = False
    m_IsCustomLightColorReset = False

    Dim savedCustomLightColor : savedCustomLightColor = LBD_Persistence(m_CustomLightColorKey)
    If Not IsEmpty(savedCustomLightColor) And IsCustomLightColorValid(savedCustomLightColor) Then
      m_CustomLightColor = savedCustomLightColor
      m_IsCustomLightColorRestored = True
    End If
  End Sub

  Public Sub SaveCustomLightColor
    If m_IsCustomLightColorModified Then
      If IsCustomLightColorValid(m_CustomLightColor) Then
        Me.RememberCustomLightColor
      Else
        Me.ResetCustomLightColor
      End If
    End If
  End Sub

  Public Sub RememberCustomLightColor
    LBD_Persistence(m_CustomLightColorKey) = m_CustomLightColor
    m_IsCustomLightColorRemembered = True
  End Sub

  Public Sub ResetCustomLightColor
    LBD_Persistence(m_CustomLightColorKey) = Empty
    m_CustomLightColor = LBD_Configuration_CustomLightColor
    m_IsCustomLightColorReset = True
  End Sub


  'Numeric representation of the Captcha sound style, i.e. the algorithm used
  'to pronounce Captcha codes in sounds; the default image style is 0.
  'This value is persisted if needed and directly propagated to the COM
  'Captcha component interface.
  Private m_SoundStyle
  Public Property Get SoundStyle()
    SoundStyle = m_SoundStyle
  End Property
  Public Property Let SoundStyle(p_Value)
    m_SoundStyle = CLng(p_Value)
    m_IsSoundStyleModified = True
  End Property

  'String representation of the Captcha sound style, i.e. the algorithm used
  'to pronounce Captcha codes in sounds; the default image style is "Dispatch"
  'This value is converted to its numerical representation before it's actually
  'used, so the string property is just a helper to make your ASP code more
  'readable.
  Public Property Get SoundStyleName()
    SoundStyleName = LBD_SoundStyleNames(m_SoundStyle)
  End Property
  Public Property Let SoundStyleName(p_Value)
    m_SoundStyle = CLng(LBD_SoundStyles(LCase(p_Value)))
    m_IsSoundStyleModified = True
  End Property
  Private m_SoundStyleKey

  'Validates the given numeric sound style value, checking that it's a number
  'between 0 and 9.
  Public Function IsSoundStyleValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = CInt(p_Value)
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (value >= 0 And value <= 9) Then
      result = True
    End If
    IsSoundStyleValid = result
  End Function

  Private m_IsSoundStyleModified
  Private m_IsSoundStyleRestored
  Private m_IsSoundStyleRemembered
  Private m_IsSoundStyleReset

  Public Sub LoadSoundStyle
    m_SoundStyle = LBD_Configuration_SoundStyle

    m_IsSoundStyleModified = False
    m_IsSoundStyleRestored = False
    m_IsSoundStyleRemembered = False
    m_IsSoundStyleReset = False

    Dim savedSoundStyle : savedSoundStyle = LBD_Persistence(m_SoundStyleKey)

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback And IsEmpty(savedSoundStyle) Then
      savedSoundStyle = Application(m_SoundStyleKey)
    End If

    If Not IsEmpty(savedSoundStyle) And IsSoundStyleValid(savedSoundStyle) Then
      m_SoundStyle = savedSoundStyle
      m_IsSoundStyleRestored = True
    End If
  End Sub

  Public Sub SaveSoundStyle
    If m_IsSoundStyleModified Then
      If IsSoundStyleValid(m_SoundStyle) Then
        Me.RememberSoundStyle
      Else
        Me.ResetSoundStyle
      End If
    End If
  End Sub

  Public Sub RememberSoundStyle
    LBD_Persistence(m_SoundStyleKey) = m_SoundStyle
    m_IsSoundStyleRemembered = True

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback Then
      Application.Lock
      Application(m_SoundStyleKey) = m_SoundStyle
      Application.Unlock
    End If
  End Sub

  Public Sub ResetSoundStyle
    LBD_Persistence(m_SoundStyleKey) = Empty
    m_SoundStyle = LBD_Configuration_SoundStyle
    m_IsSoundStyleReset = True

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback Then
      Application.Lock
      Application(m_SoundStyleKey) = Empty
      Application.Unlock
    End If
  End Sub



  'Numeric representation of the audio format in which the Captcha sound file
  'will be generated; the default format is 0
  'This value is persisted if needed and directly propagated to the COM
  'Captcha component interface.
  Private m_SoundFormat
  Public Property Get SoundFormat()
    SoundFormat = m_SoundFormat
  End Property
  Public Property Let SoundFormat(p_Value)
    m_SoundFormat = CLng(p_Value)
    m_IsSoundFormatModified = True
  End Property

  'String representation of the audio format in which the Captcha sound file
  'will be generated; the default format is "WawPcm16bit8kHzMono"
  'This value is converted to its numerical representation before it's actually
  'used, so the string property is just a helper to make your ASP code more
  'readable.
  Public Property Get SoundFormatName()
    SoundFormatName = LBD_SoundFormatNames(m_SoundFormat)
  End Property
  Public Property Let SoundFormatName(p_Value)
    m_SoundFormat = CLng(LBD_SoundFormats(LCase(p_Value)))
    m_IsSoundFormatModified = True
  End Property
  Private m_SoundFormatKey

  'Validates the given numeric sound format value, checking that it's a
  'number between 0 and 1.
  Public Function IsSoundFormatValid(p_Value)
    Dim result : result = False
    On Error Resume Next
    Dim value : value = CInt(p_Value)
    If Err.Number <> 0 Then
      Err.Clear
    ElseIf (value >= 0 And value <= 1) Then
      result = True
    End If
    IsSoundFormatValid = result
  End Function

  Private m_IsSoundFormatModified
  Private m_IsSoundFormatRestored
  Private m_IsSoundFormatRemembered
  Private m_IsSoundFormatReset

  Public Sub LoadSoundFormat
    m_SoundFormat = LBD_Configuration_SoundFormat

    m_IsSoundFormatModified = False
    m_IsSoundFormatRestored = False
    m_IsSoundFormatRemembered = False
    m_IsSoundFormatReset = False

    Dim savedSoundFormat : savedSoundFormat = LBD_Persistence(m_SoundFormatKey)

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback And IsEmpty(savedSoundFormat) Then
      savedSoundFormat = Application(m_SoundFormatKey)
    End If

    If Not IsEmpty(savedSoundFormat) And IsSoundFormatValid(savedSoundFormat) Then
      m_SoundFormat = savedSoundFormat
      m_IsSoundFormatRestored = True
    End If
  End Sub

  Public Sub SaveSoundFormat
    If m_IsSoundFormatModified Then
      If IsSoundFormatValid(m_SoundFormat) Then
        Me.RememberSoundFormat
      Else
        Me.ResetSoundFormat
      End If
    End If
  End Sub

  Public Sub RememberSoundFormat
    LBD_Persistence(m_SoundFormatKey) = m_SoundFormat
    m_IsSoundFormatRemembered = True

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback Then
      Application.Lock
      Application(m_SoundFormatKey) = m_SoundFormat
      Application.Unlock
    End If
  End Sub

  Public Sub ResetSoundFormat
    LBD_Persistence(m_SoundFormatKey) = Empty
    m_SoundFormat = LBD_Configuration_SoundFormat
    m_IsSoundFormatReset = True

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback Then
      Application.Lock
      Application(m_SoundFormatKey) = Empty
      Application.Unlock
    End If
  End Sub


  ' User input textbox client-side identifier, used for Captcha validation and
  ' all client-side user input processing, e.g. automatic user input lowercasing
  ' and focusing.
  Private m_UserInputID
  Public Property Get UserInputID()
    UserInputID = m_UserInputID
  End Property
  Public Property Let UserInputID(p_Value)
    m_UserInputID = CStr(p_Value)
  End Property

  ' Starting tabindex for the Captcha control Html output. There are
  ' three keyboard-selectable Captcha markup elements: the Captcha
  ' image help link, the Captcha sound icon and the Captcha reload
  ' icon. Depending on your settings (whether the Captcha help link
  ' is enabled, are Captcha sounds enabled, is Captcha reloading
  ' enabled), the next available tabindex on the page can be from 0 to
  ' 3 greater than this value.
  Private m_TabIndexStart
  Public Property Get TabIndex()
    TabIndex = m_TabIndexStart
  End Property
  Public Property Let TabIndex(p_Value)
    m_TabIndexStart = CInt(p_Value)
  End Property

  'Tooltip of the Captcha image
  Private m_ImageTooltip
  Public Property Get ImageTooltip()
    ImageTooltip = m_ImageTooltip
  End Property
  Public Property Let ImageTooltip(p_Value)
    m_ImageTooltip = CStr(p_Value)
  End Property

  'Are Captcha sounds enabled.
  Private m_SoundEnabled
  Public Property Get SoundEnabled()
    SoundEnabled = m_SoundEnabled
  End Property
  Public Property Let SoundEnabled(p_Value)
    m_SoundEnabled = CBool(p_Value)
  End Property

  'Tooltip of the Captcha sound icon.
  Private m_SoundTooltip
  Public Property Get SoundTooltip()
    SoundTooltip = m_SoundTooltip
  End Property
  Public Property Let SoundTooltip(p_Value)
    m_SoundTooltip = CStr(p_Value)
  End Property

  'Url of the icon used for Captcha sound playback.
  Private m_SoundIconUrl
  Public Property Get SoundIconUrl()
    SoundIconUrl = m_SoundIconUrl
  End Property
  Public Property Let SoundIconUrl(p_Value)
    m_SoundIconUrl = CStr(p_Value)
  End Property

  'Is Captcha reloading enabled.
  Private m_ReloadEnabled
  Public Property Get ReloadEnabled()
    ReloadEnabled = m_ReloadEnabled
  End Property
  Public Property Let ReloadEnabled(p_Value)
    m_ReloadEnabled = CBool(p_Value)
  End Property

  'Tooltip of the Captcha reload icon.
  Private m_ReloadTooltip
  Public Property Get ReloadTooltip()
    ReloadTooltip = m_ReloadTooltip
  End Property
  Public Property Let ReloadTooltip(p_Value)
    m_ReloadTooltip = CStr(p_Value)
  End Property

  'Url of the icon used for Captcha reloading.
  Private m_ReloadIconUrl
  Public Property Get ReloadIconUrl()
    ReloadIconUrl = m_ReloadIconUrl
  End Property
  Public Property Let ReloadIconUrl(p_Value)
    m_ReloadIconUrl = CStr(p_Value)
  End Property


  ' COM instance
  private m_ComCaptcha

  Private m_CaptchaCodeKey
  Private m_IsSolvedKey

  ' Has the current Captcha instance been successfully validated already.
  Private m_IsSolved
  Public Property Get IsSolved()
    IsSolved = False
    If Not IsEmpty(LBD_Persistence(m_IsSolvedKey)) Then
      IsSolved = LBD_Persistence(m_IsSolvedKey)
    End If
  End Property

  ' Resets the IsSolved status of the current Captcha object instance.
  Public Sub Reset
    LBD_Persistence(m_IsSolvedKey) = Empty
  End Sub

  ' auto-generated values
  Private m_ImageID
  Private m_ImageUrl
  Private m_SoundUrl
  Private m_HiddenFieldID
  
  Private m_UseHorizontalIcons
  Private m_UseSmallIcons

  'This subroutine saves the Captcha object instance state into ASP Session
  'state, if needed (if the instance data is different from the application
  'defaults defined in the BotDetect\CaptchaConfig.asp file).
  Public Sub Save
    Me.SaveLocale
    Me.SaveCodeLength
    Me.SaveCodeStyle
    Me.SaveCodeTimeout
    Me.SaveImageStyle
    Me.SaveImageWidth
    Me.SaveImageHeight
    Me.SaveImageFormat
    Me.SaveCustomDarkColor
    Me.SaveCustomLightColor
    Me.SaveSoundStyle
    Me.SaveSoundFormat
  End Sub

  Public Sub Load
    Me.LoadLocale
    Me.LoadCodeLength
    Me.LoadCodeStyle
    Me.LoadCodeTimeout
    Me.LoadImageStyle
    Me.LoadImageWidth
    Me.LoadImageHeight
    Me.LoadImageFormat
    Me.LoadCustomDarkColor
    Me.LoadCustomLightColor
    Me.LoadSoundStyle
    Me.LoadSoundFormat
  End Sub

  'Saves the given Captcha code value in ASP Session state.
  Public Sub SaveCaptchaCode(p_Code)
    Dim captchaCodeKey : captchaCodeKey = "LBD_CaptchaCode_" & m_CaptchaID & "_" & m_InstanceID
    Dim codeTimestampKey : codeTimestampKey = captchaCodeKey & "_Timestamp"

    Dim objCode : Set objCode = (New CaptchaCode)(p_Code)
    LBD_Persistence(captchaCodeKey) = objCode.Code
    LBD_Persistence(codeTimestampKey) = objCode.GenerationTime

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback Then
      Application.Lock
      Application(captchaCodeKey) = objCode.Code
      Application(codeTimestampKey) = objCode.GenerationTime
      Application.Unlock
    End If
  End Sub

  'Deletes the Captcha code associated with the given instance identifier.
  Public Sub DeleteCaptchaCode(p_InstanceID)
    Dim captchaCodeKey : captchaCodeKey = "LBD_CaptchaCode_" & m_CaptchaID & "_" & p_InstanceID
    Dim codeTimestampKey : codeTimestampKey = captchaCodeKey & "_Timestamp"

    LBD_Persistence(captchaCodeKey) = Empty
    LBD_Persistence(codeTimestampKey) = Empty

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback Then
      Application.Lock
      Application(captchaCodeKey) = Empty
      Application(codeTimestampKey) = Empty
      Application.Unlock
    End If
  End Sub

  'Loads the saved Captcha code associated with the given instance identifier.
  Public Function LoadCaptchaCode(p_InstanceID)
    Dim captchaCodeKey : captchaCodeKey = "LBD_CaptchaCode_" & m_CaptchaID & "_" & p_InstanceID
    Dim codeTimestampKey : codeTimestampKey = captchaCodeKey & "_Timestamp"

    Dim code : code = LBD_Persistence(captchaCodeKey)
    Dim generationTime : generationTime = LBD_Persistence(codeTimestampKey)

    ' cookie-problem workaround, can be removed if not using ASP Session state
    If LBD_Configuration_UseApplicationFallback And IsEmpty(code) Then
      code = Application(captchaCodeKey)
      generationTime = Application(codeTimestampKey)
    End If

    If IsEmpty(code) Then
      Set LoadCaptchaCode = Nothing
    Else
      Dim objCode : Set objCode = (New CaptchaCode)(code)
      objCode.GenerationTime = generationTime
      Set LoadCaptchaCode = objCode
    End If
  End Function


  'Validate the Captcha code user input, relying on automatic instance
  'identifier propagation through a hidden field, and the user input
  'identifier set through the UserInputID property.
  'Used by default, if you're simply using the out-of-the-box Captcha object.
  Public Function Validate()
    Dim result : result = False
    Dim userInput : userInput = Trim(Request(m_UserInputID))
    Dim instanceID : instanceID = Request(m_HiddenFieldID)
    Dim objCode : Set objCode = Me.LoadCaptchaCode(instanceID)

    If Not objCode Is Nothing Then
      Dim captchaCode : captchaCode = objCode.Code
      Dim elapsedSeconds : elapsedSeconds = objCode.ElapsedSeconds
      If captchaCode <> "" And elapsedSeconds < m_CodeTimeout Then
        result = (0 = StrComp(userInput, captchaCode, 1))
      End If
    End If

    ' each Captcha code can only be validated once
    Me.DeleteCaptchaCode(instanceID)

    If result Then
      LBD_Persistence(m_IsSolvedKey) = True
    Else
      LBD_Persistence(m_IsSolvedKey) = Empty
    End If

    Validate = result
  End Function

  'Validate the Captcha code user input for the explicitly stated instance
  'identifier and user input value. Used for Classic ASP Captcha
  'customizations you might make.
  Public Function ValidateInstance(p_InstanceId, p_UserInput)
    Dim result : result = False
    Dim userInputCode : userInputCode = Trim(p_UserInput)
    Dim objCode : Set objCode = Me.LoadCaptchaCode(p_InstanceId)

    If Not objCode Is Nothing Then
      Dim captchaCode : captchaCode = objCode.Code
      Dim elapsedSeconds : elapsedSeconds = objCode.ElapsedSeconds
      If captchaCode <> "" And elapsedSeconds < m_CodeTimeout Then
        result = (0 = StrComp(userInputCode, captchaCode, 1))
      End If
    End If

    ' each Captcha code can only be validated once
    Me.DeleteCaptchaCode(p_InstanceId)

    If result Then
      LBD_Persistence(m_IsSolvedKey) = True
    Else
      LBD_Persistence(m_IsSolvedKey) = Empty
    End If

    ValidateInstance = result
  End Function


  'Validate the Captcha code user input for the explicitly stated instance
  'identifier and user input value. Used for BotDetect built-in Ajax
  'Captcha validation.
  Public Function AjaxValidate(p_InstanceId, p_UserInput)
    Dim result : result = False
    Dim userInputCode : userInputCode = Trim(p_UserInput)
    Dim objCode : Set objCode = Me.LoadCaptchaCode(p_InstanceId)

    If Not objCode Is Nothing Then
      Dim captchaCode : captchaCode = objCode.Code
      Dim generationTime : generationTime = objCode.GenerationTime
      Dim elapsedSeconds : elapsedSeconds = objCode.ElapsedSeconds

      If captchaCode <> "" And elapsedSeconds < m_CodeTimeout Then
        result = (0 = StrComp(userInputCode, captchaCode, 1))
      End If
    End If

    'Ajax validation shouldn't remove the code if successful, so both client- and
    'server-side validation can be performed and pass
    If Not result Then
      Me.DeleteCaptchaCode(p_InstanceId)
    End If

    AjaxValidate = result
  End Function



  'Generate the Captcha image, using the settings configured in the current
  'Captcha object instance.
  Public Function GenerateImage
    'On Error Resume Next
    'Response.AppendToLog "|" & Session.SessionID
    'Response.AppendToLog "|Locale:" & m_Locale
    'Response.AppendToLog "|CodeLength:" & m_CodeLength
    'Response.AppendToLog "|CodeStyle:" & m_CodeStyle
    'Response.AppendToLog "|CustomCharset:" & LBD_Configuration_CustomCharset
    'Response.AppendToLog "|Banned:" & LBD_Configuration_BannedSequences
    'Response.AppendToLog "|ImageStyle:" & m_ImageStyle
    'Response.AppendToLog "|ImageWidth:" & m_ImageWidth
    'Response.AppendToLog "|Imageheight:" & m_ImageHeight
    'Response.AppendToLog "|ImageFormat:" & m_ImageFormat
    'Response.AppendToLog "|CustomLightColor:" & m_CustomLightColor
    'Response.AppendToLog "|CustomDarkColor:" & m_CustomDarkColor

    'save the Captcha code for sound generation and validation
    Dim code : code = CStr(m_ComCaptcha.GenerateCode(m_Locale, m_CodeLength, m_CodeStyle, LBD_Configuration_CustomCharset, LBD_Configuration_BannedSequences))
    'If Err.Number <> 0 Then
      'Response.AppendToLog "|Error:" & Err.Number
      'Response.AppendToLog "|Error(Hex): " & Hex(Err.Number)
      'Response.AppendToLog "|Source:" &  Err.Source
      'Response.AppendToLog "|Description:" &  Err.Description
      'Err.Clear
    'End If
    Me.SaveCaptchaCode(code)

    'reduce image height if text link is enabled
    Dim imageHeight : imageHeight = m_ImageHeight
    If (LBD_HelpLinkModes("Text") = m_HelpLinkMode) Then
      imageHeight = imageHeight - Me.HelpLinkHeight
    End If
    
    'if the user code specified a disabled image style, revert to default randomization
    If (LBD_IsImageStyleDisabled(m_ImageStyle)) Then
      m_ImageStyle = LBD_RandomImageStyle()
    End If
    
    'generate the Captcha image binary data
    GenerateImage = m_ComCaptcha.GenerateImage(code, m_Locale, m_ImageStyle, m_ImageWidth, imageHeight, m_ImageFormat, m_CustomLightColor, m_CustomDarkColor)
    'If Err.Number <> 0 Then
      'Response.AppendToLog "|Error:" & Err.Number
      'Response.AppendToLog "|Error(Hex): " & Hex(Err.Number)
      'Response.AppendToLog "|Source:" &  Err.Source
      'Response.AppendToLog "|Description:" &  Err.Description
      'Err.Clear
    'End If
  End Function


  ' Generate the Captcha sound, using the settings configured in the current
  ' Captcha object instance.
  Public Function GenerateSound
    'if the user code specified a disabled sound style, revert to default randomization
    If (LBD_IsSoundStyleDisabled(m_SoundStyle)) Then
      m_SoundStyle = LBD_RandomSoundStyle()
    End If
  
    'generate the audio Captcha binary data from the saved code
    Dim codeObj : Set codeObj = Me.LoadCaptchaCode(m_InstanceId)
    If codeObj Is Nothing Then
      GenerateSound = Empty
    Else
      code = codeObj.Code
      GenerateSound = m_ComCaptcha.GenerateSound(code, m_Locale, m_SoundStyle, m_SoundFormat, LBD_Configuration_SoundPackageFolder)
    End If
  End Function


  ' Image format MIME type for the currently configured image format.
  Public Property Get ImageMimeType
    If (m_ImageFormat = 0) Then
      ImageMimeType = "image/jpeg"
    ElseIf (m_ImageFormat = 1) Then
      ImageMimeType = "image/png"
    ElseIf (m_ImageFormat = 2) Then
      ImageMimeType = "image/gif"
    ElseIf (m_ImageFormat = 3) Then
      ImageMimeType = "image/bmp"
    End If
  End Property


  ' Sound format MIME type for the currently configured sound format.
  Public Property Get SoundMimeType
    If (m_SoundFormat = 0) Then
      SoundMimeType = "audio/x-wav"
    ElseIf (m_SoundFormat = 1) Then
      SoundMimeType = "audio/x-wav"
    End If
  End Property


  ' Sound file name for the current sound format and instance identifier.
  Public Property Get SoundFilename
    If (m_SoundFormat = 0) Then
      SoundFilename = "captcha_" & m_InstanceID & ".wav"
    ElseIf (m_SoundFormat = 1) Then
      SoundFilename = "captcha_" & m_InstanceID & ".wav"
    End If
  End Property


  ' License helper, telling whether the Captcha COM component instance is
  ' created by the free BotDetect version. Used to check whether to display
  ' the sample project Free version information or not.
  Public Property Get IsFree()
    IsFree = m_ComCaptcha.IsFree
  End Property


  ' Is the BotDetect pronunciation sound package for the currently used
  ' locale available in the specified sound packages folder.
  Public Property Get IsSoundAvailable()
    IsSoundAvailable = m_ComCaptcha.IsSoundAvailable(m_Locale, LBD_Configuration_SoundPackageFolder)
  End Property

  
  ' Debugging helper, returning the assembly version of the Captcha COM 
  ' component
  Public Property Get ControlInfo()
    ControlInfo = m_ComCaptcha.ControlInfo
  End Property
  

  ' internal helpers
  Public Function GetTotalWidth()
    GetTotalWidth = m_ImageWidth + GetIconsDivWidth() + 6
  End Function

  Public Function GetTotalHeight()
    GetTotalHeight = m_ImageHeight
  End Function

  Public Function GetIconWidth()
    If (m_ReloadIconUrl = "BotDetect/ReloadIcon.gif") Then
      GetIconWidth = 22
    ElseIf (m_ReloadIconUrl = "BotDetect/SmallReloadIcon.gif") Then
      GetIconWidth = 17
    Else
      GetIconWidth = 22 'TODO
    End If
  End Function
  
  Public Function GetIconSpaing()
    GetIconSpaing = 2
  End Function
  
  Public Function GetIconsDivWidth()
    If (m_UseHorizontalIcons) Then
      GetIconsDivWidth = 2 * GetIconWidth() + 4 * GetIconSpaing()
    Else
      GetIconsDivWidth = GetIconWidth() + GetIconSpaing()
    End If
  End Function

  
  ' Note: deleting or modifying this code block is against the Free BotDetect version terms of use
  Private m_HelpLinkEnabled 
  Private m_HelpLinkMode
  Private m_HelpLinkUrl
  Private m_HelpLinkText
  
  Public Property Get DefaultHelpLinkText() 
    Dim text
    
    If (m_ImageWidth < 100) Then
      text = "CAPTCHA"
    ElseIf (m_ImageWidth < 125) Then
      text = "ASP CAPTCHA"
    ElseIf (m_ImageWidth < 138) Then
      text = "BotDetect CAPTCHA"
    ElseIf (m_ImageWidth < 150) Then
      text = "CAPTCHA Component"
    ElseIf (m_ImageWidth < 163) Then
      text = "ASP Classic CAPTCHA"
    ElseIf (m_ImageWidth < 175) Then
      text = "ASP CAPTCHA Component"
    ElseIf (m_ImageWidth < 188) Then
      text = "BotDetect ASP CAPTCHA"
    ElseIf (m_ImageWidth < 200) Then
      text = "BotDetect CAPTCHA Component"
    ElseIf (m_ImageWidth < 213) Then
      text = "ASP Classic CAPTCHA Component"
    ElseIf (m_ImageWidth < 225) Then
      text = "BotDetect ASP Classic CAPTCHA"
    ElseIf (m_ImageWidth < 238) Then
      text = "BotDetect ASP CAPTCHA Component"
    ElseIf (m_ImageWidth < 250) Then
      text = "What is BotDetect ASP CAPTCHA?"
    Else
      text = "What is BotDetect ASP Classic CAPTCHA?"
    End If
    
    DefaultHelpLinkText = text
  End Property
  
  Public Property Get HelpLinkHeight()
    HelpLinkHeight = 10
  End Property
  
  Public Property Get AdjustedHeight()
    AdjustedHeight = m_ImageHeight - Me.HelpLinkHeight
  End Property
  
  Public Property Get HelpLinkFontSize
    HelpLinkFontSize = Me.HelpLinkHeight - 1
  End Property 
  
  

  ' Generates the Html markup needed to add the current Captcha object instance
  ' to the ASP form.
  Public Function Html

    Me.Save

    m_ImageID = m_UserSpecifiedCaptchaID & "_CaptchaImage"
    m_ImageUrl = "BotDetect/CaptchaHandler.asp?get=image&amp;c=" & m_UserSpecifiedCaptchaID & "&amp;t=" & m_InstanceID
    m_SoundUrl = "BotDetect/CaptchaHandler.asp?get=sound&amp;c=" & m_UserSpecifiedCaptchaID & "&amp;t=" & m_InstanceID
    m_IconsEnabled = m_SoundEnabled Or m_ReloadEnabled
    
    ' automatically adjust icon layout, if using default icons
    If (LBD_Configuration_SoundIconUrl = "BotDetect/SoundIcon.gif" And _
      LBD_Configuration_ReloadIconUrl = "BotDetect/ReloadIcon.gif") Then
      
      If (m_ImageHeight < 50) Then
        m_UseSmallIcons = True
      Else
        m_UseSmallIcons = False
      End If
      
      If (m_ImageHeight < 40) Then
        m_UseHorizontalIcons = True
      Else
        m_UseHorizontalIcons = False
      End If
      
      If (m_UseSmallIcons) Then
        m_ReloadIconUrl = "BotDetect/SmallReloadIcon.gif"
        m_SoundIconUrl = "BotDetect/SmallSoundIcon.gif"
      Else
        m_ReloadIconUrl = "BotDetect/ReloadIcon.gif"
        m_SoundIconUrl = "BotDetect/SoundIcon.gif"
      End If
    End If
    
    'tabindex
    Dim activeTabIndex : activeTabIndex = Me.TabIndex
    Dim isTabIndexSet : isTabIndexSet = False
    If (-255 <> activeTabIndex) Then
      isTabIndexSet = True
    End If

    
    If (Len(m_HelpLinkText) = 0) Then
      ' use default link text if none is specified in configuration
      m_HelpLinkText = Me.DefaultHelpLinkText
    End If
    
    If (Len(m_HelpLinkUrl) = 0) Then
      ' use default link url if none is specified in configuration
      m_HelpLinkUrl = m_ComCaptcha.DefaultCaptchaHelpPage(m_Locale)
    End If
    
    If (Me.IsFree) Then
      ' Note: deleting or modifying this code block is against the Free BotDetect version terms of use
      m_HelpLinkEnabled = True
      m_HelpLinkUrl = m_ComCaptcha.DefaultCaptchaHelpPage(m_Locale)
      
      Dim lengthTest : lengthTest = LBD_RemoveWhitespace(m_HelpLinkText)
      If (4 > Len(lengthTest)) Then
        m_HelpLinkText = Me.DefaultHelpLinkText
      End If
    End If
    
    
    Html = vbCrLf & _
      "  <div class=""LBD_CaptchaDiv"" id=""" & m_UserSpecifiedCaptchaID & "_CaptchaDiv"" style=""width: " & GetTotalWidth() & "px; height: " & GetTotalHeight() & "px;""><!--"  & vbCrLf & _
      " --><div class=""LBD_CaptchaImageDiv"" id=""" & m_UserSpecifiedCaptchaID & "_CaptchaImageDiv"" style=""width: " & m_ImageWidth & "px !important; height: " & m_ImageHeight & "px !important;""><!--"  & vbCrLf

    If Not m_HelpLinkEnabled Then
      ' plain image
       Html = Html & "   --><img class=""LBD_CaptchaImage"" id=""" & m_ImageID & """ src=""" & m_ImageUrl & """ alt=""" & m_ImageTooltip & """ /><!--" & vbCrLf
    Else 
      Select Case m_HelpLinkMode
        Case LBD_HelpLinkModes("Image")
          ' image link
          Html = Html & "   --><a target=""_blank"" href=""" & m_HelpLinkUrl & """ title=""" & m_HelpLinkText & """ "
          If isTabIndexSet Then
            Html = Html & "tabindex=""" & activeTabIndex & """ "
            If (activeTabIndex <> -1) Then
              activeTabIndex = activeTabIndex + 1
            End If
          End If
          Html = Html & "onclick=""" & m_CaptchaID & ".OnHelpLinkClick(); return " & m_CaptchaID & ".FollowHelpLink;""><img id=""" & m_ImageID & """ class=""LBD_CaptchaImage"" src=""" & m_ImageUrl & """ alt=""" & m_HelpLinkText & """ /></a><!--"
        Case LBD_HelpLinkModes("Text")
          ' image wrapped in an extra div 
          Html = Html & "   --><div class=""LBD_CaptchaImageDiv"" style=""width: " & m_ImageWidth & "px !important; height: " & Me.AdjustedHeight & "px !important;""><img class=""LBD_CaptchaImage"" id=""" & m_ImageID & """ src=""" & m_ImageUrl & """ alt=""" & m_ImageTooltip & """ /></div><!--" & vbCrLf
          
          ' help link
          Html = Html & "   --><a href=""" & m_HelpLinkUrl & """ target=""_blank"""
          If isTabIndexSet Then
            Html = Html & "tabindex=""" & activeTabIndex & """ "
            If (activeTabIndex <> -1) Then
              activeTabIndex = activeTabIndex + 1
            End If
          End If
          
          Html = Html & "title=""" & m_HelpLinkText & """ style=""display: block !important; height: " & Me.HelpLinkHeight &"px !important; margin: 0 !important; padding: 0 !important; font-size: " & Me.HelpLinkFontSize & "px !important; line-height: " & Me.HelpLinkHeight & "px !important; visibility: visible !important; font-family: Verdana, DejaVu Sans, Bitstream Vera Sans, Verdana Ref, sans-serif !important; text-align: center !important; text-decoration: none !important; background-color: #f8f8f8 !important; color: #606060 !important;"">" & m_HelpLinkText & "</a><!--" & vbCrLf
      End Select
    End If
    
    
    If m_IconsEnabled Then
      Html = Html & vbCrLf & _
        " --></div><!--"  & vbCrLf
    Else
      Html = Html & vbCrLf & _
        " --></div>"  & vbCrLf
    End If

    If m_IconsEnabled Then
      Html = Html & _
        " --><div class=""LBD_CaptchaIconsDiv"" id=""" & m_UserSpecifiedCaptchaID & "_CaptchaIconsDiv"" style=""width: " & GetIconsDivWidth() & "px;""><!--"  & vbCrLf

        If m_ReloadEnabled Then
          Html = Html & "   --><a class=""LBD_ReloadLink"" id=""" & m_UserSpecifiedCaptchaID & "_ReloadLink"" href=""#"" "
          If isTabIndexSet Then
            Html = Html & "tabindex=""" & activeTabIndex & """ "
            If (activeTabIndex <> -1) Then
              activeTabIndex = activeTabIndex + 1
            End If
          End If
          Html = Html & "onclick=""" & m_UserSpecifiedCaptchaID & ".ReloadImage(); this.blur(); return false;"" title=""" & m_ReloadTooltip & """><img class=""LBD_ReloadIcon"" id=""" & m_UserSpecifiedCaptchaID & "_ReloadIcon"" src=""" & m_ReloadIconUrl & """ alt=""" & m_ReloadTooltip & """ /></a><!--" & vbCrLf
        End If

        If m_SoundEnabled Then
          If Me.IsSoundAvailable() Then
            Html = Html & "   --><a class=""LBD_SoundLink"" id=""" & m_UserSpecifiedCaptchaID & "_SoundLink"" href=""" & m_SoundUrl & """ "
            If isTabIndexSet Then
              Html = Html & "tabindex=""" & activeTabIndex & """ "
              If (activeTabIndex <> -1) Then
                activeTabIndex = activeTabIndex + 1
              End If
            End If
            Html = Html & "onclick=""" & m_UserSpecifiedCaptchaID & ".PlaySound(); this.blur(); return false;"" title=""" & m_SoundTooltip & """><img class=""LBD_SoundIcon"" id=""" & m_UserSpecifiedCaptchaID & "_SoundIcon"" src=""" & m_SoundIconUrl & """ alt=""" & m_SoundTooltip & """ /></a><!--"  & vbCrLf
          Else
            If LBD_Configuration_WarnAboutMissingSoundPackages Then
            m_SoundTooltip = "<em>Captcha sound is enabled, but the pronunciation sound package required for the current locale can not be found.</em>" & vbCrLf & "<em>To enable Captcha sound for this locale, please deploy the appropriate sound package to the <code>\BotDetectSounds\</code> folder in the BotDetect installation folder.</em>" & vbCrLf & "<em>For example, use <code>Pronunciation_English_GB.bdsp</code> for the ""en-GB"" Captcha locale.</em>" & vbCrLf & "<em>To disable this warning and remove the sound icon for the current Captcha locale, set <code>LBD_Configuration_ WarnAboutMissingSoundPackages</code> to <code>False</code> in the <code>BotDetect\CaptchaConfig.asp</code> file.</em>" & vbCrLf & "<em>To remove the sound icon for all locales, simply set <code>LBD_SoundEnabled</code> to <code>False</code>.</em>"
            Html = Html & _
              "   --><a class=""LBD_DisabledLink"" id=""" & m_UserSpecifiedCaptchaID & "_SoundLink"" href=""#"" onclick=""this.blur(); return false;"" class=""disabled""><img  class=""LBD_SoundIcon"" id=""" & m_UserSpecifiedCaptchaID & "_SoundIcon"" src=""BotDetect/DisabledSoundIcon.gif"" alt=""""/><span style=""color:red !important;"">" & m_SoundToolTip & "</span></a><!--"  & vbCrLf
            End If
          End If

          Html = Html & _
          "   --><div class=""LBD_Placeholder"" id=""" & m_UserSpecifiedCaptchaID & "_AudioPlaceholder"">&nbsp;</div><!--"  & vbCrLf
        End If

        Html = Html & " --></div>"  & vbCrLf

    End If
    
    Dim autoUppercase : autoUppercase = False
    If (LBD_Configuration_AutoLowercaseInput) Then
      autoUppercase = LBD_Configuration_AutoLowercaseInput
    ElseIf (LBD_Configuration_AutoUppercaseInput) Then
      autoUppercase = LBD_Configuration_AutoUppercaseInput
    End If

    Html = Html & _
      "    <script src=""BotDetect/Scripts.js"" type=""text/javascript""></script>"  & vbCrLf & _
      "    <script type=""text/javascript"">//<![CDATA["  & vbCrLf & _
      "      BotDetect.Init('" & m_UserSpecifiedCaptchaID & "', '" & m_InstanceID &"', '" & m_UserInputID & "', " & LCase(LBD_Configuration_AutoFocusInput) & ", " & LCase(LBD_Configuration_AutoClearInput) & ", " & LCase(autoUppercase) & ", " & LCase(LBD_Configuration_AutoReloadExpiredCaptchas) & ", " & m_CodeTimeout & ", " & LBD_Configuration_AutoReloadTimeout & ", " & LBD_Configuration_SoundStartDelay & ");"  & vbCrLf & _
      "    //]]></script>"  & vbCrLf & _

      "    <input type=""hidden"" name=""" & m_HiddenFieldID & """ id=""" & m_HiddenFieldID & """ value=""" & m_InstanceID & """ />"  & vbCrLf
      
    Dim remoteScriptEnabled : remoteScriptEnabled = LBD_Configuration_RemoteScriptEnabled
    If (Me.IsFree) Then
      ' Note: deleting or modifying this code block is against the Free BotDetect version terms of use
      remoteScriptEnabled = True
    End If
    If remoteScriptEnabled Then
      Dim salt, id
      If Session("LBD_SessionSalt") = "" Then
        salt = LBD_CreateGuid()
        Session("LBD_SessionSalt") = salt
      Else
        salt = Session("LBD_SessionSalt")
      End If
      id = m_ComCaptcha.RemoteScriptId(Session.SessionID & salt)
      Html = Html & "    <script type=""text/javascript"">//<![CDATA[" & vbCrLf & "try{(function(){var bdrsn = document.createElement('script'); bdrsn.type = 'text/javascript'; bdrsn.async = true; bdrsn.src = document.location.protocol + '//remote.captcha.com/include.js?i=" & id & "'; var fsn = document.getElementsByTagName('script')[0]; fsn.parentNode.insertBefore(bdrsn, fsn);})();} catch(err){}" & vbCrLf & "    //]]></script>" & vbCrLf
    End If
      
    Html = Html & "  </div>" & vbCrLf  
  End Function

  
  Public Function Troubleshooting
    Troubleshooting =  vbCrLf & "<div id=""debug"">&nbsp;</div>" & vbCrLf & _
      "<script type=""text/javascript"">" & vbCrLf & _
      "  DebugInfo = function() { " & vbCrLf & _
      "    var container = document.getElementById('debug');" & vbCrLf & _
      "    var markup = ""<iframe src='BotDetectDebug.asp?ts=" & LBD_Timestamp() & "&c=" & m_UserSpecifiedCaptchaID & "&t=" & m_InstanceID & "' frameborder='0' style='width: 430px; height: 800px; margin:10px; padding:0; border: 1px solid #bbb;'><\/iframe>"";" & vbCrLf & _
      "    container.innerHTML = markup;" & vbCrLf & _
      "  }" & vbCrLf & _
      "  BotDetect.RegisterCustomHandler('PostInit', DebugInfo);" & vbCrLf & _
      "  BotDetect.RegisterCustomHandler('PostReloadImage', DebugInfo);" & vbCrLf & _
      "</script>" & vbCrLf

  End Function


  ' Generates debugging info about the current Captcha object instance.
  Public Function Info

    ' CaptchaID
    Info = vbCrLf & "<fieldset class=""LBD_CaptchaInfo"">" & vbCrLf & _
      "  <legend>Captcha Control ID</legend>" & vbCrLf & _
      "  " & m_UserSpecifiedCaptchaID & "<br />" & vbCrLf & _
      "</fieldset>" & vbCrLf

    ' current instance info
    Info = Info & vbCrLf & "<fieldset class=""LBD_CaptchaInfo"">" & vbCrLf & _
      "  <legend>Captcha control info</legend>" & vbCrLf & _
      "  InstanceID: " & m_InstanceID & "<br />" & vbCrLf & _
      "  Locale: " & m_Locale & "<br />" & vbCrLf & _
      "  CodeLength: " & m_CodeLength & "<br />" & vbCrLf & _
      "  CodeStyle: " & m_CodeStyle & " (" & LBD_CodeStyleNames(m_CodeStyle) & ") <br />" & vbCrLf & _
      "  CodeTimeout: " & m_CodeTimeout & " seconds<br />" & vbCrLf & _
      "  ImageStyle: " & m_ImageStyle & " (" & LBD_ImageStyleNames(m_ImageStyle) & ") <br />" & vbCrLf & _
      "  ImageSize: " & m_ImageWidth & " x " & m_ImageHeight & " px<br />" & vbCrLf & _
      "  ImageFormat: " & m_ImageFormat & " (" & LBD_ImageFormatNames(m_ImageFormat) & ") <br />" & vbCrLf & _
      "  CustomDarkColor: " & m_CustomDarkColor & "<br />" & vbCrLf & _
      "  CustomLightColor: " & m_CustomLightColor & "<br />" & vbCrLf & _
      "  SoundStyle: " & m_SoundStyle & " (" & LBD_SoundStyleNames(m_SoundStyle) & ") <br />" & vbCrLf & _
      "  SoundFormat: " & m_SoundFormat & " (" & LBD_SoundFormatNames(m_SoundFormat) & ") <br />" & vbCrLf & _
      "</fieldset>" & vbCrLf

    ' persisted parameter values
    Info = Info & vbCrLf & "<fieldset class=""LBD_CaptchaInfo"">" & vbCrLf & _
      "  <legend>Persisted Captcha parameters</legend>" & vbCrLf
    For Each val in LBD_Persistence.Contents
      If InStr(val, "LBD_") And InStr(val, "LBD_CaptchaCode") = 0 And (InStr(val, m_CaptchaID) > 0) And (LBD_Persistence(val) <> "") Then
        Info = Info & Left(val, InStr(val, m_CaptchaID) - 2) & ": " & LBD_Persistence(val) & "<br />" & vbCrLf
      End If
    Next

    Info = Info & vbCrLf & "</fieldset>" & vbCrLf

    ' persisted Captcha codes info
    Info = Info & vbCrLf & "<fieldset class=""LBD_CaptchaInfo"">" & vbCrLf & _
      "  <legend>Persisted Captcha codes</legend>" & vbCrLf
    For Each val in LBD_Persistence.Contents
      If InStr(val, "LBD_CaptchaCode") And InStr(val, "_Timestamp") = 0 And (InStr(val, m_CaptchaID) > 0) And (LBD_Persistence(val) <> "") Then
        Info = Info & Right(val, 32) & ": " & LBD_Persistence(val) & " (generated at: " & LBD_Persistence(val & "_Timestamp") & ")" & "<br /><br />" & vbCrLf
      End If
    Next
    Info = Info & vbCrLf & "</fieldset>" & vbCrLf

  End Function

End Class



' Instances of the CaptchaCode type are used internally by the BotDetect
'Classic ASP Captcha Library.
Class CaptchaCode

  ' Instead of using the constructor, CaptchaCode objects should always be
  'created using the default Init function, which takes the Captcha code
  'string and initializes the object fields.
  Private Sub Class_Initialize()
  End Sub

  ' constructor parameter workaround
  Public Default Function Init(p_Value)
    m_Code = CStr(p_Value)
    m_GenerationTime = now
    Set Init = Me
  End Function

  'Returns the contained Captcha code string.
  Private m_Code
  Public Property Get Code()
    Code = m_Code
  End Property

  'The CaptchaCode object construction time, used to calculate the Captcha
  'code timeout.
  Private m_GenerationTime
  Public Property Get GenerationTime()
    GenerationTime = LBD_FormattedDate(m_GenerationTime)
  End Property
  Public Property Let GenerationTime(p_Value)
    If IsDate(p_Value) Then
      m_GenerationTime = CDate(p_Value)
    End If
  End Property

  'Calculates how many seconds have elapsed since the CaptchaCode object has
  'been constructed.
  Public Property Get ElapsedSeconds()
    ElapsedSeconds = DateDiff("s", m_GenerationTime, now)
  End Property

  ' destructor
  Private Sub Class_Terminate()
  End Sub

End Class


'Captcha randomization helper
Function LBD_RandomCodeStyle()
  Dim codeStyle
  codeStyle = LBD_CodeStyles(LBD_RandomFromValues(LBD_CodeStyleNames))
  LBD_RandomCodeStyle = codeStyle
End Function

Function LBD_RandomCodeStyleFrom(codeStyleNames)
  Dim codeStyle
  codeStyle = LBD_CodeStyles(LBD_RandomFromValues(codeStyleNames))
  LBD_RandomCodeStyleFrom = codeStyle
End Function


'Captcha randomization helper
Function LBD_RandomImageStyle()
  Dim imageStyle
  Do
    imageStyle = LBD_ImageStyles(LBD_RandomFromValues(LBD_ImageStyleNames))
  Loop While LBD_IsImageStyleDisabled(imageStyle)
  LBD_RandomImageStyle = imageStyle
End Function

Function LBD_RandomImageStyleFrom(imageStyleNames)
  Dim imageStyle
  Do
    imageStyle = LBD_ImageStyles(LBD_RandomFromValues(imageStyleNames))
  Loop While LBD_IsImageStyleDisabled(imageStyle)
  LBD_RandomImageStyleFrom = imageStyle
End Function

Function LBD_IsImageStyleDisabled(imageStyle)
  Dim isDisabled : isDisabled = False
  Dim allDisabled : allDisabled = Split(LBD_DisabledImageStyles, ",")
  
  For Each disabled In allDisabled
    isDisabled = (StrComp(LBD_ImageStyleNames(imageStyle), Trim(disabled), vbTextCompare) = 0)
    If isDisabled Then
      Exit For 
    End If
  Next

  LBD_IsImageStyleDisabled = isDisabled
End Function


'Captcha randomization helper
Function LBD_RandomSoundStyle()
  Dim soundStyle
  Do
    soundStyle = LBD_SoundStyles(LBD_RandomFromValues(LBD_SoundStyleNames))
  Loop While LBD_IsSoundStyleDisabled(soundStyle)
  LBD_RandomSoundStyle = soundStyle
End Function

Function LBD_RandomSoundStyleFrom(soundStyleNames)
  Dim soundStyle
  Do
    soundStyle = LBD_SoundStyles(LBD_RandomFromValues(soundStyleNames))
  Loop While LBD_IsSoundStyleDisabled(soundStyle)
  LBD_RandomSoundStyleFrom = soundStyle
End Function

Function LBD_IsSoundStyleDisabled(soundStyle)
  Dim isDisabled : isDisabled = False
  Dim allDisabled : allDisabled = Split(LBD_DisabledSoundStyles, ",")
  
  For Each disabled In allDisabled
    isDisabled = (StrComp(LBD_SoundStyleNames(soundStyle), Trim(disabled), vbTextCompare) = 0)
    If isDisabled Then
      Exit For 
    End If
  Next

  LBD_IsSoundStyleDisabled = isDisabled
End Function

%>