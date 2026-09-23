<%

' Captcha code configuration
' ---------------------------------------------------------------------------
LBD_Configuration_CodeLength = 5
LBD_Configuration_CodeStyle = LBD_CodeStyles("Alphanumeric")
LBD_Configuration_CodeTimeout = 1200
LBD_Configuration_Locale = "en-US"
LBD_Configuration_CustomCharset = ""
LBD_Configuration_BannedSequences = ""

' Captcha image configuration
' ---------------------------------------------------------------------------
LBD_Configuration_ImageStyle = LBD_RandomImageStyle()
LBD_Configuration_ImageWidth = 250
LBD_Configuration_ImageHeight = 50
LBD_Configuration_ImageFormat = LBD_ImageFormats("JPEG")
LBD_Configuration_CustomDarkColor = ""
LBD_Configuration_CustomLightColor = ""
LBD_Configuration_ImageTooltip = "CAPTCHA"
LBD_Configuration_HelpLinkEnabled = True
LBD_Configuration_HelpLinkMode = LBD_HelpLinkModes("Image")
LBD_Configuration_HelpLinkUrl = ""
LBD_Configuration_HelpLinkText = ""

' Captcha sound configuration
' ---------------------------------------------------------------------------
LBD_Configuration_SoundEnabled = True
LBD_Configuration_SoundStyle = LBD_RandomSoundStyle()
LBD_Configuration_SoundFormat = LBD_SoundFormats("WavPcm16bit8kHzMono")
LBD_Configuration_SoundTooltip = "Speak the CAPTCHA code"
LBD_Configuration_SoundIconUrl = "BotDetect/SoundIcon.gif"
LBD_Configuration_SoundPackageFolder = "C:\Program Files\Lanapsoft\BotDetect 3 CAPTCHA Component\Asp\Redistribute\BotDetectSounds"
LBD_Configuration_WarnAboutMissingSoundPackages = True
LBD_Configuration_SoundStartDelay = 0

' Captcha reload configuration
' ---------------------------------------------------------------------------
LBD_Configuration_ReloadEnabled = True
LBD_Configuration_ReloadTooltip = "Reload the CAPTCHA code"
LBD_Configuration_ReloadIconUrl = "BotDetect/ReloadIcon.gif"
LBD_Configuration_AutoReloadExpiredCaptchas = True
LBD_Configuration_AutoReloadTimeout = 7200

' Captcha user input  configuration
' ---------------------------------------------------------------------------
LBD_Configuration_AutoFocusInput = True
LBD_Configuration_AutoClearInput = True
LBD_Configuration_AutoUppercaseInput = True

' Captcha remote script configuration
' ---------------------------------------------------------------------------
LBD_Configuration_RemoteScriptEnabled = True

' Captcha persistence configuration
' ---------------------------------------------------------------------------
Dim LBD_Persistence : Set LBD_Persistence = Session
LBD_Configuration_UseApplicationFallback = True

%>