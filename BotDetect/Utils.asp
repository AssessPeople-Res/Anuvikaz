<%

'A VbScript Array, containing all BotDetect code style names. String values
'are mapped by their numeric equivalents as indexes, so you can get a code
'style name by passing the numeric value as array index.
'E.g. LBD_CodeStyleNames(0) = "Alphanumeric".
Dim LBD_CodeStyleNames : LBD_CodeStyleNames = _
  Array("Alphanumeric", "Alpha", "Numeric")

'A VbScript Dictionary, containing all BotDetect code style numeric values.
'Numeric values are mapped by their (case-insensitive) string equivalents as
'keys, so you can get a code style numeric value by passing the string name
'as dictionary key.
'E.g. LBD_CodeStyles("Alphanumeric") = 0.
Dim LBD_CodeStyles : Set LBD_CodeStyles = _
  LBD_DictionaryFromArray(LBD_CodeStyleNames)


'A VbScript Array, containing all BotDetect image style names. String values
'are mapped by their numeric equivalents as indexes, so you can get a image
'style name by passing the numeric value as array index.
'E.g. LBD_ImageStyleNames(31) = "Chess".
Dim LBD_ImageStyleNames : LBD_ImageStyleNames = Array( _
  "Chess", "Distortion", "Jail", "Negative", "Snow", "Split", "Wave", _
  "WantedCircular", "Stitch", "Chess3D", "Circles", "Corrosion", "Chipped", _
  "Flash", "Mass", "Rough", "BlackOverlap", "Overlap", "Overlap2", "Halo", _
  "ThickThinLines", "ThickThinLines2", "Sunrays", "Sunrays2", "Darts", _
  "FingerPrints", "CrossShadow", "CrossShadow2", "Lego", "Strippy", _
  "ThinWavyLetters", "Chalkboard", "WavyColorLetters", "AncientMosaic", _
  "Vertigo", "WavyChess", "MeltingHeat", "SunAndWarmAir", "Graffiti", _
  "Graffiti2", "Cut", "SpiderWeb", "Collage", "InBandages", "Ghostly", _
  "PaintMess", "CaughtInTheNet", "CaughtInTheNet2", "Bullets", "Bullets2", _
  "Bubbles", "Electric", "MeltingHeat2", "Neon", "Neon2", "Radar", _
  "Ripple", "Ripple2", "SpiderWeb2", "Split2")

'A VbScript Dictionary, containing all BotDetect image style numeric values.
'Numeric values are mapped by their (case-insensitive) string equivalents as
'keys, so you can get a image style numeric value by passing the string name
'as dictionary key.
'E.g. LBD_ImageStyles("Chalkboard") = 31.
Dim LBD_ImageStyles : Set LBD_ImageStyles = _
  LBD_DictionaryFromArray(LBD_ImageStyleNames)


'A VbScript Array, containing all BotDetect sound style names. String values
'are mapped by their numeric equivalents as indexes, so you can get a sound
'style name by passing the numeric value as array index.
'E.g. LBD_SoundStyleNames(0) = "Dispatch".
Dim LBD_SoundStyleNames : LBD_SoundStyleNames = Array( _
  "Dispatch", "HiveMind", "Industrial", "Pulse", "Radio", _
  "RedAlert", "Robot", "Scratched", "Synth", "Workshop" )

'A VbScript Dictionary, containing all BotDetect sound style numeric values.
'Numeric values are mapped by their (case-insensitive) string equivalents as
'keys, so you can get a sound style numeric value by passing the string name
'as dictionary key.
'E.g. LBD_SoundStyles("Dispatch") = 0.
Dim LBD_SoundStyles : Set LBD_SoundStyles = _
  LBD_DictionaryFromArray(LBD_SoundStyleNames)


'A VbScript Array, containing all BotDetect image format names. String values
'are mapped by their numeric equivalents as indexes, so you can get a image
'format name by passing the numeric value as array index.
'E.g. LBD_ImageFormatNames(0) = "JPEG".
Dim LBD_ImageFormatNames : LBD_ImageFormatNames = _
  Array("JPEG", "GIF", "PNG", "BMP")

'A VbScript Dictionary, containing all BotDetect image format numeric values.
'Numeric values are mapped by their (case-insensitive) string equivalents as
'keys, so you can get an image format numeric value by passing the string name
'as dictionary key.
'E.g. LBD_ImageFormats("Jpeg") = 0.
Dim LBD_ImageFormats : Set LBD_ImageFormats = _
  LBD_DictionaryFromArray(LBD_ImageFormatNames)


'A VbScript Array, containing all BotDetect sound format names. String values
'are mapped by their numeric equivalents as indexes, so you can get a sound
'format name by passing the numeric value as array index.
'E.g. LBD_SoundFormatNames(0) = "WavPcm16bit8kHzMono".
Dim LBD_SoundFormatNames : LBD_SoundFormatNames = _
  Array("WavPcm16bit8kHzMono", "WavPcm8bit8kHzMono")

'A VbScript Dictionary, containing all BotDetect sound format numeric values.
'Numeric values are mapped by their (case-insensitive) string equivalents as
'keys, so you can get a sound format numeric value by passing the string name
'as dictionary key.
'E.g. LBD_SoundFormats( "WavPcm16bit8kHzMono") = 0.
Dim LBD_SoundFormats : Set LBD_SoundFormats = _
  LBD_DictionaryFromArray(LBD_SoundFormatNames)

'A VbScript Array, containing all help link mode names. String values
'are mapped by their numeric equivalents as indexes, so you can get a sound
'format name by passing the numeric value as array index.
'E.g. LBD_HelpLinkModeNames(0) = "Image".
Dim LBD_HelpLinkModeNames : LBD_HelpLinkModeNames = _
  Array("Image", "Text")

'A VbScript Dictionary, containing all help link mode numeric values.
'Numeric values are mapped by their (case-insensitive) string equivalents as
'keys, so you can get a sound format numeric value by passing the string name
'as dictionary key.
'E.g. LBD_HelpLinkModes("Image") = 0.
Dim LBD_HelpLinkModes : Set LBD_HelpLinkModes = _
  LBD_DictionaryFromArray(LBD_HelpLinkModeNames)
  
  

'Creates a simple timestamp string based on the current date and time,
'e.g. on the 21st of September at 13:45:27 it would return 20100921134527.
Function LBD_Timestamp()
  LBD_Timestamp = year(now) & right("0" & month(now),2) & _
    right("0" & day(now),2) & right("0" & hour(now),2) & _
    right("0" & minute(now),2) & right("0" & second(now),2)
End Function

'Formats the given date/time value in the YYYY/MM/DD hh:mm:ss format for
'logging & troubleshooting purposes.
Function LBD_FormattedDate(p_Date)
  LBD_FormattedDate = year(p_Date) & "/" & right("0" & month(p_Date),2) & _
    "/" & right("0" & day(p_Date),2) & " " & right("0" & hour(p_Date),2) & _
    ":" & right("0" & minute(p_Date),2) & ":" & right("0" & second(p_Date),2)
End Function

'Automatically uses the LBD_FormattedDate formatting on the current date/time.
Function LBD_Now()
  LBD_Now = LBD_FormattedDate(now)
End Function

'Replaces all non-Ascii characters in the given string with '?'.
Function LBD_Asciify(p_Input)
  Dim output : output = CStr(p_Input)
  Dim objRegExp : Set objRegExp = New RegExp
  objRegExp.Global = True
  objRegExp.IgnoreCase = True
  objRegExp.Pattern = "[^\x01-\x7E]"
  output = objRegExp.Replace(output, "?")
  LBD_Asciify = output
  Set objRegExp = Nothing
End Function

'Removes all non-word (word characters = alphanumeric Ascii + underscore)
'characters from the given string.
Function LBD_Normalize(p_Input)
  Dim output : output = LCase(CStr(p_Input))
  Dim objRegExp : Set objRegExp = New RegExp
  objRegExp.Global = True
  objRegExp.IgnoreCase = True
  objRegExp.Pattern = "\W"
  output = objRegExp.Replace(output, "")
  LBD_Normalize = output
  Set objRegExp = Nothing
End Function

'Removes all whitespace chars from the given string.
Function LBD_RemoveWhitespace(p_Input)
  Dim output : output = LCase(CStr(p_Input))
  Dim objRegExp : Set objRegExp = New RegExp
  objRegExp.Global = True
  objRegExp.IgnoreCase = True
  objRegExp.Pattern = "\s"
  output = objRegExp.Replace(output, "")
  LBD_RemoveWhitespace = output
  Set objRegExp = Nothing
End Function

'Creates a new GUID value, formatted as a continuous string
'(e.g. 9DAA001BF32B4B32AE741B109922ECC5).
Function LBD_CreateGuid()
  Dim TypeLib : Set TypeLib = Server.CreateObject("Scriptlet.TypeLib")
  Dim tg : tg = TypeLib.Guid
  Set TypeLib = Nothing
  Dim guid : guid = Left(tg, len(tg)-2)
  Dim objRegExp : Set objRegExp = New RegExp
  objRegExp.IgnoreCase = False
  objRegExp.Global = True
  objRegExp.Pattern = "[{}-]"
  LBD_CreateGuid = objRegExp.Replace(guid, "")
  Set objRegExp = Nothing
End Function

'Given an array of values, converts it to a VbScript Dictionary, with the
'array values becoming dictionary keys and array indexes becoming dictionary
'values.
'In other words, it reverses the array [index -> value] mapping, creating the
'equivalent [value -> index] dictionary mapping.
Function LBD_DictionaryFromArray(p_Array)
  Dim dict : Set dict = CreateObject("Scripting.Dictionary")
  dict.CompareMode = 1
  Dim i : i = 0
  For Each element In p_Array
    dict.Add element, i
    i = i + 1
  Next
  Set LBD_DictionaryFromArray = dict
End Function

'Returns a random integer from the numeric range defined by the given
'(inclusive) lower and upper bound.
'E.g. LBD_RandomFromRange(1, 5) could randomly return 1, 2, 3, 4 or 5.
Function LBD_RandomFromRange(p_LowerLimit, p_UpperLimit)
  Randomize
  Dim num : num = CInt((p_UpperLimit - p_LowerLimit)*Rnd() + p_LowerLimit)
  LBD_RandomFromRange = num
End Function

'Returns a random value from the given VbScript Array.
'E.g. LBD_RandomFromValues(Array("one", "two", "three")) could randomly
'return "one", "two" or "three".
Function LBD_RandomFromValues(p_Values)
  Randomize
  Dim max : max = UBound(p_Values)
  Dim num : num = LBD_RandomFromRange(0, max)
  LBD_RandomFromValues = p_Values(num)
End Function

'Name of the currently executing script file,
'e.g. returns "Default" when called from Default.asp.
Function LBD_ScriptName
  Dim fullname : fullname = Request.ServerVariables("SCRIPT_NAME")
  Dim parts : parts = split(fullname, "/")
  Dim count : count = UBound(parts)
  Dim scriptname : scriptname = parts(count)
  LBD_ScriptName = Left(scriptname, len(scriptname)-4)
End Function


'Name of the next-to-last path fragment for the currently executing script
'file, e.g. returns "Folder2" when called from
'http://localhost/Folder1/Folder2/Default.asp.
Function LBD_AppName
  Dim fullname : fullname = Request.ServerVariables("SCRIPT_NAME")
  Dim parts : parts = split(fullname, "/")
  Dim count : count = UBound(parts)
  Dim appname
  If count > 0 Then
    appname = parts(count-1)
  Else
    appname = "root"
  End If
  LBD_AppName = appname
End Function

%>