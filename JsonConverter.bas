Option Explicit
'=======================================================================================================
' JSON String Converter
'-------------------------------------------------------------------------------------------------------
' Author(s)   :
'       Ryo Yokoyama
'       Tim Hall
' Contributors:
'       Sanbi
' Last Update (xx/xx/2025)
'       - Fixed parse number
'       - Fixed decode \xXX for jsonParseString
'       - Use class StringBuffer (Cristian Buse)- https://github.com/cristianbuse/VBA-StringBuffer
'       - Use class Dictionary (Cristian Buse)- https://github.com/cristianbuse/VBA-FastDictionary


'-----------------------------------------------------------------------------------------------------------
' Notes       :
'       This class should only be initiated internally by the CDPBrowser class and should not need to be
'       initiated directly at any times.
' References  :
'       Nil
' Sources     :
'       Tim Hall: github.com/VBA-tools/VBA-JSON
'=======================================================================================================
 
' VBA-JSON (c) Tim Hall - github.com/VBA-tools/VBA-JSON

' Errors:
' 10001 - JSON parse error
'
' @class CDPJsonConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on vba-json (with extensive changes)
' BSD license included below
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
 
'===================================
' Win APIs Declarations
'===================================
 

' === VBA-UTC Headers
#If Mac Then

#If VBA7 Then

' 64-bit Mac (2016)
Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
Private Declare PtrSafe Function utc_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" _
    (ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_fread Lib "/usr/lib/libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As LongPtr, ByVal utc_Number As LongPtr, ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_feof Lib "/usr/lib/libc.dylib" Alias "feof" _
    (ByVal utc_File As LongPtr) As LongPtr

#Else

' 32-bit Mac
Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" _
    (ByVal utc_File As Long) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As Long, ByVal utc_Number As Long, ByVal utc_File As Long) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" _
    (ByVal utc_File As Long) As Long

#End If

#ElseIf VBA7 Then

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#Else

Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#End If

#If Mac Then

#If VBA7 Then
Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As LongPtr
End Type

#Else

Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#End If

#Else

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

#End If
'===================================
' Types & Structures Declarations
'===================================

' VBA only stores 15 significant digits, so any numbers larger than that are truncated
' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
' See: http://support.microsoft.com/kb/269370
Private Type json_Options
 
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `CDPJsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean
 
    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    ' Tiï¿½u chuï¿½ï¿½n JSON yï¿½u cï¿½ï¿½u cï¿½c khï¿½a ï¿½ï¿½ï¿½i tï¿½ï¿½ï¿½ng phaï¿½i ï¿½ï¿½ï¿½ï¿½c trï¿½ch dï¿½ï¿½n ( " hoï¿½ï¿½c ' ), haï¿½y sï¿½ï¿½ duï¿½ng tï¿½y choï¿½n nï¿½y ï¿½ï¿½ï¿½ cho phï¿½p cï¿½c khï¿½a khï¿½ng ï¿½ï¿½ï¿½ï¿½c trï¿½ch dï¿½ï¿½n
    AllowUnquotedKeys As Boolean
 
    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
    EscapeSolidus As Boolean
    keyEscapeSolidus As Long
    CompareMode As Boolean
End Type
 
Private JsonOptions As json_Options
Private tJson$, lJson&, vv$, idx&, t$, s2d As Boolean
' ============================================= '
' Class Functions
' ============================================= '
Public Property Let setCompareMode(b As Boolean)
  JsonOptions.CompareMode = b
End Property
Public Property Let setAllowUnquotedKeys(b As Boolean)
  JsonOptions.AllowUnquotedKeys = b
End Property
Public Property Let setUseDoubleForLargeNumbers(b As Boolean)
  JsonOptions.UseDoubleForLargeNumbers = b
End Property
Public Property Let setEscapeSolidus(b As Boolean)
  JsonOptions.EscapeSolidus = b
  JsonOptions.keyEscapeSolidus = -b * 47
End Property
Public Function removeObjectKeys(ByVal dict As Object, Optional delimiter$ = ",", Optional backet$ = "'", Optional skipNull As Boolean) As String
  Dim i, s As New StringBuffer, s0$, v$, X$
  X = IIf(skipNull, "", "null")
  For Each i In dict.Keys()
    If IsObject(dict(i)) Then v = backet & ConvertToJson(dict(i)) & backet Else v = ConvertToJson(dict(i), , , backet = "'")
    s.Append s0$ & v: s0 = delimiter
  Next
  removeObjectKeys = s.value
End Function
Public Function ParseJson(ByVal JsonString As String, _
               Optional ByVal StringToDate As Boolean = True, _
               Optional ByVal EscapeSolidus As Boolean = True, _
               Optional ByVal CompareMode As Boolean = True, _
               Optional ByVal ErrRaise As Boolean = False) As Object
'----------------------------------------------------------------
' Convert JSON string to object (Dictionary/Collection)
' @method ParseJson
' @param {String} tjson
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
'----------------------------------------------------------------
  idx = 0: lJson = Len(JsonString): tJson = JsonString: s2d = StringToDate
  setEscapeSolidus = EscapeSolidus
  JsonOptions.CompareMode = CompareMode
  'Remove vbCr, vbLf, and vbTab from tjson
  jsonSkipSpaces
  Select Case vv
  Case "{": Set ParseJson = jsonParseObject
  Case "[": Set ParseJson = jsonParseArray
  Case Else:  If ErrRaise Then Err.Raise 10001, "JSONConverter", jsonParseErrorMessage("Expecting '{' or '['")   'Error: Invalid JSON string
  End Select
End Function
Public Function ConvertToJson(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal jCurrentIndentation As Long = 0, Optional ByVal stringSingleQuote% = 2) As String
  '--------------------------------------------------------------------------------------------------------------------------------
  ' Convert object (Dictionary/Collection/Array) to JSON
  ' @method ConvertToJson
  ' @param {Variant} JsonValue (Dictionary, Collection, or Array)
  ' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
  ' @return {String}
  '--------------------------------------------------------------------------------------------------------------------------------
  Dim jBufferPosition As Long, jBufferLength As Long, idx As Long
  Dim jLBound As Long, jUBound As Long, jIsFirstItem As Boolean
  Dim jIndex2D As Long, jLBound2D As Long, jUBound2D As Long, jIsFirstItem2D As Boolean
  Dim jKey As Variant, jValue As Variant, jDateStr As String, jConverted As String
  Dim jSkipItem As Boolean, jPrettyPrint As Boolean, jIndentation As String
  Dim jInnerIndentation As String, jBuffer As New StringBuffer, ss$
  Select Case stringSingleQuote
  Case 0:
  Case 1: ss = "'"
  Case 2: ss = """"
  End Select
  jLBound = -1: jUBound = -1: jIsFirstItem = True
  jLBound2D = -1: jUBound2D = -1: jIsFirstItem2D = True
  jPrettyPrint = Not IsMissing(Whitespace)
  Select Case VarType(JsonValue)
  Case vbNull: ConvertToJson = "null"
  Case vbDate: jDateStr = ConvertToIso(CDate(JsonValue)): ConvertToJson = """" & jDateStr & """"
  Case vbString   ' String (or large number encoded as string)
    If Not JsonOptions.UseDoubleForLargeNumbers And jsonStringIsLargeNumber(JsonValue) Then
      ConvertToJson = JsonValue
    Else
      ConvertToJson = ss & jsonEncode(JsonValue) & ss
    End If
  Case vbBoolean: If JsonValue Then ConvertToJson = "true" Else ConvertToJson = "false"
  Case vbArray To vbArray + vbByte
    If jPrettyPrint Then
      If VarType(Whitespace) = vbString Then
        jIndentation = String$(jCurrentIndentation + 1, Whitespace)
        jInnerIndentation = String$(jCurrentIndentation + 2, Whitespace)
      Else
        jIndentation = Space$((jCurrentIndentation + 1) * Whitespace)
        jInnerIndentation = Space$((jCurrentIndentation + 2) * Whitespace)
      End If
    End If
    ' Array
    jBuffer.Append "["
    On Error Resume Next
    jLBound = LBound(JsonValue): jUBound = UBound(JsonValue): jLBound2D = LBound(JsonValue, 2): jUBound2D = UBound(JsonValue, 2)
    If jLBound >= 0 And jUBound >= 0 Then
      For idx = jLBound To jUBound
        If jIsFirstItem Then jIsFirstItem = False Else jBuffer.Append ","
        If jLBound2D >= 0 And jUBound2D >= 0 Then
          ' 2D Array
          If jPrettyPrint Then jBuffer.Append vbNewLine
          jBuffer.Append jIndentation & "["
          For jIndex2D = jLBound2D To jUBound2D
            If jIsFirstItem2D Then jIsFirstItem2D = False Else jBuffer.Append ","
            jConverted = ConvertToJson(JsonValue(idx, jIndex2D), Whitespace, jCurrentIndentation + 2)
            ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
            If jConverted = "" Then
              ' (nest to only check if converted = "")
              If jsonIsUndefined(JsonValue(idx, jIndex2D)) Then jConverted = "null"
            End If
            If jPrettyPrint Then jConverted = vbNewLine & jInnerIndentation & jConverted
            jBuffer.Append jConverted
          Next jIndex2D
          If jPrettyPrint Then jBuffer.Append vbNewLine
          jBuffer.Append jIndentation & "]"
          jIsFirstItem2D = True
        Else
          ' 1D Array
          jConverted = ConvertToJson(JsonValue(idx), Whitespace, jCurrentIndentation + 1)
          ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
          If jConverted = "" Then
            ' (nest to only check if converted = "")
            If jsonIsUndefined(JsonValue(idx)) Then jConverted = "null"
          End If
          If jPrettyPrint Then jConverted = vbNewLine & jIndentation & jConverted
          jBuffer.Append jConverted
        End If
      Next idx
    End If
    On Error GoTo 0
    If jPrettyPrint Then
      jBuffer.Append vbNewLine
      If VarType(Whitespace) = vbString Then
        jIndentation = String$(jCurrentIndentation, Whitespace)
      Else
        jIndentation = Space$(jCurrentIndentation * Whitespace)
      End If
    End If
    jBuffer.Append jIndentation & "]"
    ConvertToJson = jBuffer.value
    ' Dictionary or Collection
  Case vbObject
    If jPrettyPrint Then
      If VarType(Whitespace) = vbString Then
        jIndentation = String$(jCurrentIndentation + 1, Whitespace)
      Else
        jIndentation = Space$((jCurrentIndentation + 1) * Whitespace)
      End If
    End If
    ' Dictionary
    If TypeName(JsonValue) = "Dictionary" Then
      jBuffer.Append "{"
      For Each jKey In JsonValue.Keys
        ' For Objects, undefined (Empty/Nothing) is not added to object
        jConverted = ConvertToJson(JsonValue(jKey), Whitespace, jCurrentIndentation + 1)
        If jConverted = "" Then
          jSkipItem = jsonIsUndefined(JsonValue(jKey))
        Else
          jSkipItem = False
        End If
        If Not jSkipItem Then
          If jIsFirstItem Then jIsFirstItem = False Else jBuffer.Append ","
          If jPrettyPrint Then
            jConverted = vbNewLine & jIndentation & """" & jKey & """: " & jConverted
          Else
            jConverted = """" & jKey & """:" & jConverted
          End If
          jBuffer.Append jConverted
        End If
      Next jKey
      If jPrettyPrint Then
        jBuffer.Append vbNewLine
        If VarType(Whitespace) = vbString Then
          jIndentation = String$(jCurrentIndentation, Whitespace)
        Else
          jIndentation = Space$(jCurrentIndentation * Whitespace)
        End If
      End If
      jBuffer.Append jIndentation & "}"
      ' Collection
    ElseIf TypeName(JsonValue) = "Collection" Then
      jBuffer.Append "["
      For Each jValue In JsonValue
        If jIsFirstItem Then jIsFirstItem = False Else jBuffer.Append ","
        jConverted = ConvertToJson(jValue, Whitespace, jCurrentIndentation + 1)
        ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
        If jConverted = "" Then
          ' (nest to only check if converted = "")
          If jsonIsUndefined(jValue) Then jConverted = "null"
        End If
        If jPrettyPrint Then jConverted = vbNewLine & jIndentation & jConverted
        jBuffer.Append jConverted
      Next jValue
      If jPrettyPrint Then
        jBuffer.Append vbNewLine
        If VarType(Whitespace) = vbString Then
          jIndentation = String$(jCurrentIndentation, Whitespace)
        Else
          jIndentation = Space$(jCurrentIndentation * Whitespace)
        End If
      End If
      jBuffer.Append jIndentation & "]"
    End If
    ConvertToJson = jBuffer.value
  Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal: ConvertToJson = Replace$(JsonValue, ",", ".")
  Case Else
    ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
    ' Use VBA's built-in to-string
    On Error Resume Next
    ConvertToJson = JsonValue
    On Error GoTo 0
  End Select
End Function
Private Function jsonParseObject() As Object
  Set jsonParseObject = Interaction.CreateObject("Scripting.Dictionary"): jsonParseObject.CompareMode = -JsonOptions.CompareMode
  Dim k$
  Do
    k = jsonParseKey(): If k = "" Then Exit Function
    jsonSkipSpaces
    If vv Like "[{[]" Then Set jsonParseObject.item(k) = jsonParseValue Else jsonParseObject.item(k) = jsonParseValue
    jsonSkipSpaces
    Select Case vv
    Case "}": Exit Function
    Case ",":
    Case " ", vbCr, vbLf, vbTab: jsonSkipSpaces
    Case Else: Err.Raise 10001, "JSONConverter", jsonParseErrorMessage("Expecting Json Object")
    End Select
  Loop Until idx >= lJson
End Function
Private Function jsonParseArray() As Collection
  Set jsonParseArray = New Collection
  Do
    jsonSkipSpaces
    Select Case vv
    Case "]": Exit Function
    Case ",": jsonSkipSpaces
    End Select
    jsonParseArray.add jsonParseValue
  Loop Until idx >= lJson
End Function
Private Function jsonParseValue() As Variant
  Select Case vv
  Case "{": Set jsonParseValue = jsonParseObject
  Case "[": Set jsonParseValue = jsonParseArray
  Case """", "'": vv = jsonParseString(vv)
    If Not s2d Then
      jsonParseValue = vv
    Else
      Select Case True
      Case vv Like "####[-/]##[-/]##T##:##:##.###*", _
           vv Like "####[-/]##[-/]##T##:##:##Z", _
           vv Like "####[-/]##[-/]##T##:##:##": jsonParseValue = ParseIso(vv)
      Case vv Like "####[-/]##[-/]##", vv Like "####[-/]##[-/]## ##:##:##": jsonParseValue = CDate(vv)
      Case Else: jsonParseValue = vv
      End Select
    End If
  Case "0" To "9": jsonParseValue = json_ParseNumber(vv)
  Case "+", "-": idx = idx + 1: t = Mid$(tJson, idx, 1): If Not t Like "#" Then Err.Raise 10001, "JSONConverter", jsonParseErrorMessage("Expecting 'NUMBER'")
    jsonParseValue = json_ParseNumber(vv & t)
  Case Else
    Select Case True
    Case Mid$(tJson, idx, 4) Like "true": jsonParseValue = True: idx = idx + 3
    Case Mid$(tJson, idx, 5) Like "false": jsonParseValue = False: idx = idx + 4
    Case Mid$(tJson, idx, 4) Like "null": jsonParseValue = Null: idx = idx + 3
    Case Else: Err.Raise 10001, "JSONConverter", jsonParseErrorMessage("Expecting value json")
    End Select
  End Select
End Function
 
Private Function jsonParseString(ByVal json_Quote$) As String
  With New StringBuffer
    Do
      idx = idx + 1: vv = Mid$(tJson, idx, 1)
      Select Case vv
      Case "\"
        ' Escaped string, \\, or \/
        idx = idx + 1: vv = Mid$(tJson, idx, 1)
        Select Case vv
        Case """", "\", "/", "'": .Append vv
        Case "b": .Append vbBack
        Case "f": .Append vbFormFeed
        Case "n": .Append vbCrLf
        Case "r": .Append vbCr
        Case "t": .Append vbTab
        Case "x"
          idx = idx + 1: t = Mid$(tJson, idx, 2)
          If Not t Like "[0-9A-Fa-f][0-9A-Fa-f]" Then Err.Raise 10001, "JSONConverter", jsonParseErrorMessage("Expecting Encode character: " & "&H" + t)
          .Append ChrW(Val("&H" + t)): idx = idx + 1
        Case "u"
          idx = idx + 1: t = Mid$(tJson, idx, 4)
          If Not t Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then Err.Raise 10001, "JSONConverter", jsonParseErrorMessage("Expecting Encode character: " & "&H" + t)
          .Append ChrW(Val("&H" + t)): idx = idx + 3
        End Select
      Case json_Quote: jsonParseString = .value: Exit Function
      Case Else: .Append vv
      End Select
    Loop Until idx >= lJson
  End With
End Function
 
Private Function json_ParseNumber(ByVal value$) As Variant
  Dim json_Char As String, d&, t$
  Dim json_IsLargeNumber As Boolean
  Do: idx = idx + 1: vv = Mid$(tJson, idx, 1)
    Select Case d
    Case 0:
      Select Case vv
      Case "0" To "9":
      Case ".": idx = idx + 1: t = Mid$(tJson, idx, 1): If Not t Like "#" Then GoTo ee
        d = 1: vv = vv & t
      Case "e", "E":
eee:
        idx = idx + 1: t = Mid$(tJson, idx, 1)
        Select Case t
        Case "0" To "9":
        Case "+": idx = idx + 1: t = t & Mid$(tJson, idx, 1): If Not t Like "+#" Then GoTo ee
        Case Else: GoTo ee
        End Select
        d = 2: vv = vv & t
      Case ",", "}", "]", " ", vbCr, vbLf, vbTab:  idx = idx - 1: Exit Do
      Case Else: GoTo ee
      End Select
    Case 1, 3:
      Select Case vv
      Case "0" To "9": d = 3
      Case "e", "E": GoTo eee
      Case ",", "}", "]", " ", vbCr, vbLf, vbTab: idx = idx - 1: Exit Do
      Case Else: GoTo ee
      End Select
    Case 2, 4:
      Select Case vv
      Case "0" To "9": d = 4
      Case ",", "}", "]", " ", vbCr, vbLf, vbTab: idx = idx - 1: Exit Do
      Case Else: GoTo ee
      End Select
    End Select
    value = value & vv
  Loop Until idx >= lJson

  json_IsLargeNumber = IIf(InStr(value, "."), Len(value) >= 17, Len(value) >= 16)
  If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
    json_ParseNumber = value
  Else
    ' Val does not use regional settings, so guard for comma is not needed
    json_ParseNumber = Val(value)
  End If
Exit Function
ee:
   Err.Raise 10001, "JSONConverter", jsonParseErrorMessage("Expecting 'NUMBER'")
End Function
 
Private Function jsonParseKey() As String
  jsonSkipSpaces
  
  If vv = "}" Then
    Exit Function
  ElseIf vv Like "[""']" Then
    jsonParseKey = jsonParseString(vv)
  ElseIf JsonOptions.AllowUnquotedKeys Then
    Do
      Select Case vv
      Case ":", " ", vbCr, vbLf, vbTab: Exit Do
      Case Else: jsonParseKey = jsonParseKey & vv: idx = idx + 1: vv = Mid$(tJson, idx, 1)
      End Select
    Loop Until idx > lJson
  Else
    Err.Raise 10001, "JSONConverter", jsonParseErrorMessage("Expecting '""' or '''")
  End If
  ' Check for colon and skip if present or throw if not present
  jsonSkipSpaces
  If vv <> ":" Or jsonParseKey = "" Or idx >= lJson Then Err.Raise 10001, "JSONConverter", jsonParseErrorMessage("Expecting Key:Value")
End Function
 
 
Private Function jsonIsUndefined(ByVal json_Value As Variant) As Boolean
  ' Empty / Nothing -> undefined
  Select Case VarType(json_Value)
  Case vbEmpty: jsonIsUndefined = True
  Case vbObject
    Select Case TypeName(json_Value)
    Case "Empty", "Nothing": jsonIsUndefined = True
    End Select
  End Select
End Function
 
 
Private Function jsonEncode(ByVal json As Variant) As String
  ' Reference: http://www.ietf.org/rfc/rfc4627.txt
  ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
  Dim i As Long, s As String, m As Long, b As String, json_Buffer As New StringBuffer
  For i = 1 To Len(json)
    ' When AscW returns a negative number, it returns the twos complement form of that number.
    ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
    ' support.microsoft.com/en-us/kb/272138
    s = Mid$(json, i, 1): m = AscW(s) And &HFFFF&
    ' From spec, ", \, and control characters must be escaped (solidus is optional)
    Select Case m
    Case 34: s = "\"""
    Case 92: s = "\\"
    Case 8: s = "\b"
    Case 12: s = "\f"
    Case 13: s = "\r"
    Case 10: s = "\n"
    Case 9: s = "\t"
    Case 0 To 31, 127 To 65535
      ' Non-ascii characters -> convert to 4-digit hex
      s = "\u" & Right$("0000" & LCase(Hex$(m)), 4)
    Case JsonOptions.keyEscapeSolidus: s = "\/"
    End Select
    json_Buffer.Append s
  Next i
  jsonEncode = json_Buffer.value
End Function
 
Private Sub jsonSkipSpaces()
  Do
    idx = idx + 1: vv$ = Mid$(tJson, idx, 1)
    Select Case vv
    Case " ", vbCr, vbLf, vbTab:
    Case Else: Exit Sub
    End Select
  Loop Until idx >= lJson
  vv = vbNullString
End Sub
 
 
Private Function jsonStringIsLargeNumber(tJson As Variant) As Boolean
  ' Check if the given string is considered a "large number"
  ' (See json_ParseNumber)
  Dim json_Length As Long
  Dim json_CharIndex As Long
  json_Length = lJson
  ' Length with be at least 16 characters and assume will be less than 100 characters
  If json_Length >= 16 And json_Length <= 100 Then
    Dim json_CharCode As String
    jsonStringIsLargeNumber = True
    For json_CharIndex = 1 To json_Length
      json_CharCode = Asc(Mid$(tJson, json_CharIndex, 1))
      Select Case json_CharCode
        ' Look for .|0-9|E|e
      Case 46, 48 To 57, 69, 101: ' Continue through characters
      Case Else
        jsonStringIsLargeNumber = False
        Exit Function
      End Select
    Next json_CharIndex
  End If
End Function
 
 
Private Function jsonParseErrorMessage(ErrorMessage As String)
'------------------------------------------------------------------------------------
' Provide detailed parse error message, including details of where and what occurred
'
' Example:
' Error parsing JSON:
' {"abcde":True}
'          ^
' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['
'------------------------------------------------------------------------------------
 
    Dim json_StartIndex As Long
    Dim json_StopIndex As Long
    ' Include 10 characters before and after error (if possible)
    json_StartIndex = idx - 10
    json_StopIndex = idx + 10
    If json_StartIndex < 1 Then json_StartIndex = 1
    If json_StopIndex > lJson Then json_StopIndex = lJson

    jsonParseErrorMessage = "Error parsing JSON:" & vbNewLine & _
                             Mid$(tJson, json_StartIndex, json_StopIndex - json_StartIndex + 1) & vbNewLine & _
                             Space$(idx - json_StartIndex) & "^" & vbNewLine & _
                             ErrorMessage
   
    Debug.Print jsonParseErrorMessage
End Function
 

Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
 
    If json_BufferPosition > 0 Then
        json_BufferToString = Left$(json_Buffer, json_BufferPosition)
    End If
    
End Function
 
''
' VBA-UTC v1.0.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
'
' @module UtcConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @method ParseUtc
' @param {Date} UtcDate
' @return {Date} Local date
' @throws 10011 - UTC parsing error
''
Public Function ParseUtc(utc_UtcDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ParseUtc = utc_ConvertDate(utc_UtcDate)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_LocalDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate

    ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
#End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.number & " - " & Err.description
End Function

''
' Convert local date to UTC date
'
' @method ConvertToUrc
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' @throws 10012 - UTC conversion error
''
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate

    ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
#End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.number & " - " & Err.description
End Function


''
' Parse ISO 8601 date string to local date
'
' @method ParseIso
' @param {Date} utc_IsoString
' @return {Date} Local date
' @throws 10013 - ISO 8601 parsing error
''
 
Private Function ParseIso(utc_IsoString As String) As Date
  '-----------------------------------------------------
  ' Parse ISO 8601 date string to local date
  ' @method ParseIso
  ' @param {Date} utc_IsoString
  ' @return {Date} Local date
  ' @throws 10013 - ISO 8601 parsing error
  '-----------------------------------------------------
  On Error GoTo utc_ErrorHandling
  Dim utc_Parts() As String
  Dim utc_DateParts() As String
  Dim utc_TimeParts() As String
  Dim utc_OffsetIndex As Long
  Dim utc_HasOffset As Boolean
  Dim utc_NegativeOffset As Boolean
  Dim utc_OffsetParts() As String
  Dim utc_Offset As Date
  utc_DateParts = Split(utc_IsoString, "-")
  ParseIso = DateSerial(Val(utc_IsoString), Val(utc_DateParts(1)), Val(utc_DateParts(2)))
  utc_Parts = Split(utc_IsoString, "T")
  If UBound(utc_Parts) > 0 Then
    If InStr(utc_Parts(1), "Z") Then
      utc_TimeParts = Split(Replace(utc_Parts(1), "Z", ""), ":")
    Else
      utc_OffsetIndex = InStr(1, utc_Parts(1), "+")
      If utc_OffsetIndex = 0 Then
        utc_NegativeOffset = True: utc_OffsetIndex = InStr(1, utc_Parts(1), "-")
      End If
      If utc_OffsetIndex > 0 Then
        utc_HasOffset = True
        utc_TimeParts = Split(Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
        utc_OffsetParts = Split(Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")
        Select Case UBound(utc_OffsetParts)
        Case 0: utc_Offset = TimeSerial(CInt(utc_OffsetParts(0)), 0, 0)
        Case 1: utc_Offset = TimeSerial(CInt(utc_OffsetParts(0)), CInt(utc_OffsetParts(1)), 0)
        Case 2
          ' Val does not use regional settings, use for seconds to avoid decimal/comma issues
          utc_Offset = TimeSerial(CInt(utc_OffsetParts(0)), CInt(utc_OffsetParts(1)), Int(Val(utc_OffsetParts(2))))
        End Select
        If utc_NegativeOffset Then: utc_Offset = -utc_Offset
      Else
        utc_TimeParts = Split(utc_Parts(1), ":")
      End If
    End If
    Select Case UBound(utc_TimeParts)
    Case 0: ParseIso = ParseIso + TimeSerial(CInt(utc_TimeParts(0)), 0, 0)
    Case 1: ParseIso = ParseIso + TimeSerial(CInt(utc_TimeParts(0)), CInt(utc_TimeParts(1)), 0)
    Case 2
      ' Val does not use regional settings, use for seconds to avoid decimal/comma issues
      ParseIso = ParseIso + TimeSerial(CInt(utc_TimeParts(0)), CInt(utc_TimeParts(1)), Int(Val(utc_TimeParts(2))))
    End Select
    ParseIso = ParseUtc(ParseIso)
    If utc_HasOffset Then ParseIso = ParseIso - utc_Offset
  End If
  Exit Function
utc_ErrorHandling:
  Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & Err.number & " - " & Err.description
End Function
 
 
''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIso(utc_LocalDate As Date) As String
'-------------------------------------------------------------
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
'-------------------------------------------------------------
 
    On Error GoTo utc_ErrorHandling
    ConvertToIso = Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")
    Exit Function
 
utc_ErrorHandling:
 
    Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.number & " - " & Err.description
    
End Function
 
 ' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then

Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
    Dim utc_ShellCommand As String
    Dim utc_Result As utc_ShellResult
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String

    If utc_ConvertToUtc Then
        utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
            " +'%s'` +'%Y-%m-%d %H:%M:%S'"
    Else
        utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
            "+'%Y-%m-%d %H:%M:%S'"
    End If

    utc_Result = utc_ExecuteInShell(utc_ShellCommand)

    If utc_Result.utc_Output = "" Then
        Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
    Else
        utc_Parts = Split(utc_Result.utc_Output, " ")
        utc_DateParts = Split(utc_Parts(0), "-")
        utc_TimeParts = Split(utc_Parts(1), ":")

        utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
            TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
    End If
End Function

Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult
#If VBA7 Then
    Dim utc_File As LongPtr
    Dim utc_Read As LongPtr
#Else
    Dim utc_File As Long
    Dim utc_Read As Long
#End If

    Dim utc_Chunk As String

    On Error GoTo utc_ErrorHandling
    utc_File = utc_popen(utc_ShellCommand, "r")

    If utc_File = 0 Then: Exit Function

    Do While utc_feof(utc_File) = 0
        utc_Chunk = VBA.Space$(50)
        utc_Read = CLng(utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_File))
        If utc_Read > 0 Then
            utc_Chunk = VBA.Left$(utc_Chunk, CLng(utc_Read))
            utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
        End If
    Loop

utc_ErrorHandling:
    utc_ExecuteInShell.utc_ExitCode = CLng(utc_pclose(utc_File))
End Function

#Else

Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
    utc_DateToSystemTime.utc_wYear = VBA.year(utc_Value)
    utc_DateToSystemTime.utc_wMonth = VBA.month(utc_Value)
    utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
    utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
    utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
    utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
    utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function

#End If


