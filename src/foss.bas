Option Explicit

'**********************************************************************************************************
' Name:                 endsWith
' Author:               mielk | 2012-06-21
'
' Comment:              Checks if the given string ends with the specified substrign.
'
' Parameters:
'   str                 String to be checked.
'   substr              Substring to be checked.
'   isCaseSensitive     Optional parameter of Boolean type.
'                       It determines if text matching is case sensitive.
'                       If this value is set to True, searching is case sensitive - a letter in lowercase
'                       is treated as different than the same letter in uppercase (i.e. a <> A).
'                       If this value is set to False, it doesn't matter if a letter is in lowercase or in
'                       uppercase, since both of them are considered as the same character (i.e. a = A).
'                       Default value of this parameter is True.
'
' Returns:
'   Boolean             True - if string [str] ends with the given substring.
'                       False - otherwise.
'
'
' --- Changes log -----------------------------------------------------------------------------------------
' 2012-06-21        mielk           Function created.
'**********************************************************************************************************
' http://www.mielk.pl/en/download/code/text/endsWith.php
Public Function endsWith(str As String, substr As String, Optional isCaseSensitive As Boolean = True) _
                                                                                                 As Boolean
    Const METHOD_NAME As String = "endsWith"
    '------------------------------------------------------------------------------------------------------
    Dim uCompareMethod As VBA.VbCompareMethod
    '------------------------------------------------------------------------------------------------------


    'Convert [isCaseSensitive] parameter of Boolean type to the [VbCompareMethod] enumeration. ----------|
    If isCaseSensitive Then                                                                             '|
        uCompareMethod = VBA.vbBinaryCompare                                                            '|
    Else                                                                                                '|
        uCompareMethod = VBA.vbTextCompare                                                              '|
    End If                                                                                              '|
    '----------------------------------------------------------------------------------------------------|


    '----------------------------------------------------------------------------------------------------|
    If VBA.StrComp(VBA.Right$(str, VBA.Len(substr)), substr, uCompareMethod) = 0 Then                   '|
        endsWith = True                                                                                 '|
    End If                                                                                              '|
    '----------------------------------------------------------------------------------------------------|

End Function
