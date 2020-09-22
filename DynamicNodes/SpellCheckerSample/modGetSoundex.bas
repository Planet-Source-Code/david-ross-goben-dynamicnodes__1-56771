Attribute VB_Name = "modGetSoundex"
Option Explicit
'~modGetSoundex.bas;
'Retrives the 4-character soundex code for a word
'*****************************************************************************
' modGetSoundex: The GetSoundex() function retrives the 4-character soundex code
'                for a word. This can be used when the user types in a response,
'                but the word is not a recognized response. The soundex code can
'                be compared the the soundex codes for known words, and those that
'                match (sound like), can be displayed in a list, for example.
'
' Convert Surname to Soundex Code. As detailed on the National Archives and
' Records Administration Web site. The census soundex coded surname is
' indexed on the way the Surname sounds rather than spelt. This allows for
' flexible searching on a surname. i. e. Smith, Smyth, Smythe all contain
' the same soundex code S530. This allows for finding surnames listed under
' different spelling variants.
'
' This code is based on the code by Darrell SPARTI dated 20-Sep-1998 which
' I found to record incorrect Soundex Code. In part this may be due to
' not functioning in VB6 correctly, but it did lack some of the claims
' it made regarding early functions. Noticably missing in this version was
' the rule regarding the letter "H" and "W", which if found separating two
' equally valued soundex consonants, the consonant to the right of the letter
' should not be coded.
'
' The coding has also been shortened, avoiding repetative code. Although it
' does appear quite long due to the amount of Remarks.
'*****************************************************************************

Public Function GetSoundex(Instring As String) As String
  Dim sInstring As String           'local processing of supplied word
  Dim sSoundexCode As String        'Holds the 4 character Soundex Code
  Dim sCurrentCharacter As String   'Current character being checked
  Dim sPreviousCharacter As String  'Previous character being checked
  Dim sSpecialCharacter As String   'Character previous to previous if current _
                                    character H or W This is for the special _
                                    check. If H or W seprates two two _
                                    consonants of the same value, the consonant _
                                    to the right of the vowel is not coded.
  Dim iCharacterCount As Integer    'Counter incremented as each character checked.
  Dim sWorkSurname As String        'Used during striping out not alphabetic _
                                    characters of InString Used to temporarily _
                                    hold surname value.
  Dim SoundexChar(6) As String      'Hold Characters to Check for relevant Soundex _
                                    code. SoundexChar(0) ignored for ease of use _
                                    when checking.
  Dim Counter As Integer, F As Integer
  
  'Assign values to Check
  SoundexChar(1) = "BFPV"     'Soundex Value of 1
  SoundexChar(2) = "CGJKQSXZ" 'Soundex Value of 2
  SoundexChar(3) = "DT"       'Soundex Value of 3
  SoundexChar(4) = "L"        'Soundex Value of 4
  SoundexChar(5) = "MN"       'Soundex Value of 5
  SoundexChar(6) = "R"        'Soundex Value of 6
'
'Convert the string to upper case letters and remove spaces
'
  sInstring = UCase$(Trim$(Instring))
'
'Strip Non-Alphabetic (Uppercase) Characters Ascii values btw B and Z (no vowels)
'
  For Counter = 1 To Len(sInstring)
    If InStr(1, "BCDFGJKLMNPQRSTVXZ", Mid$(sInstring, Counter, 1)) Then
      sWorkSurname = sWorkSurname & Mid$(sInstring, Counter, 1)
    End If
  Next Counter
  sInstring = sWorkSurname    'Reset words
'
'If surname of zero length end without conversion.
'
  If Len(sInstring) < 1 Then Exit Function
'
'The soundex code will start with the first character of the String
'
  sSoundexCode = Left(sWorkSurname, 1)
'
'Check the other characters starting at the second character
'
  iCharacterCount = 2
'
'Continue the conversion until the soundex code is 4 characters Long regarless of the length of the String
'
  Do While Not Len(sSoundexCode) = 4
'
'If the previous character has the same soundex code as current character
'or the previous character is the same as the current character, ignore it
'and move onto the Next. A special rule applies if "H" or "W" separate two
'consonants that have the same soundex code, the consonant to the right of
'H or W is not coded. Example Ashcroft is A261 not A226.
'
'If counter is greater than length of name being checked add "0"
'
    If iCharacterCount > Len(sWorkSurname) Then
      sSoundexCode = sSoundexCode & "0"
'
'Otherwise, concatenate a number to the soundex code base On soundex rules
'
    Else
      If iCharacterCount > 2 Then
        sSpecialCharacter = Mid$(sWorkSurname, iCharacterCount - 2, 1)
      End If
      sCurrentCharacter = Mid$(sWorkSurname, iCharacterCount, 1)
      sPreviousCharacter = Mid$(sWorkSurname, iCharacterCount - 1, 1)
'
'Check 6 variations
'
      For F = 1 To 6
        If InStr(1, SoundexChar(F), sCurrentCharacter) Then
          If InStr(1, SoundexChar(F), sPreviousCharacter) Then
            'Character has same soundex value as previous
          Else
            sSoundexCode = sSoundexCode & CStr(F) 'Increment Soundex
          End If 'Using value of F
          Exit For
        End If 'Nb: use of Trim
      Next F
    End If
    iCharacterCount = iCharacterCount + 1
  Loop
  GetSoundex = sSoundexCode
End Function

