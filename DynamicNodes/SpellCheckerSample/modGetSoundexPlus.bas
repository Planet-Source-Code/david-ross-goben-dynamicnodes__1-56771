Attribute VB_Name = "modGetSoundexPlus"
Option Explicit

'*******************************************************************************
' Function Name     : GetSoundPlus
' Purpose           : This is a special modification of the old Soundex routine
'*******************************************************************************
Public Function GetSoundPlus(Instring As String) As String
  Dim sInstring As String           'handle copy of input string
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
  Dim SoundexChar(8) As String      'Hold Characters to Check for relevant Soundex _
                                    code. SoundexChar(0) ignored for ease of use _
                                    when checking.
  Dim Counter As Long, F As Long
  
  'Assign values to Check
  SoundexChar(1) = "BFPV"     'Soundex Value of 1
  SoundexChar(2) = "CGJKQSXZ" 'Soundex Value of 2
  SoundexChar(3) = "DT"       'Soundex Value of 3
  SoundexChar(4) = "L"        'Soundex Value of 4
  SoundexChar(5) = "MN"       'Soundex Value of 5
  SoundexChar(6) = "R"        'Soundex Value of 6
  SoundexChar(6) = "R"        'Soundex Value of 6
  SoundexChar(7) = "EIY"      'Soundex Value of 6
  SoundexChar(8) = "OUW"      'Soundex Value of 6
'
'Convert the string to upper case letters and remove spaces
'
  sInstring = UCase$(Trim$(Instring))
'
'Strip Non-Alphabetic (Uppercase) Characters Ascii values btw B and Z (no vowels)
'
  sWorkSurname = vbNullString
  For Counter = 1 To Len(sInstring)
    If InStr(1, "BCDEFGIJKLMNOPQRSTUVWXYZ", Mid$(sInstring, Counter, 1)) Then
      sWorkSurname = sWorkSurname & Mid$(sInstring, Counter, 1)
    End If
  Next Counter
'
' check right end of word
'
  If CBool(Len(sWorkSurname)) Then
    F = Len(sWorkSurname)
    Do While F > 1 And InStr(1, "YIES", Right$(sWorkSurname, 1)) > 0
      sWorkSurname = Left$(sWorkSurname, F - 1)
      F = Len(sWorkSurname)
    Loop
    
    If F > 3 And Right$(sWorkSurname, 3) = "ING" Then
      sWorkSurname = Left$(sWorkSurname, F - 3)
      F = Len(sWorkSurname)
    End If
    
    If F > 2 And Right$(sWorkSurname, 2) = "ER" Then
      If InStr(1, "BCDFGJKLMNPQRSTVWXZ", Mid$(sWorkSurname, F - 2, 1)) Then
        sWorkSurname = Left$(sWorkSurname, F - 2)
        F = Len(sWorkSurname)
      End If
    End If
        
'''    If F > 2 And Right$(sWorkSurname, 2) = "IC" Then
'''      If InStr(1, "BCDFGJKLMNPQRSTVWXZ", Mid$(sWorkSurname, F - 2, 1)) Then
'''        sWorkSurname = Left$(sWorkSurname, F - 2)
'''        F = Len(sWorkSurname)
'''      End If
'''    End If
  End If
'
' if the result is null, then use 1st char and 7 zeros
'
  If Not CBool(F) Then
    GetSoundPlus = Left$(sInstring, 1) & "0000000"
    Exit Function
  End If
'
' get first character
'
  sSoundexCode = Left(sWorkSurname, 1)
'
' if left char of sWorkSurname is not same as orig, set orig
'
  If Left(sInstring, 1) <> sSoundexCode Then
    sSoundexCode = Left(sInstring, 1)
    sWorkSurname = sSoundexCode & sWorkSurname
  End If
'
' start scanning from the second character. Go until length is 8
'
  iCharacterCount = 2
  Do While Not Len(sSoundexCode) = 8
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
      sSoundexCode = sSoundexCode & String$(8 - Len(sSoundexCode), "0")
      Exit Do
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
      For F = 1 To 8
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
  GetSoundPlus = sSoundexCode
End Function

