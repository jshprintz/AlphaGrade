VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QA_Checklist 
   Caption         =   "QA Checklist"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   OleObjectBlob   =   "QA_Checklist.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QA_Checklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCurrentRng As Range
Private i, n, mResearchCount, mPCount, mWERcount, mVarCntOne, mVarCntTwo, mVarCntThree, mVarCntFour, mVarCntFive, mVarCntSix, mVarCntSeven, mFinalMWCount As Integer
Private Response As Variant
Private mRWCount, mstrIHRep As String
Private aVar As Variable
Private mTagDisp As Boolean

Private Sub cmdBack_Click()

'Resets the Main Menu

PreGrade_explain.Visible = False
PostGrade_explain.Visible = False
ClientReady_explain.Visible = False
TagCount_explain.Visible = False
Links_explain.Visible = False

cmdContinueOne.Locked = True
cmdContinueOne.Visible = False
cmdContinueTwo.Locked = True
cmdContinueTwo.Visible = False
cmdContinueThree.Locked = True
cmdContinueThree.Visible = False
cmdContinueFour.Locked = True
cmdContinueFour.Visible = False

cmdBack.BackColor = vbButtonFace

Post_Grade.Locked = False
Post_Grade.Visible = True
Pre_Grade.Locked = False
Pre_Grade.Visible = True
Client_Ready.Locked = False
Client_Ready.Visible = True
QA_Tag_Count.Locked = False
QA_Tag_Count.Visible = True
Tools.Locked = False
Tools.Visible = True
Links.Locked = False
Links.Visible = True
HelpClick.Locked = False
HelpClick.Visible = True

cmdBack.Locked = True
cmdBack.Visible = False
Help_Desk.Locked = True
Help_Desk.Visible = False
Tutorial.Locked = True
Tutorial.Visible = False
StyleGuide.Locked = True
StyleGuide.Visible = False
TermDB.Locked = True
TermDB.Visible = False
SwitchSC.Locked = True
SwitchSC.Visible = False
SwitchSC_AV.Locked = True
SwitchSC_AV.Visible = False
cmdConvertStd.Locked = True
cmdConvertStd.Visible = False
cmdConvertAV.Locked = True
cmdConvertAV.Visible = False


End Sub

Private Sub cmdContinueFour_Click()

' Final tutorial screen

cmdContinueFour.Locked = True
cmdContinueFour.Visible = False

QA_Tag_Count.Locked = True
QA_Tag_Count.Visible = False

Tools.Locked = True
Tools.Visible = True
Links.Locked = True
Links.Visible = True

TagCount_explain.Visible = False

Links_explain.Font.Bold = True
Links_explain.Visible = True


cmdBack.BackColor = RGB(7, 179, 64)

End Sub

Private Sub cmdContinueOne_Click()

' Second Tutorial Screen

cmdContinueOne.Locked = True
cmdContinueOne.Visible = False

Pre_Grade.Locked = True
Pre_Grade.Visible = False

Post_Grade.Locked = True
Post_Grade.Visible = True

PreGrade_explain.Visible = False

PostGrade_explain.Font.Bold = True
PostGrade_explain.Visible = True

cmdContinueTwo.Locked = False
cmdContinueTwo.BackColor = RGB(7, 179, 64)
cmdContinueTwo.Visible = True

End Sub

Private Sub cmdContinueThree_Click()

' Fourth Tutorial Screen

cmdContinueThree.Locked = True
cmdContinueThree.Visible = False

Client_Ready.Locked = True
Client_Ready.Visible = False
QA_Tag_Count.Locked = True
QA_Tag_Count.Visible = True

ClientReady_explain.Visible = False

TagCount_explain.Font.Bold = True
TagCount_explain.Visible = True

cmdContinueFour.Locked = False
cmdContinueFour.BackColor = RGB(7, 179, 64)
cmdContinueFour.Visible = True

End Sub

Private Sub cmdContinueTwo_Click()

' Third Tutorial Screen

cmdContinueTwo.Locked = True
cmdContinueTwo.Visible = False

Post_Grade.Locked = True
Post_Grade.Visible = False

Client_Ready.Locked = True
Client_Ready.Visible = True

PostGrade_explain.Visible = False

ClientReady_explain.Font.Bold = True
ClientReady_explain.Visible = True

cmdContinueThree.Locked = False
cmdContinueThree.BackColor = RGB(7, 179, 64)
cmdContinueThree.Visible = True

End Sub

Private Sub cmdConvertAV_Click()

'
' Convert Standard formatting to AV
'
'

Dim ExpertName, IntName, RepWords As String
Dim ParaCount, i, n As Integer
Dim CRange, NRange As Range
Dim CheckNext As Boolean

i = 1
ParaCount = ActiveDocument.Paragraphs.Count
    
    'Sets the whole document to black
    ActiveDocument.Content.Select
    Selection.Font.Color = 3808512

Call RemoveHeaders

    Do While i <= ParaCount
    CheckNext = True
    n = 1
            Set CRange = ActiveDocument.Paragraphs(i).Range ' Current paragaph
            CRange.Find.ClearFormatting
            CRange.Find.Replacement.ClearFormatting
            
                With CRange.Find
                .Text = "Expert:" ' Searches for Expert ID
                
                    If .Execute = True Then 'If found, changes color to desired color
                        CRange.Copy
                        ActiveDocument.Paragraphs(i).Range.Select
                        Selection.Font.Color = 78077
                                                
                        Do While CheckNext = True And (i + n) <= ParaCount 'Checks the next paragraph for Interviewer tag. If doesn't detect
                        Set NRange = ActiveDocument.Paragraphs(i + n).Range ' changes the next paragraph to be Expert tag color. Repeats this until
                                                                            ' Interviewer tag is detected.
                            With NRange.Find
                            .Text = "Interviewer:"
                                                    
                                If .Execute = False Then
                                    NRange.Copy
                                    ActiveDocument.Paragraphs(i + n).Range.Select
                                    Selection.Font.Color = 78077
                                    n = n + 1
                                Else
                                    CheckNext = False
                                End If
                                
                            End With
                            
                        Loop
                            
                    End If
                End With
        
        i = i + 1
    Loop

Selection.HomeKey Unit:=wdStory

End Sub

Private Sub cmdConvertStd_Click()
'
' Convert AV formatting to Standard
'
'

Dim ExpertName, IntName, RepWords As String
Dim ParaCount, i, n As Integer
Dim CRange, NRange As Range
Dim CheckNext As Boolean

i = 1
ParaCount = ActiveDocument.Paragraphs.Count
    
    'Sets the whole document to black
    ActiveDocument.Content.Select
    Selection.Font.Color = -587137025


    Do While i <= ParaCount
    CheckNext = True
    n = 1
            Set CRange = ActiveDocument.Paragraphs(i).Range ' Current paragaph
            CRange.Find.ClearFormatting
            CRange.Find.Replacement.ClearFormatting
            
                With CRange.Find
                .Text = "Expert:" ' Searches for Expert ID
                
                    If .Execute = True Then 'If found, changes color to desired color
                        CRange.Copy
                        ActiveDocument.Paragraphs(i).Range.Select
                        Selection.Font.Color = 2893715
                                                
                        Do While CheckNext = True And (i + n) <= ParaCount 'Checks the next paragraph for Interviewer tag. If doesn't detect
                        Set NRange = ActiveDocument.Paragraphs(i + n).Range ' changes the next paragraph to be Expert tag color. Repeats this until
                                                                            ' Interviewer tag is detected.
                            With NRange.Find
                            .Text = "Interviewer:"
                                                    
                                If .Execute = False Then
                                    NRange.Copy
                                    ActiveDocument.Paragraphs(i + n).Range.Select
                                    Selection.Font.Color = 2893715
                                    n = n + 1
                                Else
                                    CheckNext = False
                                End If
                                
                            End With
                            
                        Loop
                            
                    End If
                End With
        
        i = i + 1
    Loop

Call RemoveHeaders
Call HeaderInsert

Selection.HomeKey Unit:=wdStory

End Sub

Private Sub Help_Desk_Click()

    ' Opens Help Desk Ticket submission form
    ActiveDocument.FollowHyperlink ("https://docs.google.com/forms/d/e/1FAIpQLSc4B3uXu0hYmnlpnR4UBh64W0TAlvkCbwNrBLZqDNW1Wy1dHA/viewform?usp=sf_link")

End Sub

Private Sub HelpClick_Click()

' Resets Help Menu

PreGrade_explain.Visible = False

Post_Grade.Locked = True
Post_Grade.Visible = False
Pre_Grade.Locked = True
Pre_Grade.Visible = False
Client_Ready.Locked = True
Client_Ready.Visible = False
QA_Tag_Count.Locked = True
QA_Tag_Count.Visible = False
StyleGuide.Locked = True
StyleGuide.Visible = False
TermDB.Locked = True
TermDB.Visible = False
HelpClick.Locked = True
HelpClick.Visible = False
Tools.Locked = True
Tools.Visible = False
Links.Locked = True
Links.Visible = False

cmdBack.Locked = False
cmdBack.Visible = True
Help_Desk.Locked = False
Help_Desk.Visible = True
Tutorial.Locked = False
Tutorial.Visible = True

End Sub

Private Sub Links_Click()

' Resets Link Menu

Post_Grade.Locked = True
Post_Grade.Visible = False
Pre_Grade.Locked = True
Pre_Grade.Visible = False
Client_Ready.Locked = True
Client_Ready.Visible = False
QA_Tag_Count.Locked = True
QA_Tag_Count.Visible = False

HelpClick.Locked = True
HelpClick.Visible = False
Tools.Locked = True
Tools.Visible = False
Links.Locked = True
Links.Visible = False

StyleGuide.Locked = False
StyleGuide.Visible = True
TermDB.Locked = False
TermDB.Visible = True

cmdBack.Locked = False
cmdBack.Visible = True

End Sub

Private Sub Post_Grade_Click()

On Error GoTo ErrorHandler

'------------------------------------------------------------------------------------------------------------------
' This sub runs a series of functions and procedures designed to grade the file and display the grade to the user.

    ' Macs have an issue with Removing Personal Information from a document. To work around, all mac
    ' users have their Username saved and then recovered at the end while they use a temp username to edit
    ' the document.
    
    #If Mac Then
    
        ' STORES THE Iniital User Name and Initals
          For Each aVar In ActiveDocument.Variables
              If aVar.Name = "InitialUN" Then mVarCntSix = aVar.Index
          Next aVar
              If mVarCntSix = 0 Then
               ActiveDocument.Variables.Add Name:="InitialUN", Value:=Application.UserName
               ActiveDocument.Variables.Add Name:="InitialUI", Value:=Application.UserInitials
              End If
    
    Application.UserName = "Author" ' Temp user name
    Application.UserInitials = ""
    
    #Else
        ' Windows Users: Removes personal information
        
        ActiveDocument.RemoveDocumentInformation (wdRDIRemovePersonalInformation)
        ActiveDocument.Save
        
    #End If

'-------------------------------------------------------
'This is to check to see if Pre Grade was ever ran

        For Each aVar In ActiveDocument.Variables           ' For each variable in the document
            If aVar.Name = "InitialCount" Then mVarCntFour = aVar.Index  ' If the variable name is this then store this variable as the index number
        Next aVar
            
            If mVarCntFour = 0 Then      ' if the variable = 0, meaning it doesn't exist
             MsgBox ("WARNING: Pre Grade was never used." & vbCr & vbCr & "Always use Pre Grade for a more accurate result.")
             ActiveDocument.Variables.Add Name:="InitialCount", Value:=0         'create dummy value
             ActiveDocument.Variables.Add Name:="InitialUserScore", Value:=100         'create dummy value
            End If
'----------------------------------------------------------

'This is to store the version number

        For Each aVar In ActiveDocument.Variables
            If aVar.Name = "InHouseReport" Then mVarCntSeven = aVar.Index
        Next aVar
            If mVarCntSeven = 0 Then
             ActiveDocument.Variables.Add Name:="InHouseReport", Value:=mstrIHRep
            Else
             ActiveDocument.Variables(mVarCntSeven).Value = mstrIHRep
            End If
'----------------------------------------------------------
    

    Dim MWClearedPoints, FinalPreUserScore, QACommentScore As Variant
    
    'For Missing Words Cleared
    MWClearedPoints = 0
    
    'For Final Points
    FinalPreUserScore = ActiveDocument.Variables("InitialUserScore").Value
    
    'For QA Comment grades
    Dim WrongOrder, CapErrors, SGErrors, PunctErrors, MissingWord, MissingSent, MissingUnTwo, MissingOvTwo, _
        MissingOvFive, MissingNT, HeaderWrong, TagWrong, TemplateWrong, Consistency, WordsNotSpoken, AVWrong, _
        Mishears, Typo, SpeakerColor, SpeakerID, IUS, IC, ICount, GCount As Integer

    Dim GradeComments As String
    
    Dim SaveDefault As Boolean
    
    Dim AcceptDoc, CurDoc As Document
    
    Dim MyData As DataObject

     '--------- HOW TO CHECK AGAINST THE UNACCEPTED TRACKCHANGES
        
        Set CurDoc = ActiveDocument
        
        Set MyData = New DataObject             ' Clears the clipboard
        MyData.SetText ""
        MyData.PutInClipboard
        
            CurDoc.Content.Copy         ' Copys content
            
            Set AcceptDoc = Documents.Add           'CREATES NEW DOCUMENT
                Documents(AcceptDoc).Activate
        
                AcceptDoc.Content.Paste             ' PASTES INTO NEW DOCUMENT
        
                AcceptDoc.AcceptAllRevisions        ' ACCEPTS CHANGES
                AcceptDoc.Revisions.AcceptAll
                    
            ICount = CountThis("[inaudible") 'Counts listed tags
        
            GCount = GCount + CountThis("? 0")
            GCount = GCount + CountThis("? 1")
            GCount = GCount + CountThis("? 2")
            GCount = GCount + CountThis("? 3")
            GCount = GCount + CountThis("? 4")
            GCount = GCount + CountThis("? 5")
  '--------------------------------------------------------------------------------------------------------
  'This is to create a document variable that stores the final amount of missing words

            For Each aVar In AcceptDoc.Variables
                If aVar.Name = "FinalCount" Then mVarCntThree = aVar.Index
                
            Next aVar
                If mVarCntThree = 0 Then
                 AcceptDoc.Variables.Add Name:="FinalCount", Value:=GCount + ICount
                Else
                 AcceptDoc.Variables(mVarCntThree).Value = GCount + ICount
                End If
            
            'Stores the final amount of missing words
            mFinalMWCount = AcceptDoc.Variables("FinalCount").Value
            
'------------------------------------------------------------------------------------------------------------

            AcceptDoc.Close _
            SaveChanges:=wdDoNotSaveChanges             ' closes new document without saving
              
            Documents(CurDoc).Activate
        
        ' Calculates the difference in tags
       mFinalMWCount = CurDoc.Variables("InitialCount").Value - ICount - GCount
       
       
       If mFinalMWCount < 0 Then    ' To account for instances in which the QA adds more MW than the TR had.
          mFinalMWCount = 0
       End If
       
'---------------------------------------------------------------------------------------------------------
       'GRADING FOR MISSING WORDS CLEARED
'---------------------------------------------------------------------------------------------------------
       
       If mFinalMWCount >= 3 And mFinalMWCount <= 6 Then
          MWClearedPoints = (mFinalMWCount * 0.25) - 0.75
       ElseIf mFinalMWCount >= 7 Then
          MWClearedPoints = (mFinalMWCount * 0.5) - 2.25
       End If
       
'---------------------------------------------------------------------------------------------------------
       'CALCULATE TOTAL AMOUNT OF QA ENTERED COMMENTS
'---------------------------------------------------------------------------------------------------------
            
    ' Searches all comments in the document for specified word. If word appears,
    ' the number of errors associated with the word increases
        
        Typo = CheckComments("Typo")
        SpeakerID = CheckComments("ID")
        SpeakerColor = CheckComments("color")
        SGErrors = CheckComments("Style")
        
        MissingWord = CheckComments("word(s)")
        MissingSent = CheckComments("sentence(s)")
        MissingUnTwo = CheckComments("Under 2 minutes")
        MissingOvTwo = CheckComments("Over 2 minutes")
        MissingOvFive = CheckComments("Over 5 minutes")
        MissingNT = CheckComments("Nothing")
        
        WordsNotSpoken = CheckComments("Adding")
        Consistency = CheckComments("Consistency")
        HeaderWrong = CheckComments("Header")
        TagWrong = CheckComments("Tag")
        TemplateWrong = CheckComments("template")
        AVWrong = CheckComments("AlphaView")
        
        CapErrors = CheckComments("Cap")
        PunctErrors = CheckComments("Punc")
        Mishears = CheckComments("Mish")
                
        ' Caps Consistency errors and Wrong Tag errors at 5 each
                
        If Consistency > 5 Then
            Consistency = 5
        End If
        
        If TagWrong > 5 Then
            TagWrong = 5
        End If
        
'------------------------------------------SCORING------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------
        
       'Calculates user score based on scoring structure
        
        If MissingNT > 0 Then
             QACommentScore = 100
        Else
             QACommentScore = (CapErrors * 0.1) + (SGErrors * 0.1) + (PunctErrors * 0.1) _
             + (WordsNotSpoken * 0.25) + (Consistency * 0.25) + (HeaderWrong * 0.5) + (TagWrong * 0.25) _
             + (TemplateWrong * 0.5) + (AVWrong * 0.5) + (Mishears * 0.5) + (Typo * 0.5) + SpeakerID + SpeakerColor _
             + (MissingWord * 0.5) + MissingSent + (MissingUnTwo * 2) + (MissingOvTwo * 5) + (MissingOvFive * 15)
        End If
        
        
        IUS = ActiveDocument.Variables("InitialUserScore").Value
        IC = ActiveDocument.Variables("InitialCount").Value
       
        
        If QACommentScore < 0 Then ' Accounts for calculation error
            QACommentScore = 0
        End If
        
        'Calculates the final score based on QA Comments, Missing Words Cleared, and Pre Grade score (Initial User Score).
        FinalPreUserScore = IUS - MWClearedPoints - QACommentScore
        
        If FinalPreUserScore < 0 Then
            FinalPreUserScore = 0
        End If
          
'------------------------------------------------------------------------------------------------------------------------------------
'                                  DISPLAYS FINAL SCORE
'------------------------------------------------------------------------------------------------------------------------------------

        'If the QA marks that the TR didn't transcribe anything
        
        If MissingNT > 0 Then
                
                Response = MsgBox("|Final Grade|" & vbCr & vbCr & _
                          "Final User Score: " & FinalPreUserScore & vbCr & vbCr & _
                          "Nothing was transcribed." _
                            , vbOKOnly + vbInformation, "Post Grade Assessment")
                Exit Sub
        End If
                
                Response = MsgBox("|Final Grade|" & vbCr & vbCr & _
                          "Final User Score: " & FinalPreUserScore & vbCr & vbCr & _
                          "Pre Grade Score: " & IUS & vbCr & vbCr & _
                            mFinalMWCount & " cleared missing words resulted in a loss of " & _
                            MWClearedPoints & " points." & vbCr & vbCr & "Lost points due to edits: " & _
                            QACommentScore & " points." _
                            , vbOKOnly + vbInformation, "Post Grade Assessment")
                            
        
        Set MyData = New DataObject ' Clear the clipboard
        MyData.SetText ""
        MyData.PutInClipboard
        
        Exit Sub
        
        
        
ErrorHandler:

Select Case Err.Number

    Case 4605
    AcceptDoc.Close _
    SaveChanges:=wdDoNotSaveChanges
        
    Set AcceptDoc = Documents.Add
    Documents(AcceptDoc).Activate
    

    Case 4198
    AcceptDoc.Close _
    SaveChanges:=wdDoNotSaveChanges
    
    CurDoc.Content.Copy
    
    Set AcceptDoc = Documents.Add
    Documents(AcceptDoc).Activate
    
    AcceptDoc.Content.Paste
    
End Select

Resume

End Sub

Private Sub Pre_Grade_Click()
    
    #If Mac Then
    
        ' STORES THE Initial User Name and Initials
          For Each aVar In ActiveDocument.Variables
              If aVar.Name = "InitialUN" Then mVarCntSix = aVar.Index
          Next aVar
              If mVarCntSix = 0 Then
               ActiveDocument.Variables.Add Name:="InitialUN", Value:=Application.UserName
               ActiveDocument.Variables.Add Name:="InitialUI", Value:=Application.UserInitials
              End If
    
    Application.UserName = "Author"
    Application.UserInitials = ""
    
    #Else
        
        ActiveDocument.RemoveDocumentInformation (wdRDIRemovePersonalInformation)
        ActiveDocument.Save
        
    #End If
    
    ' STORES THE SEED
              For Each aVar In ActiveDocument.Variables
                  If aVar.Name = "InHouseReport" Then mVarCntSeven = aVar.Index
              Next aVar
                  If mVarCntSeven = 0 Then
                   ActiveDocument.Variables.Add Name:="InHouseReport", Value:=mstrIHRep
                  Else
                   ActiveDocument.Variables(mVarCntSeven).Value = mstrIHRep
                  End If


ActiveDocument.TrackRevisions = False

Call NoBorders
Call OneSpace
Call ReplaceQuotes

    Selection.WholeStory
    Selection.Font.Size = 11
    Selection.Font.Name = "Arial"


ActiveDocument.TrackRevisions = True

    'For TGCount
    Dim TGCount, ICount, GCount As Integer
               
    'For Timestamps
    Dim TimeStampPoints As Variant
    
    'For Punctuation
    Dim PunctPoints As Variant
    
    'For Function Calls
    Dim TSCount As Integer
    
    'For WER
    Dim WERCountPoints As Variant
    
    Dim PreUserScore As Variant

    PreUserScore = 100
    
    TSCount = 0
    mPCount = 0
    TGCount = 0
    ICount = 0
    GCount = 0
    mWERcount = 0
    
    
    mRWCount = mRWCount + ReplaceWords(Chr(150), Chr(45))    ' en dash
    mRWCount = mRWCount + ReplaceWords(Chr(151), Chr(45))    ' em dash

    mRWCount = mRWCount + ReplaceWords("1 percent ", "1%")
    mRWCount = mRWCount + ReplaceWords("2 percent ", "2%")
    mRWCount = mRWCount + ReplaceWords("3 percent ", "3%")
    mRWCount = mRWCount + ReplaceWords("4 percent ", "4%")
    mRWCount = mRWCount + ReplaceWords("5 percent ", "5%")
    mRWCount = mRWCount + ReplaceWords("6 percent ", "6%")
    mRWCount = mRWCount + ReplaceWords("7 percent ", "7%")
    mRWCount = mRWCount + ReplaceWords("8 percent ", "8%")
    mRWCount = mRWCount + ReplaceWords("9 percent ", "9%")
    
    mRWCount = mRWCount + ReplaceWords("alright", "all right")
    mRWCount = mRWCount + ReplaceWords("Alright", "All right")
    mRWCount = mRWCount + ReplaceWords("appdev", "AppDev")
    mRWCount = mRWCount + ReplaceWords("b2b", "B2B")
    mRWCount = mRWCount + ReplaceWords("b2b2c", "B2B2C")
    mRWCount = mRWCount + ReplaceWords("b2c", "B2C")
    mRWCount = mRWCount + ReplaceWords("Brownfield", "brownfield")
    mRWCount = mRWCount + ReplaceWords("c2b", "C2B")
    mRWCount = mRWCount + ReplaceWords("c2c", "C2C")
    mRWCount = mRWCount + ReplaceWords("CAAS", "CaaS")
    mRWCount = mRWCount + ReplaceWords("capex", "CapEx")
    mRWCount = mRWCount + ReplaceWords("CAPEX", "CapEx")
    mRWCount = mRWCount + ReplaceWords("codev", "co-dev")
    mRWCount = mRWCount + ReplaceWords("cross talk", "crosstalk")
    mRWCount = mRWCount + ReplaceWords("crossstalk", "crosstalk")
    mRWCount = mRWCount + ReplaceWords("devops", "DevOps")
    mRWCount = mRWCount + ReplaceWords("DRAAS", "DRaaS")
    mRWCount = mRWCount + ReplaceWords("ecommerce", "e-commerce")
    mRWCount = mRWCount + ReplaceWords("eCommerce", "e-commerce")
    mRWCount = mRWCount + ReplaceWords("e-signature", "e-signature")
    mRWCount = mRWCount + ReplaceWords("eSignature", "e-signature")
    mRWCount = mRWCount + ReplaceWords("et cetera", "etc.")
    mRWCount = mRWCount + ReplaceWords("FinTech", "fintech")
    mRWCount = mRWCount + ReplaceWords("gonna", "going to")
    mRWCount = mRWCount + ReplaceWords("Greenfield", "greenfield")
    mRWCount = mRWCount + ReplaceWords("HIPPA", "HIPAA")
    mRWCount = mRWCount + ReplaceWords("InsurTech", "insurtech")
    mRWCount = mRWCount + ReplaceWords("IOT", "IoT")
    mRWCount = mRWCount + ReplaceWords("Ok", "Okay")
    mRWCount = mRWCount + ReplaceWords("opex", "OpEx")
    mRWCount = mRWCount + ReplaceWords("OPEX", "OpEx")
    mRWCount = mRWCount + ReplaceWords("owner/operator", "owner-operator")
    mRWCount = mRWCount + ReplaceWords("PAAS", "PaaS")
    mRWCount = mRWCount + ReplaceWords("payor", "payer")
    mRWCount = mRWCount + ReplaceWords("payors", "payers")
    mRWCount = mRWCount + ReplaceWords("SAAS", "SaaS")
    mRWCount = mRWCount + ReplaceWords("smart watch", "smartwatch")
    mRWCount = mRWCount + ReplaceWords("software as a service", "Software-as-a-Service")
    mRWCount = mRWCount + ReplaceWords("software-as-a-service", "Software-as-a-Service")
    mRWCount = mRWCount + ReplaceWords("Software-as-a-service", "Software-as-a-Service")
    mRWCount = mRWCount + ReplaceWords("three-D", "3D")
    mRWCount = mRWCount + ReplaceWords("two-D", "2D")
    mRWCount = mRWCount + ReplaceWords("wanna", "want to")
    mRWCount = mRWCount + ReplaceWords("zip code", "ZIP Code")
    mRWCount = mRWCount + ReplaceWords("Zip Code", "ZIP Code")
    mRWCount = mRWCount + ReplaceWords("Zip code", "ZIP Code")
    mRWCount = mRWCount + ReplaceWords("Zip", "ZIP")
                                
                If mWERcount < 1 Then
                    mWERcount = 0
                End If
                
                WERCountPoints = mWERcount * 0.5 ' Stores point deduction for WER
      
    mPCount = mPCount + CheckPunctuation(Chr(133), "Punctuation: Ellipses")
    mPCount = mPCount + CheckPunctuation("...", "Punctuation: Ellipses")
    mPCount = mPCount + CheckPunctuation("; ", "Punctuation: Semicolon")
    mPCount = mPCount + CheckPunctuation("--", "Punctuation: Double Dash")
    mPCount = mPCount + CheckPunctuation("~", "Punctuation: Tilde")
    mPCount = mPCount + CheckPunctuation("<", "Punctuation: Wrong Bracket")
    mPCount = mPCount + CheckPunctuation(">", "Punctuation: Wrong Bracket")
    mPCount = mPCount + CheckPunctuation("{", "Punctuation: Wrong Bracket")
    mPCount = mPCount + CheckPunctuation("}", "Punctuation: Wrong Bracket")
       
                PunctPoints = mPCount * 0.5 ' Stores point deduction for wrong punctuation
    
    ICount = CountThis("[inaudible") 'Counts listed tags
    GCount = GCount + CountThis("? 0")
    GCount = GCount + CountThis("? 1")
    GCount = GCount + CountThis("? 2")
    GCount = GCount + CountThis("? 3")
    GCount = GCount + CountThis("? 4")
    GCount = GCount + CountThis("? 5")
    
                TGCount = ICount + GCount

    TSCount = TSCount + TimeStampCheck("[0-9][0-9]:[0-9][0-9]:")
    TSCount = TSCount + TimeStampCheck(":[0-9][0-9].[0-9]")
    TSCount = TSCount + TimeStampCheck(":[0-9][0-9][0-9]")
                 
                TimeStampPoints = TSCount * 0.5 ' Stores point deduction for wrong timestamp

                    
          PreUserScore = PreUserScore - TimeStampPoints - PunctPoints - WERCountPoints
          
        If PreUserScore < 99 Then
            PreUserScore = 99
        End If
          
          ' STORES THE NUMBER OF INITIAL MISSING WORDS
                    For Each aVar In ActiveDocument.Variables
                        If aVar.Name = "InitialCount" Then mVarCntOne = aVar.Index
                    Next aVar
                        If mVarCntOne = 0 Then
                         ActiveDocument.Variables.Add Name:="InitialCount", Value:=TGCount
                        Else
                         ActiveDocument.Variables(mVarCntOne).Value = TGCount
                        End If
      
          ' STORES THE INITIAL USER SCORE
                    For Each aVar In ActiveDocument.Variables
                        If aVar.Name = "InitialUserScore" Then mVarCntTwo = aVar.Index
                    Next aVar
                        If mVarCntTwo = 0 Then
                         ActiveDocument.Variables.Add Name:="InitialUserScore", Value:=PreUserScore
                        Else
                         ActiveDocument.Variables(mVarCntTwo).Value = PreUserScore
                        End If
                        
          ' STORES THE INITIAL REPLACED WORDS
                    For Each aVar In ActiveDocument.Variables
                        If aVar.Name = "InitialReplacedWords" Then mVarCntFive = aVar.Index
                    Next aVar
                        If mVarCntFive = 0 Then
                         ActiveDocument.Variables.Add Name:="InitialReplacedWords", Value:=mRWCount
                        Else
                         ActiveDocument.Variables(mVarCntFive).Value = mRWCount
                        End If
                    
                        
    ' DISPLAY FINAL PRE GRADE SCORE FOR EVERYTHING
    
    Response = MsgBox("Pre Grade Score: " & PreUserScore & vbCr & vbCr & vbCr & _
                        TSCount & " timestamp error(s)." & vbCr & _
                        mPCount & " punctuation error(s)." & vbCr & _
                        mWERcount & " spelling/formatting error(s)." & vbCr & _
                        vbCr & mRWCount, vbOKOnly + vbInformation, "Pre Grade Assessment")
                        
    Unload Me
    
End Sub

Private Function ReplaceWords(TestWord, RepWord As String) As String

        Set mCurrentRng = ActiveDocument.Content
        
        mCurrentRng.Find.ClearFormatting

        ' This function will replace the first word with the second word
        
        With mCurrentRng.Find                        ' Replaces test word
            .Text = TestWord
            .Forward = True
            .MatchWildcards = False
            .MatchCase = True
            .MatchWholeWord = True
            .Wrap = 0
            
                Do While .Execute           ' counts the instances
                    ReplaceWords = TestWord + " with " + RepWord + vbCr
                    mCurrentRng.Find.Replacement.Text = RepWord
                    mWERcount = mWERcount + 1
                Loop
        
        End With
             ' replaces the word
        Set mCurrentRng = ActiveDocument.Content
        With mCurrentRng.Find
            .MatchWholeWord = True
            .MatchCase = True
            .Execute Findtext:=TestWord, replacewith:=RepWord, _
            Replace:=wdReplaceAll
        End With
        
End Function
    
Private Sub Client_Ready_Click()
mResearchCount = 0

On Error GoTo ErrorHandler


    ActiveDocument.TrackRevisions = False       '   Turns off track changes
    ActiveDocument.AcceptAllRevisions           '   Accepts all revisions
    
    Dim comrange As Range
    Dim TSCount As Integer
    
    
    ' STORES THE SEED
              For Each aVar In ActiveDocument.Variables
                  If aVar.Name = "InHouseReport" Then mVarCntSeven = aVar.Index
              Next aVar
                  If mVarCntSeven = 0 Then
                   ActiveDocument.Variables.Add Name:="InHouseReport", Value:=mstrIHRep
                  Else
                   ActiveDocument.Variables(mVarCntSeven).Value = mstrIHRep
                  End If
    
    
    TSCount = 0
        
    Set comrange = ActiveDocument.Content
        
    If comrange.Application.ActiveDocument.Comments.Count <> 0 Then     ' Deletes all comments
        comrange.Application.ActiveDocument.DeleteAllComments
    End If
    
    Call ReplaceQuotes
    Call NoBorders
    Call OneSpace
    
    TSCount = TSCount + TimeStampCheck("[0-9][0-9]:[0-9][0-9]:") 'Checks for QA entered incorrect timestamps
    TSCount = TSCount + TimeStampCheck(":[0-9][0-9].[0-9]")
    TSCount = TSCount + TimeStampCheck(":[0-9][0-9][0-9]")
    
    If TSCount > 0 Then         'Alerts the QA that there are incorrectly formatted timestamps
        MsgBox ("Please check timestamps for formatting accuracy.")
        Exit Sub
    End If
        
    
    ActiveDocument.CheckSpelling        'Spellcheck
    
    
     
    mTagDisp = False    ' Turns off display for Tag Count
    
        Call QA_Tag_Count_Click ' Captures count for research tags
    
    mTagDisp = True     ' Turns display back on for Tag Count
     
    If mResearchCount <> 0 Then
        MsgBox ("WARNING: There are still research tags in this document." & vbCr & _
                vbCr & "Please delete before continuing.")
    Else
        MsgBox ("This document is client ready.")
                      

               
        #If Mac Then
        
            Application.UserName = ActiveDocument.Variables("InitialUN").Value
            Application.UserInitials = ActiveDocument.Variables("InitialUI").Value
        
        #Else
        
            ActiveDocument.RemoveDocumentInformation (wdRDIRemovePersonalInformation)
            ActiveDocument.Save
            
        #End If
               

    End If


Unload Me
Exit Sub


ErrorHandler:

Select Case Err.Number

    Case 5825

End Select
Resume Next

End Sub

Private Sub QA_Leave_Click()

    Unload Me
    
End Sub

Private Sub QA_Tag_Count_Click()

    Dim ICount, GCount, CCount, FCount, SCount, DCount, RCount As Integer
    Dim InDisp, GuessDisp, CrossDisp As String

    ICount = CountThis("[inaudible") 'Counts listed tags
    GCount = GCount + CountThis("? 0")
    GCount = GCount + CountThis("? 1")
    GCount = GCount + CountThis("? 2")
    GCount = GCount + CountThis("? 3")
    GCount = GCount + CountThis("? 4")
    GCount = GCount + CountThis("? 5")
    CCount = CountThis("[crosstalk")
    FCount = CountThis("[foreign")
    SCount = CountThis("[silence")
    DCount = CountThis("[call dropped")
    
    'Special loop used to search for research tags
    Set mCurrentRng = ActiveDocument.Content
    With mCurrentRng.Find
        .MatchWildcards = True
        .Text = "\[*\]"
        .Format = False
        .Wrap = 0
        .Forward = False
        
            Do While .Execute
                RCount = RCount + 1
            Loop
    End With
    
    'STORES THE NUMBER OF RESEARCH TAGS
    
    RCount = RCount - FCount - CCount - GCount - ICount - SCount - DCount
    If RCount < 0 Then
        RCount = 0
    End If

    mResearchCount = RCount
   
    
' CHANGES GRAMMAR AS NEEDED
    
    If ICount <> 1 Then
        InDisp = "inaudibles"
    Else
        InDisp = "inaudible"
    End If
        
    If GCount <> 1 Then
        GuessDisp = "guesses"
    Else
        GuessDisp = "guess"
    End If
    
    If CCount <> 1 And CCount <> 0 Then
        CrossDisp = "crosstalks"
    Else
        CrossDisp = "crosstalk"
    End If
    
'DISPLAYS MESSAGE
    If mTagDisp = True Then
    
            Response = MsgBox("This document contains" & vbCr & "the following tags:" & vbCr & vbCr & _
                    ICount & " " & InDisp & vbCr _
                    & GCount & " " & GuessDisp & vbCr _
                    & CCount & " " & CrossDisp & _
                    vbCr & FCount & " foreign" & vbCr & _
                    RCount & " research", vbOKOnly + vbInformation, "Tag Count")
    End If
    

End Sub

Private Function CheckComments(QAComment As String) As Integer

Dim n As Integer
Dim CommentsRange As Range

n = 0
i = 1

    Set CommentsRange = ActiveDocument.Content
        CommentsRange.Find.ClearFormatting
        
            If CommentsRange.Application.ActiveDocument.Comments.Count <> 0 Then    'Checks to see if there are any comments
                For i = 1 To CommentsRange.Application.ActiveDocument.Comments.Count        ' Goes through every comment
                                                  
                      With ActiveDocument.Comments(i).Range.Find        ' Seaches the comment for the specified word
                          .MatchWholeWord = False
                          .MatchCase = True
                          .Text = QAComment
                          .Format = False
                          .Wrap = 0
                          .Forward = False
                              If .Execute = True Then                   ' If word is found, adds to total number of found instances
                                   n = n + 1
                              End If
                      End With
                      
                Next i
            End If
 
    CheckComments = n

End Function

Private Function CheckPunctuation(PunctSymbol, PunctDisp As String) As Integer
i = 0
                'Looks for incorrect Punctuation
                    
                    Set mCurrentRng = ActiveDocument.Content
                    With mCurrentRng.Find
                        .MatchWildcards = False
                        .Text = PunctSymbol
                        .Format = False
                        .Wrap = 0
                        .Forward = False
                        
                            Do While .Execute
                                ActiveDocument.Comments.Add _
                                Range:=mCurrentRng, Text:=PunctDisp
                                i = i + 1
                            Loop
                    End With
                    
                    CheckPunctuation = i

End Function

Private Function TimeStampCheck(WrongTime As String) As Integer
i = 0
                    ' Run a loop that will count incorrect time stamps
                    
                    Set mCurrentRng = ActiveDocument.Content
                    With mCurrentRng.Find
                        .MatchWildcards = True
                        .Text = WrongTime
                        .Format = False
                        .Wrap = 0
                        .Forward = False

                            Do While .Execute
                                ActiveDocument.Comments.Add _
                                Range:=mCurrentRng, Text:="Style Guide: Timestamps should be in the following format: H:MM:SS."
                                i = i + 1
                            Loop
                    End With
                    
                    TimeStampCheck = i
                    
End Function

Private Function CountThis(CntThis As String) As Integer
i = 0
                    ' Run a loop that will count whatever you want
                    
                    Set mCurrentRng = ActiveDocument.Content
                    With mCurrentRng.Find
                        .ClearFormatting
                        .MatchWildcards = False
                        .Text = CntThis
                        .Wrap = 0
                        .Forward = False

                            Do While .Execute
                                i = i + 1
                            Loop
                            
                    End With
                    
                    CountThis = i
                    
End Function

Private Sub ReplaceQuotes()

    ' This is to account for graves and acute accents for those who use US International keyboards
        
        ActiveDocument.Content.Find.Execute Findtext:=Chr(96), MatchCase:=False, _
        MatchWholeWord:=False, MatchWildcards:=False, Wrap:=wdFindStop, _
        Format:=False, replacewith:=Chr(39), Replace:=wdReplaceAll
        
        ActiveDocument.Content.Find.Execute Findtext:=Chr(180), MatchCase:=False, _
        MatchWholeWord:=False, MatchWildcards:=False, Wrap:=wdFindStop, _
        Format:=False, replacewith:=Chr(39), Replace:=wdReplaceAll

    ' This macro is designed to replace all straight quotes with curly quotes without changing the users predefined settings.
    
    Dim blnQuotes As Boolean

        blnQuotes = Application.Options.AutoFormatAsYouTypeReplaceQuotes
    
    If Application.Options.AutoFormatAsYouTypeReplaceQuotes = True Then ' If the user already has curly quotes selected
    
        ActiveDocument.Content.Find.Execute Findtext:="'", MatchCase:=False, _
        MatchWholeWord:=False, MatchWildcards:=False, Wrap:=wdFindStop, _
        Format:=False, replacewith:="'", Replace:=wdReplaceAll
        
    ElseIf Application.Options.AutoFormatAsYouTypeReplaceQuotes = False Then ' If the user does not have curly quotes selected
    
        Application.Options.AutoFormatAsYouTypeReplaceQuotes = True
        ActiveDocument.Content.Find.Execute Findtext:="'", MatchCase:=False, _
        MatchWholeWord:=False, MatchWildcards:=False, Wrap:=wdFindStop, _
        Format:=False, replacewith:="'", Replace:=wdReplaceAll
        Application.Options.AutoFormatAsYouTypeReplaceQuotes = False ' Changes it back to the user preferred setting
        
    End If

End Sub

Private Sub NoBorders()
'
' NoBorders Macro
'
'
    Selection.WholeStory
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone

    
End Sub

Private Sub StyleGuide_Click()

    ActiveDocument.FollowHyperlink ("https://alphasights.docsend.com/view/4thtb7qpj7abpy5j/d/883mnb4cs4mui9zw")
    

End Sub

Private Sub SwitchSC_AV_Click()

'
' Switches Speaker IDs and Colors in a document
'
'

Dim ExpertName, IntName, RepWords As String
Dim ParaCount, i As Integer
Dim CRange, NRange As Range
Dim CheckNext As Boolean

ActiveDocument.TrackRevisions = False

i = 1

    ' Replaces the Speaker tags
    RepWords = ReplaceWordsColor("Expert:", "Xpert:", 78077)
    RepWords = ReplaceWordsColor("Interviewer:", "Expert:", 78077)
    RepWords = ReplaceWordsColor("Xpert:", "Interviewer:", 3808512)
    
    'Gets the paragraph count
    ParaCount = ActiveDocument.Paragraphs.Count
    
    'Sets the whole document to black
    ActiveDocument.Content.Select
    Selection.Font.Color = 3808512
    
    
    ' i is used to search through each paragraph with a speaker ID.CheckNext and n are used to account for paragraphs
    ' without a speaker id.
    
    Do While i <= ParaCount
    CheckNext = True
    n = 1
            Set CRange = ActiveDocument.Paragraphs(i).Range ' Current paragaph
            CRange.Find.ClearFormatting
            CRange.Find.Replacement.ClearFormatting
            
                With CRange.Find
                .Text = "Expert:" ' Searches for Expert ID
                
                    If .Execute = True Then 'If found, changes color to desired color
                        CRange.Copy
                        ActiveDocument.Paragraphs(i).Range.Select
                        Selection.Font.Color = 78077
                                                
                        Do While CheckNext = True And (i + n) <= ParaCount 'Checks the next paragraph for Interviewer tag. If doesn't detect
                        Set NRange = ActiveDocument.Paragraphs(i + n).Range ' changes the next paragraph to be Expert tag color. Repeats this until
                                                                            ' Interviewer tag is detected.
                            With NRange.Find
                            .Text = "Interviewer:"
                                                    
                                If .Execute = False Then
                                    NRange.Copy
                                    ActiveDocument.Paragraphs(i + n).Range.Select
                                    Selection.Font.Color = 78077
                                    n = n + 1
                                Else
                                    CheckNext = False
                                End If
                                
                            End With
                            
                        Loop
                            
                    End If
                End With
        
        i = i + 1
    Loop
    
    ActiveDocument.TrackRevisions = True

End Sub

Private Sub SwitchSC_Click()
'
' Switches Speaker IDs and Colors in a document
'
'

Dim ExpertName, IntName, RepWords As String
Dim ParaCount, i As Integer
Dim CRange, NRange As Range
Dim CheckNext As Boolean

ActiveDocument.TrackRevisions = False

i = 1

    ' Replaces the Speaker tags
    RepWords = ReplaceWordsColor("Expert:", "Xpert:", 2893715)
    RepWords = ReplaceWordsColor("Interviewer:", "Expert:", 2893715)
    RepWords = ReplaceWordsColor("Xpert:", "Interviewer:", -587137025)
    
    'Gets the paragraph count
    ParaCount = ActiveDocument.Paragraphs.Count
    
    'Sets the whole document to black
    ActiveDocument.Content.Select
    Selection.Font.Color = -587137025
    
    ' i is used to search through each paragraph with a speaker ID.CheckNext and n are used to account for paragraphs
    ' without a speaker id.
    
    Do While i <= ParaCount
    CheckNext = True
    n = 1
            Set CRange = ActiveDocument.Paragraphs(i).Range ' Current paragaph
            CRange.Find.ClearFormatting
            CRange.Find.Replacement.ClearFormatting
            
                With CRange.Find
                .Text = "Expert:" ' Searches for Expert ID
                
                    If .Execute = True Then 'If found, changes color to desired color
                        CRange.Copy
                        ActiveDocument.Paragraphs(i).Range.Select
                        Selection.Font.Color = 2893715
                                                
                        Do While CheckNext = True And (i + n) <= ParaCount 'Checks the next paragraph for Interviewer tag. If doesn't detect
                        Set NRange = ActiveDocument.Paragraphs(i + n).Range ' changes the next paragraph to be Expert tag color. Repeats this until
                                                                            ' Interviewer tag is detected.
                            With NRange.Find
                            .Text = "Interviewer:"
                                                    
                                If .Execute = False Then
                                    NRange.Copy
                                    ActiveDocument.Paragraphs(i + n).Range.Select
                                    Selection.Font.Color = 2893715
                                    n = n + 1
                                Else
                                    CheckNext = False
                                End If
                                
                            End With
                            
                        Loop
                            
                    End If
                End With
        
        i = i + 1
    Loop
    
    ActiveDocument.TrackRevisions = True
    
End Sub

Private Sub TermDB_Click()

    'Opens Terminology Database
    ActiveDocument.FollowHyperlink ("https://docs.google.com/spreadsheets/d/1npD88Ud6XB4xIplKW8_OJjLkTd0O105e3WlfMPGzt5A/edit#gid=1097598184")

End Sub

Private Sub Tools_Click()

' Resets screen for Tools Menu

Post_Grade.Locked = True
Post_Grade.Visible = False
Pre_Grade.Locked = True
Pre_Grade.Visible = False
Client_Ready.Locked = True
Client_Ready.Visible = False
QA_Tag_Count.Locked = True
QA_Tag_Count.Visible = False

HelpClick.Locked = True
HelpClick.Visible = False
Tools.Locked = True
Tools.Visible = False
Links.Locked = True
Links.Visible = False

SwitchSC.Locked = False
SwitchSC.Visible = True
SwitchSC_AV.Locked = False
SwitchSC_AV.Visible = True
cmdConvertStd.Locked = False
cmdConvertStd.Visible = True
cmdConvertAV.Locked = False
cmdConvertAV.Visible = True

cmdBack.Locked = False
cmdBack.Visible = True


End Sub

Private Sub Tutorial_Click()

'Resets screen for Tutorial stage 1

Tutorial.Locked = True
Tutorial.Visible = False
Help_Desk.Locked = True
Help_Desk.Visible = False

Pre_Grade.Locked = True
Pre_Grade.Visible = True

PreGrade_explain.Visible = True
PreGrade_explain.Font.Bold = True

cmdContinueOne.Locked = False
cmdContinueOne.BackColor = RGB(7, 179, 64)
cmdContinueOne.Visible = True


End Sub

Private Sub UserForm_Initialize()

    mTagDisp = True

'-------------------Version Number--------------------
                mstrIHRep = "1.9.4.3"
'-------------------Version Number--------------------

    #If Mac Then
    
        ResizeUserForm Me
        
    #End If


End Sub

Private Sub ResizeUserForm(frm As Object, Optional dResizeFactor As Double = 0#)
  
  ' Resizes user form for macs
  
  Dim ctrl As Control
  Dim sColWidths As String
  Dim vColWidths As Variant
  Dim iCol As Long

  If dResizeFactor = 0 Then dResizeFactor = 1.333333
  With frm
    .Height = .Height * dResizeFactor
    .Width = .Width * dResizeFactor

    For Each ctrl In frm.Controls
      With ctrl
        .Height = .Height * dResizeFactor
        .Width = .Width * dResizeFactor
        .Left = .Left * dResizeFactor
        .Top = .Top * dResizeFactor
        On Error Resume Next
        .Font.Size = .Font.Size * dResizeFactor
        On Error GoTo 0

        ' multi column listboxes, comboboxes
        Select Case TypeName(ctrl)
          Case "ListBox", "ComboBox"
            If ctrl.ColumnCount > 1 Then
              sColWidths = ctrl.ColumnWidths
              vColWidths = Split(sColWidths, ";")
              For iCol = LBound(vColWidths) To UBound(vColWidths)
                vColWidths(iCol) = Val(vColWidths(iCol)) * dResizeFactor
              Next
              sColWidths = Join(vColWidths, ";")
              ctrl.ColumnWidths = sColWidths
            End If
        End Select
      End With
    Next
  End With
End Sub

Private Sub OneSpace()

        Set mCurrentRng = ActiveDocument.Content
        mCurrentRng.Find.ClearFormatting

        ' This sub will replace two spaces with one
        
        With mCurrentRng.Find                        ' Replaces test word
            .Text = "  "
            .Forward = True
            .MatchWildcards = False
            .MatchCase = True
            .MatchWholeWord = True
            .Wrap = 0
            
                Do While .Execute
                    mCurrentRng.Find.Replacement.Text = " "
                    .Execute Replace:=wdReplaceAll
                Loop
        End With
End Sub

Private Function ReplaceWordsColor(TestWord, RepWord As String, Col As Long) As String
    
        'This sub will replace word and switch color

    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = Col
    End With
    With Selection.Find
        .Text = TestWord
        .Replacement.Text = RepWord
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
End Function

Private Sub HeaderInsert()
'
' HeaderInsert Macro
'
'

Dim strDisclaimer As String

    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    
    With ActiveDocument.Sections(1)
        .Headers(wdHeaderFooterFirstPage).Range.Text = strDisclaimer
        .PageSetup.DifferentFirstPageHeaderFooter = True
    End With
    
    Selection.HomeKey Unit:=wdStory
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    
    Selection.Font.Bold = True
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 11
    Selection.TypeText Text:="Project Topic"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" - Expert Name"
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Font.Bold = True
    Selection.TypeText Text:="Interaction ID:"
    Selection.Font.Bold = False
    Selection.TypeText Text:=" XXXXXXXX "
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Font.Italic = True
    Selection.Font.Size = 9
    
    Selection.TypeText Text:= _
        "This transcript was prepared by a third party transcriptioni"
    Selection.TypeText Text:= _
        "st who is acting independently of AlphaSights, and is subjec"
    Selection.TypeText Text:= _
        "t to confidentiality obligations to AlphaSights. Any editing"
    Selection.TypeText Text:= _
        " by AlphaSights has been done for ease of reading and the tr"
    Selection.TypeText Text:= _
        "anscript is not endorsed by AlphaSights, nor does it represe"
    Selection.TypeText Text:= _
        "nt the views of AlphaSights as a company or that of its empl"
    Selection.TypeText Text:= _
        "oyees. This transcript is provided for your general informat"
    Selection.TypeText Text:= _
        "io"
    Selection.TypeText Text:= _
        "n and it is not and does not purport to be: (i) investment, "
    Selection.TypeText Text:= _
        "financial, legal or professional advice; (ii) a recommendati"
    Selection.TypeText Text:= _
        "on or invitation as to how to proceed with any business or i"
    Selection.TypeText Text:= _
        "nvestment decision or decision as to whether to enter into o"
    Selection.TypeText Text:= _
        "r offer to enter into any agreement; (iii) independently ver"
    Selection.TypeText Text:= _
        "ified by AlphaSights (and AlphaSights accepts no liability f"
    Selection.TypeText Text:= _
        "or any inaccuracies of the contents); (iv) an accurate, comp"
    Selection.TypeText Text:= _
        "rehensive or complete summary of any interaction with an ind"
    Selection.TypeText Text:= _
        "ustry expert; or (v) a comprehensive summary of all matters "
    Selection.TypeText Text:= _
        "which may be relevant to the subject area. This transcript"
    Selection.TypeText Text:= _
        " is neither applicable to, nor should"
    Selection.TypeText Text:= _
        " relied on by, any third party."
        
    Selection.Font.Bold = False
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 11
    
    
End Sub

Sub RemoveHeaders()
    Dim sec As Section
    Dim hdr As HeaderFooter
    
    For Each sec In ActiveDocument.Sections
        For Each hdr In sec.Headers
            If hdr.Exists Then
                hdr.Range.Delete
            End If
        Next hdr
    Next sec
    
End Sub
