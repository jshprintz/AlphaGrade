VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QA_Comments 
   Caption         =   "QA Comments"
   ClientHeight    =   9525.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   OleObjectBlob   =   "QA_Comments.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "QA_Comments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i As Integer
Private mCommentDispSuffix, mHeadlineStr As String
Private mSize As Variant
Private mResize As Double

Private Sub Cancel_Click()

   Unload Me

End Sub

Private Sub Insert_Click()

On Error GoTo ErrorHandler

Dim FinalDisp As Boolean

FinalDisp = True
               
    ' Display the final error in a comment on the righthand margin
    If FinalDisp = True Then
    
            If mCommentDispSuffix = "Header incorrect" Then
                Selection.HomeKey Unit:=wdStory
            End If
            
                ActiveDocument.Comments.Add _
                Range:=Selection.Range, Text:=mCommentDispSuffix
                
            #If Mac Then
            
            #Else
                ActiveDocument.RemoveDocumentInformation (wdRDIRemovePersonalInformation)
                ActiveDocument.Save
            #End If
    End If
    
    
    Exit Sub
    
ErrorHandler:

Select Case Err.Number

    Case 5935
    MsgBox ("Please do not make comments in the header." & vbCr & vbCr & "Please make all comments in the body of the text.")
    
    Selection.HomeKey Unit:=wdStory
    
End Select
Resume Next

End Sub

Private Sub Capitalization_Click()

Call Marquee

Headline.Caption = mHeadlineStr

Call Clean
        
        mCommentDispSuffix = "Capitalization"

End Sub


Private Sub Mishear_Click()

Call Marquee

Headline.Caption = mHeadlineStr

Call Clean
        
        mCommentDispSuffix = "Mishear"

End Sub


Private Sub Header_Wrong_Click()

Call Marquee

Headline.Caption = mHeadlineStr
        
Call Clean
        
        mCommentDispSuffix = "Header incorrect"

End Sub

Private Sub MissingContent_Click()

Call Marquee

Headline.Caption = "Please select an option below."

Call Clean
        
        mCommentDispSuffix = "Missing word(s)"
        
        Call MCSelect_Change
        
        MCSelect.Locked = False
        MCSelect.Visible = True
        
End Sub

Private Sub MCSelect_Change()


        If MCSelect.Value = 0 Then
            mCommentDispSuffix = "Missing word(s)"
        ElseIf MCSelect.Value = 1 Then
            mCommentDispSuffix = "Missing sentence(s)"
        ElseIf MCSelect.Value = 2 Then
            mCommentDispSuffix = "Under 2 minutes missing"
        ElseIf MCSelect.Value = 3 Then
            mCommentDispSuffix = "Over 2 minutes, under 5 minutes missing"
        ElseIf MCSelect.Value = 4 Then
            mCommentDispSuffix = "Over 5 minutes missing"
        ElseIf MCSelect.Value = 5 Then
            mCommentDispSuffix = "Nothing transcribed"
        End If
        
End Sub


Private Sub AlphaView_Click()

Call Marquee

Headline.Caption = "Please select an option below."

Call Clean
        
        mCommentDispSuffix = "AlphaView| Remove greetings"

        Call cboxAlphaView_Change
        
        cboxAlphaView.Locked = False
        cboxAlphaView.Visible = True
        
End Sub

Private Sub cboxAlphaView_Change()

        If cboxAlphaView.Value = 0 Then
            mCommentDispSuffix = "AlphaView| Remove greetings"
        ElseIf cboxAlphaView.Value = 1 Then
            mCommentDispSuffix = "AlphaView| Remove interviewer background"
        ElseIf cboxAlphaView.Value = 2 Then
            mCommentDispSuffix = "AlphaView| Remove names"
        ElseIf cboxAlphaView.Value = 3 Then
            mCommentDispSuffix = "AlphaView| Wrong formatting"
        End If

End Sub


Private Sub Punctuation_Error_Click()

Call Marquee

        Headline.Caption = mHeadlineStr

Call Clean
        
        mCommentDispSuffix = "Punctuation"

End Sub

Private Sub SpeakerID_Click()

Call Marquee

        Headline.Caption = mHeadlineStr

Call Clean

        mCommentDispSuffix = "Wrong speaker ID"

End Sub

Private Sub StyleGuide_Click()

Call Marquee

        Headline.Caption = "Please select an option below."

Call Clean
        
        mCommentDispSuffix = "Style Guide: Always use US English"
        
        Call SGSelect_Change
        
        SGSelect.Visible = True
        SGSelect.Locked = False

End Sub

Private Sub SGSelect_Change()

        If SGSelect.Value = 0 Then
            mCommentDispSuffix = "Style Guide: Always use US English"
        ElseIf SGSelect.Value = 1 Then
            mCommentDispSuffix = "Style Guide: Currency"
        ElseIf SGSelect.Value = 2 Then
            mCommentDispSuffix = "Style Guide: Measurement"
        ElseIf SGSelect.Value = 3 Then
            mCommentDispSuffix = "Style Guide: Numbers"
        End If
End Sub
Private Sub Typo_Click()

Call Marquee

        Headline.Caption = mHeadlineStr

Call Clean
        
        mCommentDispSuffix = "Typo/Spelling"

End Sub

Private Sub Wrong_Color_Select_Click()

Call Marquee

Headline.Caption = mHeadlineStr

Call Clean

        mCommentDispSuffix = "Wrong speaker color"

End Sub
Private Sub IncorrectTag_Click()

Headline.Caption = "Please select an option below."

Call Clean
        
        mCommentDispSuffix = "Tag| Incorrect Tag Format"
        
        Call ITSelect_Change
        
        ITSelect.Locked = False
        ITSelect.Visible = True

End Sub

Private Sub ITSelect_Change()

        If ITSelect.Value = 0 Then
            mCommentDispSuffix = "Tag| Incorrect Format"
        ElseIf ITSelect.Value = 1 Then
            mCommentDispSuffix = "Tag| Incorrect Timestamp Format"
        ElseIf ITSelect.Value = 2 Then
            mCommentDispSuffix = "Tag| M/W Word in Research Bracket"
        ElseIf ITSelect.Value = 3 Then
            mCommentDispSuffix = "Tag| Wrong Timestamp"
        End If

End Sub

Private Sub WrongTemplate_Click()

Call Marquee

Headline.Caption = mHeadlineStr

Call Clean
        
        mCommentDispSuffix = "Wrong template formatting"

End Sub

Private Sub Inconsistent_Click()

Call Marquee

Headline.Caption = mHeadlineStr

Call Clean
        
        mCommentDispSuffix = "Consistency "
        
End Sub

Private Sub Extra_Words_Click()

Call Marquee

Headline.Caption = mHeadlineStr

Call Clean
        
        mCommentDispSuffix = "Adding words not spoken"
      
End Sub
Private Sub UserForm_Initialize()

mResize = 1
        
        With MCSelect
            .Locked = True
            .Visible = False
            .Font.Size = 9
            .AddItem "a word or two"
            .AddItem "a sentence or two"
            .AddItem "< 2 min"
            .AddItem "> 2 < 5 min"
            .AddItem "> 5 min"
            .AddItem "Nothing Transcribed"
            .Style = fmStyleDropDownList
            .BoundColumn = 0
            .ListIndex = 0
            .MatchRequired = True
        End With
        
        With ITSelect
            .Locked = True
            .Visible = False
            .Font.Size = 9
            .AddItem "Incorrect Tag Format"
            .AddItem "Incorrect Timestamp Format"
            .AddItem "M/W word in Research Tag"
            .AddItem "Wrong Timestamp"
            .Style = fmStyleDropDownList
            .BoundColumn = 0
            .ListIndex = 0
            .MatchRequired = True
        End With
        
        With SGSelect
            .Locked = True
            .Visible = False
            .Font.Size = 9
            .AddItem "Always use US English"
            .AddItem "Currency"
            .AddItem "Measurement"
            .AddItem "Numbers"
            .Style = fmStyleDropDownList
            .BoundColumn = 0
            .ListIndex = 0
            .MatchRequired = True
        End With
                   
        With cboxAlphaView
            .Locked = True
            .Visible = False
            .Font.Size = 9
            .AddItem "Remove greetings"
            .AddItem "Remove interviewer background"
            .AddItem "Remove names"
            .AddItem "Wrong formatting"
            .Style = fmStyleDropDownList
            .BoundColumn = 0
            .ListIndex = 0
            .MatchRequired = True
        End With
             
             mCommentDispSuffix = "Mishear"
             

End Sub

Private Sub ResizeUserForm(frm As Object, Optional dResizeFactor As Double = 0#)

'THIS RESIZES THE USERFORM WINDOW
  
  Dim ctrl As Control
  Dim sColWidths As String
  Dim vColWidths As Variant
  Dim iCol As Long

    dResizeFactor = mResize
  
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

Private Sub Marquee()

Dim RndNum As Integer
RndNum = Int((5 * Rnd) + 1)    ' Generate random value between 1 and 5.

If RndNum = 1 Then
    mHeadlineStr = "Remember, not every edit is an error!"
ElseIf RndNum = 2 Then
    mHeadlineStr = "Don't double ding!"
ElseIf RndNum = 3 Then
    mHeadlineStr = "It's only a crucial error if it changes the meaning of the sentence!"
ElseIf RndNum = 4 Then
    mHeadlineStr = "Always take audio quality into consideration!"
ElseIf RndNum = 5 Then
    mHeadlineStr = "The macro automatically deducts for cleared tags. Don't double ding!"
End If

End Sub

Private Sub Clean()

        'resets the form from the combo boxes
        
        ITSelect.Locked = True
        ITSelect.Visible = False
        MCSelect.Locked = True
        MCSelect.Visible = False
        SGSelect.Locked = True
        SGSelect.Visible = False
        cboxAlphaView.Locked = True
        cboxAlphaView.Visible = False
             
End Sub

Private Sub Plus_Click()

mResize = 1.05
ResizeUserForm Me

End Sub

Private Sub Minus_Click()

mResize = 0.95
ResizeUserForm Me

End Sub
