Attribute VB_Name = "NewMacros"
Sub runTemplateFix()
'
' runTemplateFix Macro
'
'

' Introduction
MsgBox "You have selected the Aurecon Template Automation Tool Ver 0.1.6 . Please click Ok to run"

' Body Text
templateFix_Body

' Headings
templateFix_Heading1
templateFix_Heading2
templateFix_Heading3
templateFix_Heading4
templateFix_Heading5
templateFix_Heading6
templateFix_Heading7

' Bullets
templateFix_Bullet1
templateFix_Bullet2
templateFix_Bullet3

' General
templateFix_Footer
templateFix_FooterText
templateFix_NoSpacing
templateFix_Normal

' Final Message
MsgBox "Aurecon Template Automation Tool has finished. Please direct any feedback to David.Petrie@aurecongroup.com"

End Sub

Sub templateFix_Heading1()
Attribute templateFix_Heading1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Heading1Fix"
'
' templateFix_Heading1 Macro
'
'
    With ActiveDocument.Styles("Heading 1").Font
        .Name = "+Headings"
        .Size = 14
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = -570392321
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 14
        .NameBi = "+Headings"
        .BoldBi = True
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Heading 1").ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.5)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 14.5
        .SpaceBeforeAuto = False
        .SpaceAfter = 5.65
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(-1.5)
        .OutlineLevel = wdOutlineLevel1
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Heading 1").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Heading 1")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Body Text"
    End With
    
    'templateFix_Bullet1
End Sub

Sub templateFix_Bullet1()
Attribute templateFix_Bullet1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Bullet1"
'
' templateFix_Bullet1 Macro
'
'
    With ActiveDocument.Styles("Bullet 1").Font
        .Name = "+Body"
        .Size = 10
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Body CS"
        .BoldBi = False
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Bullet 1").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.4)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 11
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(-0.4)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Bullet 1").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Bullet 1")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Bullet 1"
    End With
End Sub
Sub templateFix_Body()
Attribute templateFix_Body.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Body"
'
' templateFix_Body Macro
'
'
    With ActiveDocument.Styles("Body Text").Font
        .Name = "+Body"
        .Size = 10
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Body CS"
        .BoldBi = False
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Body Text").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 7.1
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Body Text").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Body Text")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Body Text"
    End With
    Selection.Style = ActiveDocument.Styles("Body Text")
End Sub
Sub templateFix_Bullet2()
Attribute templateFix_Bullet2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Bullet2"
'
' templateFix_Bullet2 Macro
'
'
    With ActiveDocument.Styles("Bullet 2").Font
        .Name = "+Body"
        .Size = 10
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Body CS"
        .BoldBi = False
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Bullet 2").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.8)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(-0.4)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Bullet 2").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Bullet 2")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Bullet 2"
    End With
End Sub
Sub templateFix_Heading2()
Attribute templateFix_Heading2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Heading2"
'
' templateFix_Heading2 Macro
'
'
    With ActiveDocument.Styles("Heading 2").Font
        .Name = "+Headings"
        .Size = 12
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = -721354753
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 12
        .NameBi = "+Body CS"
        .BoldBi = True
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Heading 2").ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.5)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 14.5
        .SpaceBeforeAuto = False
        .SpaceAfter = 5.65
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(-1.5)
        .OutlineLevel = wdOutlineLevel2
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Heading 2").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Heading 2")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Body Text"
    End With
End Sub
Sub templateFix_Bullet3()
Attribute templateFix_Bullet3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Bullet3"
'
' templateFix_Bullet3 Macro
'
'
    With ActiveDocument.Styles("Bullet 3").Font
        .Name = "+Body"
        .Size = 10
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Body CS"
        .BoldBi = False
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Bullet 3").ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.2)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 11
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(-0.4)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Bullet 3").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Bullet 3")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Bullet 3"
    End With
End Sub
Sub templateFix_Footer()
Attribute templateFix_Footer.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Footer"
'
' templateFix_Footer Macro
'
'
    With ActiveDocument.Styles("Footer").Font
        .Name = "+Body"
        .Size = 10
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Body CS"
        .BoldBi = False
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Footer").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Footer").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Footer")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Footer"
    End With
End Sub
Sub templateFix_FooterText()
Attribute templateFix_FooterText.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_FooterText"
'
' templateFix_FooterText Macro
'
'
    With ActiveDocument.Styles("Footer Text").Font
        .Name = "+Body"
        .Size = 6
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 6
        .NameBi = "+Body CS"
        .BoldBi = False
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Footer Text").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Footer Text").NoSpaceBetweenParagraphsOfSameStyle = _
         False
    With ActiveDocument.Styles("Footer Text")
        .AutomaticallyUpdate = False
        .BaseStyle = "Footer"
        .NextParagraphStyle = "Footer Text"
    End With
End Sub
Sub templateFix_Heading3()
Attribute templateFix_Heading3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Heading3"
'
' templateFix_Heading3 Macro
'
'
    With ActiveDocument.Styles("Heading 3").Font
        .Name = "+Headings"
        .Size = 10
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Headings"
        .BoldBi = True
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Heading 3").ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.5)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 14.5
        .SpaceBeforeAuto = False
        .SpaceAfter = 5.65
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(-1.5)
        .OutlineLevel = wdOutlineLevel3
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Heading 3").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Heading 3")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Body Text"
    End With
End Sub
Sub templateFix_Heading4()
Attribute templateFix_Heading4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Heading4"
'
' templateFix_Heading4 Macro
'
'
    With ActiveDocument.Styles("Heading 4").Font
        .Name = "+Body"
        .Size = 10
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = -721354753
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Headings CS"
        .BoldBi = True
        .ItalicBi = True
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Heading 4").ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.5)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 14.5
        .SpaceBeforeAuto = False
        .SpaceAfter = 5.65
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(-1.5)
        .OutlineLevel = wdOutlineLevel4
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Heading 4").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Heading 4")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Body Text"
    End With
End Sub
Sub templateFix_Heading5()
Attribute templateFix_Heading5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Heading5"
'
' templateFix_Heading5 Macro
'
'
    With ActiveDocument.Styles("Heading 5").Font
        .Name = "+Headings"
        .Size = 10
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = -721354753
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Headings CS"
        .BoldBi = False
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Heading 5").ParagraphFormat
        .LeftIndent = CentimetersToPoints(1.5)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 16
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = True
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(-1.5)
        .OutlineLevel = wdOutlineLevel5
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Heading 5").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Heading 5")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Body Text"
    End With
End Sub
Sub templateFix_Heading6()
Attribute templateFix_Heading6.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Heading6"
'
' templateFix_Heading6 Macro
'
'
    With ActiveDocument.Styles("Heading 6").Font
        .Name = "+Headings"
        .Size = 10
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Headings CS"
        .BoldBi = False
        .ItalicBi = True
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Heading 6").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 10
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = True
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevel6
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Heading 6").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Heading 6")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
    End With
End Sub
Sub templateFix_Heading7()
Attribute templateFix_Heading7.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Heading7"
'
' templateFix_Heading7 Macro
'
'
    With ActiveDocument.Styles("Heading 7").Font
        .Name = "+Headings"
        .Size = 10
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = -738131969
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Headings CS"
        .BoldBi = False
        .ItalicBi = True
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Heading 7").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 10
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = True
        .KeepTogether = True
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevel7
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Heading 7").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Heading 7")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
    End With
End Sub
Sub templateFix_NoSpacing()
Attribute templateFix_NoSpacing.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_NoSpacing"
'
' templateFix_NoSpacing Macro
'
'
    With ActiveDocument.Styles("No Spacing").Font
        .Name = "+Body"
        .Size = 10
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Body CS"
        .BoldBi = False
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("No Spacing").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("No Spacing").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("No Spacing")
        .AutomaticallyUpdate = False
        .BaseStyle = "Normal"
        .NextParagraphStyle = "No Spacing"
    End With
End Sub
Sub templateFix_Normal()
Attribute templateFix_Normal.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.templateFix_Normal"
'
' templateFix_Normal Macro
'
'
    With ActiveDocument.Styles("Normal").Font
        .Name = "+Body"
        .Size = 10
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 0
        .Animation = wdAnimationNone
        .SizeBi = 11
        .NameBi = "+Body CS"
        .BoldBi = False
        .ItalicBi = False
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Normal").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .ReadingOrder = wdReadingOrderLtr
    End With
    ActiveDocument.Styles("Normal").NoSpaceBetweenParagraphsOfSameStyle = _
        False
    With ActiveDocument.Styles("Normal")
        .AutomaticallyUpdate = False
        .BaseStyle = ""
        .NextParagraphStyle = "Normal"
    End With
End Sub
