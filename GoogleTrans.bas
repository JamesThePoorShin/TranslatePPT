Attribute VB_Name = "Module1"
Option Explicit

'// 번역할 소스언어와 타겟언어 선택
'// en, ko ja, zh-CN, fr, eo, de, it, fa, ru, hi ....
'// https://translate.google.com/m?sl=auto&tl=en&q=a&mui=tl&hl=en
Const SourceLanguage As String = "auto"         '"en"   '"auto"
Const TargetLanguage As String = "en"

'// 속도를 위해 전역변수 선언
'// 미리 변수를 정의하려면 도구-참조에서 MSXML 과 MS Html Object 라이브러리에 체크
'// to use early binding, goto Tools - References, check MSXML 6.0,  Microsoft HTML Object Library
Dim Http As Object  'MSXML2.ServerXMLHTTP
Dim Html As MSHTML.HTMLDocument

Sub TranslateSlides()

    Set Http = CreateObject("MSXML2.ServerXMLHTTP")
    'Set Html = CreateObject("HTMLfile")
    'Set Http = New MSXML2.ServerXMLHTTP
    Set Html = New MSHTML.HTMLDocument
    'Dim elem As Object  'IHTMLElementCollection
    Dim sldRng As SlideRange
    Dim sld As Slide
    Dim shp As Shape
    Dim i As Long

    Set sldRng = ActivePresentation.Slides.Range  '모든 슬라이드
   
    'UserForm1.Show vbModeless
    
    For Each sld In sldRng
    
        For Each shp In sld.Shapes
            'Debug.Print shp.Name
            TranslateShape shp
        
        Next shp
        
        'Update Progress Bar
        i = i + 1
        'UserForm1.Caption = "Translating " & i & " / " & sldRng.Count
        'UserForm1.ProgressBar1.Value = CInt(i * 100 / sldRng.Count)
        'Exit Sub
    Next sld
    
    'Unload UserForm1
    Set Html = Nothing
    Set Http = Nothing
    
End Sub

Function TranslateShape(oShp As Shape)

    Dim Txt As String
    Dim tr As TextRange
    Dim r As Integer, C As Integer
    Dim cShp As Shape
    
    '그룹 도형인 경우
    If oShp.Type = msoGroup Then
        
        For Each cShp In oShp.GroupItems
            TranslateShape cShp
        Next cShp
            
    '테이블(표)인 경우
    ElseIf oShp.Type = msoTable Then
        With oShp.Table
            For r = 1 To .Rows.Count
                For C = 1 To .Columns.Count
                    If Not IsMerged(oShp.Table, r, C) Or isTopLeftCell(oShp.Table, r, C) Then
                        Txt = .Cell(r, C).Shape.TextFrame.TextRange.Text
                        .Cell(r, C).Shape.TextFrame.TextRange.Text = GoogleTransTbl(Txt, TargetLanguage)
                    End If
                Next C
            Next r
        End With
            
    '기타
    ElseIf oShp.HasTextFrame Then
        With oShp.TextFrame
        If .HasText Then
            With .TextRange
                For Each tr In .Paragraphs
                    '구글번역으로 기존 텍스트를 한글로 변환
                    'Txt = Replace(tr.Text, "&", "") ' & 문자 제거
                    ' 빈 칸은 +로 대체
                    'Txt = Replace(Txt, " ", "+")
                    Txt = ENCODEURL(tr.Text)
                    'Target Language
                    '"en, ko ja, zh-CN, fr, eo, de, it, fa, ru, hi ...."
                    tr.Text = GoogleTrans(Txt, TargetLanguage) '해당 언어로 번역
                    'Debug.Print Txt
                    'Debug.Print GoogleTrans(Txt, "ko")
                Next tr
            End With
        End If
        End With
    End If
    
End Function

'simply check if the cell is merged
Function IsMerged(oTbl As Table, rr As Integer, cc As Integer) As Boolean

    Dim C As Cell
    
    'the current cell
    Set C = oTbl.Cell(rr, cc)
    
    'Check the width and height
    If C.Shape.width <> oTbl.Columns(cc).width Then IsMerged = True
    If C.Shape.height <> oTbl.Rows(rr).height Then IsMerged = True
    
End Function

'list the top-left cells of each merged area
Private Sub Test_isTopLeftCell()
    Dim r As Integer, C As Integer, i As Integer
    Dim shp As Shape
    
    If ActiveWindow.Selection.Type = ppSelectionNone Then _
        MsgBox "Select a table first": Exit Sub
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    If shp.Type = msoTable Then
        For r = 1 To shp.Table.Rows.Count
            For C = 1 To shp.Table.Columns.Count
                If isTopLeftCell(shp.Table, r, C) Then
                    i = i + 1
                    Debug.Print "First cell of the merged area #" & i & " starting at " & r; ", " & C
                End If
            Next C
        Next r
    End If
    
End Sub

'Check if the cell is the first(Top-Left) cell of the merged area
Function isTopLeftCell(oTbl As Table, rr As Integer, cc As Integer) As Boolean
    Dim i As Integer
    
    With oTbl.Cell(rr, cc).Shape
        'horozontally merged
        If .width <> oTbl.Columns(cc).width Then
            'count the left cells merged from the currnet cell
            For i = 1 To cc - 1
                If oTbl.Cell(rr, cc - i).Shape.Left <> .Left Then Exit For
            Next i
            'count the rows above
            If i = 1 Then
                For i = 1 To rr - 1
                    If oTbl.Cell(rr - i, cc).Shape.Top <> .Top Then Exit For
                Next i
                If i = 1 Then isTopLeftCell = True: Exit Function
            End If
        'vertically merged
        ElseIf .height <> oTbl.Rows(rr).height Then
            For i = 1 To rr - 1
                If oTbl.Cell(rr - i, cc).Shape.Top <> .Top Then Exit For
            Next i
            If i = 1 Then isTopLeftCell = True: Exit Function
        Else
            'isFirstCell = False
        End If
    End With
    
End Function

Function GoogleTransTbl(str As String, lang_out As String) As String
    Dim URL As String
    Dim lang_in As String

    GoogleTransTbl = ""
    ' INPUT LANGUAGE
    lang_in = SourceLanguage '"en, ko ja, zh-CN, fr, eo, de, it, fa, ru, hi ...."

    ' Construct the URL
    URL = "https://translate.google.com/m?hl=en&sl=" & lang_in _
            & "&tl=" & lang_out & "&ie=UTF-8&q=" & ENCODEURL(str)

    ' Create HTTP request
    With Http
        .Open "GET", URL, False
        .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; AS; rv:11.0) like Gecko"
        .Send
        Html.body.innerHTML = .responseText
    End With

    ' Get the translated text
    On Error Resume Next ' Ignore errors for now
    If Html.getElementsByClassName("t0").Length = 0 Then
        If Html.getElementsByClassName("result-container").Length = 0 Then
            GoogleTransTbl = "<Parsing Error>"
        Else
            GoogleTransTbl = Html.getElementsByClassName("result-container")(0).innerText
        End If
    Else
        GoogleTransTbl = Html.getElementsByClassName("t0")(0).innerText
    End If
    On Error GoTo 0 ' Reset error handling
End Function


Function GoogleTrans(str As String, lang_out As String) As String

    Dim URL As String
    Dim lang_in As String
    
    GoogleTrans = ""
    'INPUT LANGUAGE
    lang_in = SourceLanguage '"en, ko ja, zh-CN, fr, eo, de, it, fa, ru, hi ...."
    
    'open mobile website since the mobile web site is much simpler....
    URL = "https://translate.google.com/m?hl=en&sl=" & lang_in _
            & "&tl=" & lang_out & "&ie=UTF-8&q=" & str
    'Example: https://translate.google.com/m?hl=en&sl=auto&tl=fr&ie=UTF-8&q=morning
    'Debug.Print URL
 
    With Http
        .Open "GET", URL, False
        .SetRequestHeader "User-Agent", "Mobile"    '"Mozilla/5.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        '.SetRequestHeader "accept-language", "ko,en;q=0.9,ko-KR;q=0.8,en-US;q=0.7"
        .Send

        Html.body.innerHTML = .responseText
        'Debug.Print .responseText
    End With
    
    'On Error GoTo SKip:
    'MsgBox Html.body.innerHTML
    'Set elem = Html.getElementsByClassName("result-container")(0)   'late binding 의 경우 작동하지 않음.
    'Set elem = IE8_GetElementsByClassName(Html.body, "result-container")(0)
    'Debug.Print Html.body.innerHTML
    'Exit Function
    If Html.getElementsByClassName("t0").Length = 0 Then
        If Html.getElementsByClassName("result-container").Length = 0 Then
            GoogleTrans = "<Parsing Error>"
        Else
            GoogleTrans = Html.getElementsByClassName("result-container")(0).innerText
        End If
    Else
        GoogleTrans = Html.getElementsByClassName("t0")(0).innerText
    End If
SKip:
    'Debug.Print GoogleTrans
    If Err.Number Then GoogleTrans = str: Debug.Print Err.Number, Err.Description

End Function

Function ENCODEURL(varText As Variant, Optional blnEncode = True)
    Static objHtmlfile As Object
    If objHtmlfile Is Nothing Then
        Set objHtmlfile = CreateObject("htmlfile")
        With objHtmlfile.parentWindow
            .execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
        End With
    End If
    If blnEncode Then
        ENCODEURL = objHtmlfile.parentWindow.encode(varText)
    End If
End Function

'// 현재 선택된 도형 텍스트만 번역
Sub TranslateCurrentShape()
    
    Set Http = CreateObject("MSXML2.ServerXMLHTTP")
    Set Html = New MSHTML.HTMLDocument
    
    Dim shp As Shape
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    TranslateShape shp
    
    Set Html = Nothing
    Set Http = Nothing
    
End Sub

'//현재 선택한 도형 Language ID 변경
Private Sub ChangeLangID()
    Dim tr As TextRange
    
    For Each tr In ActiveWindow.Selection.ShapeRange(1).TextFrame.TextRange.Runs
        'tr.LanguageID = msoLanguageIDEnglishUS
        tr.LanguageID = msoLanguageIDJapanese
        'tr.LanguageID = msoLanguageIDKorean
        'Debug.Print tr.LanguageID '1041:Jap, 1042:Kor, 1033:EngUS
    Next tr
End Sub
Private Sub ChangeLangIDTextRange()
    Dim tr As TextRange
    
    Set tr = ActiveWindow.Selection.TextRange
    'tr.LanguageID = msoLanguageIDEnglishUS
    tr.LanguageID = msoLanguageIDJapanese

End Sub
'// 현재 선택한 도형 폰트 정보 조회
Private Sub ViewFont()

    Dim shp As Shape
    Dim tr As TextRange
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    Set tr = ActiveWindow.Selection.TextRange.Runs(1)   '.Characters(1)
    
    'For Each tr In shp.TextFrame.TextRange.Runs
    
        Debug.Print "Text: ["; tr.Text; "]"
        Debug.Print "Name: "; tr.Font.Name
        Debug.Print "NameAscii: "; tr.Font.NameAscii
        Debug.Print "NameComplexScript: "; tr.Font.NameComplexScript
        Debug.Print "NameFarEast: "; tr.Font.NameFarEast
        Debug.Print "NameOther: "; tr.Font.NameOther
        Debug.Print "Lang: "; tr.LanguageID '1041:Jap, 1042:Kor, 1033:Eng
        Debug.Print "================"
        
    'Next tr
End Sub

'// 현재 도형에 폰트 적용
Private Sub ApplyFont()

    Dim shp As Shape
    Dim tr As TextRange
    Dim FontName As String, FontNameFE As String
    
    FontName = "Arial Narrow"
    FontNameFE = "KoPub돋움체 Bold"
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    'Set tr = shp.TextFrame.TextRange.Characters(2)
    'Set tr = ActiveWindow.Selection.TextRange   '.Characters(1)
    For Each tr In shp.TextFrame.TextRange.Runs
        'tr.LanguageID = msoLanguageIDEnglishUS '1041:Jap, 1042:Kor, 1033:Eng
        tr.LanguageID = msoLanguageIDEnglishUS  '
        
        With tr.Font
            .Name = FontName
            .NameAscii = FontName
            .NameComplexScript = "+mn-cs"   'FontName
            .NameFarEast = FontNameFE
            .NameOther = ""     'FontName
        End With
        
    Next tr
    
End Sub
