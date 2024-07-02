Attribute VB_Name = "Module2"
Option Explicit

'// �ӵ��� ���� �������� ����
'// �̸� ������ �����Ϸ��� ����-�������� MSXML �� MS Html Object ���̺귯���� üũ
'// to use early binding, goto Tools - References, check MSXML 6.0,  Microsoft HTML Object Library
Dim Http As Object  'MSXML2.ServerXMLHTTP
Dim Html As MSHTML.HTMLDocument



Sub EngZip()

    Set Http = CreateObject("MSXML2.ServerXMLHTTP")
    'Set Html = CreateObject("HTMLfile")
    'Set Http = New MSXML2.ServerXMLHTTP
    Set Html = New MSHTML.HTMLDocument
    'Dim elem As Object  'IHTMLElementCollection
    Dim sldRng As SlideRange
    Dim sld As Slide
    Dim shp As Shape
    Dim i As Long

    Set sldRng = ActivePresentation.Slides.Range  '��� �����̵�
   
    'UserForm1.Show vbModeless
    
    For Each sld In sldRng
    
        For Each shp In sld.Shapes
            'Debug.Print shp.Name
            TransShp shp
        
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

'// ���� ���õ� ���� �ؽ�Ʈ�� ����
Sub EngZipCurrentShape()
    
    Set Http = CreateObject("MSXML2.ServerXMLHTTP")
    Set Html = New MSHTML.HTMLDocument
    
    Dim shp As Shape
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    TransShp shp
        
    Set Html = Nothing
    Set Http = Nothing
    
End Sub


Function RemoveLB(inputString As String) As String 'Line Breaks (�ٹٲ�) ����
    Dim resultString As String
    
    ' Remove line breaks
    resultString = Replace(inputString, vbCrLf, "")
    resultString = Replace(resultString, vbCr, "")
    resultString = Replace(resultString, vbLf, "")
    
    RemoveLB = resultString
End Function


Function TransShp(oShp As Shape)

    Dim Txt As String
    Dim tr As TextRange
    Dim r As Integer, C As Integer
    Dim cShp As Shape
    
    '�׷� ������ ���
    If oShp.Type = msoGroup Then
        
        For Each cShp In oShp.GroupItems
            TransShp cShp
        Next cShp
            
    '���̺�(ǥ)�� ���
    ElseIf oShp.Type = msoTable Then
        With oShp.Table
            For r = 1 To .Rows.Count
                For C = 1 To .Columns.Count
                    If Not IsMerged(oShp.Table, r, C) Or isTopLeftCell(oShp.Table, r, C) Then
                        Txt = .Cell(r, C).Shape.TextFrame.TextRange.Text
                        If CountWords(Txt) > 2 Then
                            .Cell(r, C).Shape.TextFrame.TextRange.Text = RemoveLB(PyEngSumm(Txt))
                        End If
                    End If
                Next C
            Next r
        End With
            
    '��Ÿ
    ElseIf oShp.HasTextFrame Then
        With oShp.TextFrame
        If .HasText Then
            With .TextRange
                For Each tr In .Paragraphs
                    If CountWords(tr.Text) > 2 Then
                        tr.Text = RemoveLB(PyEngSumm(tr.Text)) '���� ����
                    End If
                Next tr
            End With
        End If
        End With
    End If
    
End Function


Function PyEngSumm(inputSentence As String)
   Dim pythonPath As String
   Dim pythonExe As String
   Dim pythonScriptPath As String
   'Dim inputSentence As String
   Dim outputSentence As String
   
   ' Set the path to Python executable
   pythonPath = "C:\Users\skb.3895\PycharmProjects\pythonProject\.venv\Scripts\"
   pythonExe = "python.exe"
   
   ' Set the path to your Python script
   pythonScriptPath = "C:\Users\skb.3895\PycharmProjects\pythonProject\.venv\Scripts\EngSumm.py"
   
   ' Get the input sentence from the user
   'inputSentence = InputBox("Enter a sentence:", "Input")
   
   ' Call the Python script with the input sentence as an argument
   PyEngSumm = CallPythonScript(pythonPath, pythonExe, pythonScriptPath, inputSentence)
   
   ' Print the output sentence returned from the Python function
   'MsgBox "Output from Python function: " & outputSentence, vbInformation

End Function

Function CallPythonScript(pythonPath As String, pythonExe As String, pythonScriptPath As String, inputSentence As String) As String

   Dim objShell As Object
   Dim scriptFile As String
   Dim cmd As String
   Dim output As String
   
   ' Construct the command to run the Python script with input sentence as argument
   scriptFile = pythonPath & pythonExe
   cmd = """" & scriptFile & """" & " " & """" & pythonScriptPath & """" & " " & """" & inputSentence & """"
   
   ' Create a Shell object to run the command
   Set objShell = VBA.CreateObject("WScript.Shell")
   
   ' Execute the command and capture the output
   output = objShell.Exec(cmd).StdOut.ReadAll
   
   ' Close the Shell object
   Set objShell = Nothing
   
   ' Return the output from the Python function
   CallPythonScript = output

End Function

Function CountWords(inputText As String) As Long
    ' Function to count the number of words in the input text
    Dim words() As String
    words = Split(inputText, " ")
    CountWords = UBound(words) + 1
End Function
