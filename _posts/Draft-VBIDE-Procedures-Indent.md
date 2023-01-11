---
title: VBIDE - Procedures - Indent
# author: 'Anastasiou Alex'
# date: 2022-12-02 12:00:00 # if missing it's taken from the filename
# last_modified_at: #2022-12-02 12:00:00 
categories: [VBIDE, Procedures] # can handle 1 category and 1 subcategory eg [Category, Subcategory]
tags: [Indent] # [Tag1, Tag2 ...]
---



```vb

Function StringFormatIndent(txt As String) As String

'@AssignedModule F_Vbe_Indent
'@INCLUDE PROCEDURE IsBlockEnd
'@INCLUDE PROCEDURE IsBlockStart
    Dim str As Variant
    str = Split(txt, vbNewLine)
    Dim strNewLine As String
    Dim nIndent As Integer
    Dim i As Long
    For i = LBound(str) To UBound(str)
        strNewLine = str(i)
        strNewLine = LTrim$(strNewLine)
        If IsBlockEnd(strNewLine) Then
            nIndent = nIndent - 1
        End If
        If nIndent < 0 Then
            nIndent = 0
        End If
        If strNewLine <> "" Then
            str(i) = Space$(nIndent * 4) & strNewLine
        End If
        If IsBlockStart(LTrim$(strNewLine)) Then
            nIndent = nIndent + 1
        End If
    Next
    StringFormatIndent = Join(str, vbNewLine)
End Function

Public Sub IndentModule(Module As VBComponent)
'@AssignedModule F_Vbe_Indent
'@INCLUDE PROCEDURE StringFormatIndent
'@INCLUDE PROCEDURE ModuleCode
'@INCLUDE PROCEDURE ModuleCodeRemove
    Dim code As String
        code = ModuleCode(Module)
        code = StringFormatIndent(code)
    ModuleCodeRemove Module
    Module.CodeModule.AddFromString code
End Sub

Public Sub IndentProcedure( _
                          Optional TargetWorkbook As Workbook, _
                          Optional Module As VBComponent, _
                          Optional Procedure As String)
'@AssignedModule F_Vbe_Indent
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE PROCEDURE ProcedureReplace
'@INCLUDE PROCEDURE StringFormatIndent
'@INCLUDE PROCEDURE ProcedureCode
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Exit Sub
    Dim code As String
        code = ProcedureCode(TargetWorkbook, Module, Procedure)
        code = StringFormatIndent(code)
    ProcedureReplace Module, Procedure, code
End Sub

Public Sub IndentWorkbook(Optional TargetWorkbook As Workbook)
'@AssignedModule F_Vbe_Indent
'@INCLUDE PROCEDURE ActiveCodepaneWorkbook
'@INCLUDE PROCEDURE IndentModule
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        IndentModule Module
    Next
End Sub


Public Function IsBlockEnd(strLine As String) As Boolean
'@AssignedModule F_VBE
    strLine = Replace(strLine, Chr(13), "")
    Dim bOK As Boolean
    Dim nPos As Integer
    Dim strTemp As String
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = left$(strLine, nPos)
    Select Case strTemp
    Case "Next", "Loop", "Wend", "Case", "Else", "#Else", "Else:", "#Else:", "ElseIf", "#ElseIf", "#End"
        bOK = True
    Case "End"
        bOK = (Len(strLine) > 3)
    End Select
    IsBlockEnd = bOK
End Function

Public Function IsBlockStart(strLine As String) As Boolean
'@AssignedModule F_VBE
    strLine = Replace(strLine, Chr(13), "")
    Dim bOK As Boolean
    Dim nPos As Integer
    Dim strTemp As String
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = left$(strLine, nPos)
    Select Case strTemp
    Case "With", "For", "Do", "While", "Select", "Case", "Else", "Else:", "#Else", "#Else:", "Sub", "Function", "Property", "Enum", "Type"
        bOK = True
    Case "If", "#If", "ElseIf", "#ElseIf"
        '        bOK = (Len(strLine) = (InStr(1, strLine, " Then") + 4))
        bOK = (Right(strLine, 4) = "Then" Or Right(strLine, 1) = "_")
    Case "Private", "Public", "Friend"
        nPos = InStr(1, strLine, " Static ")
        If nPos Then
            nPos = InStr(nPos + 7, strLine, " ")
        Else
            nPos = InStr(Len(strTemp) + 1, strLine, " ")
        End If
        On Error GoTo skip
        Select Case Mid$(strLine, nPos + 1, InStr(nPos + 1, strLine, " ") - nPos - 1)
        Case "Sub", "Function", "Property", "Enum", "Type"
            bOK = True
        End Select
skip:
        On Error GoTo 0
    End Select
    IsBlockStart = bOK
End Function


Public Sub ProcedureReplace( _
                            Module As VBComponent, _
                            Procedure As String, _
                            code As String)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE ModuleOfProcedure
    
    Dim startLine As Integer
    Dim NumLines As Integer
    With Module.CodeModule
        startLine = .ProcStartLine(Procedure, vbext_pk_Proc)
        NumLines = .ProcCountLines(Procedure, vbext_pk_Proc)
        .DeleteLines startLine, NumLines
        .InsertLines startLine, code
    End With
End Sub

```