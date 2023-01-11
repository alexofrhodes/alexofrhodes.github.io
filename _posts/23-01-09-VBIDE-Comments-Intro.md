---
title: VBIDE - Comments used well
# author: 'Anastasiou Alex'
# date: 2022-12-02 12:00:00 # if missing it's taken from the filename
last_modified_at: #2022-12-02 12:00:00 
categories: [VBIDE, Comments] # can handle 1 category and 1 subcategory eg [Category, Subcategory]
# tags: [comments] # [Tag1, Tag2 ...]
excerpt_separator: <!--more-->
---

<!--more-->

# Getting started

Comments are non executable lines in your modules and procedure. They start with a **single quote (')** or with the keyword **Rem** (from Remember) and they don't need to start at column 1. Single quote comments can also be placed inline, at it's end.

```vb
'Own line comment
If something then 'In Line comment
Rem Own Line Comment
```

```vb
'Comment before procedure title
Sub something()
'comment after procedure title

'comment in own-line
<code> 'in-line comment

'comment before procedure end 
End Sub
'comment after procedure end
```

# Use cases

**Add placeholders to your Module**, kind of like a story, before you even start coding, for example: 

```vb
'-------------MAIN IDEA----------------
'Indent code stored in a text file
'--------------------------------------

'---Sub---
'replace the content of a txt file
'or create file if it doesn't exit.

'---Function---
'return the content of a txt file

'---Sub---
'Format a string variable to proper code indentation

'---Function---
'return the number of spaces needed for each line
```

**Explain Reasoning** but avoid stating the obvious
```vb
Sub Test(var1 as variant)
'var1 needs to be variant because...
if isMissing(var1)

End Sub
```

**Mention External References**

```vb
Sub Test()
'source: https://github.com/...
'further reading: www...
'                 www...
End Sub
```

**Todo Notes**

```vb
'* @TODO Created: 04-01-2023 15:05 Author: username
'* @TODO Add tests
'* @TODO Refactor

Sub Test()

<code>

'* @TODO Created: 04-01-2023 17:05 Author: username
'* @TODO Room for speed improvement

<code>

End Sub
```
**Modification Notes**

```vb
'* Modified   : Date and Time       Author              Description
'* Updated    : 01-01-2023 16:18    Alex                Added this
'* Updated    : 02-01-2023 16:18    Alex                Modified this
'* Updated    : 03-01-2023 16:18    Alex                Deleted this

Sub Test()
'@LastModDate 2301101156

End Sub
```

**Module or Procedure Information**

```vb
'* * * * * * * * * * * * * * * * * * * * * * * 
'* Module     : F_Vbe_Insert
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-01-2023 17:36    Alex
'* * * * * * * * * * * * * * * * * * * * * * *
```

```vb
'* * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetGivenRowOfData
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'* Note       :
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-01-2023 17:36    Alex
'* Updated    : 04-01-2023 17:38    Alex                
'*
'* Argument(s):             Description
'*
'* Data As Variant       :
'* rowIndexes As Variant :
'* colIndexes As Variant :
'*
'* * * * * * * * * * * * * * * * * * * * * * *

Sub GetGivenRowOfData(Data As Variant, rowIndexes As Variant, colIndexes As Variant) As Variant

End Sub
```

**List Prerequisites**

```vb
Sub Test()
'@INCLUDE Reference VBIDE
'@INCLUDE Declaration Sleep
'@INCLUDE Procedure SecondTest
'@INCLUDE Userform uCustomMessage
'@INCLUDE Class clsTimer
'@INCLUDE Sheet UserSettings
<code>
End Sub
```

**Hold Actionable Information**

```vb
Sub Test()
'@IgnoreProcedure
'@AssignedModule M_Arrays
'@AssignedFolder Conversion
'@INCLUDE Declaration Sleep
'@INCLUDE Procedure SecondTest
End Sub
```

**Footnotes**

```vb
Sub Test()
<code> '[1]
<code> '[FurtherReading]
<code> '[Tag]
...
...
...
'[1] : an alternative way to do this is..
'[FurtherReading] : www...
'[FurtherReading] : www...
'[Tag] :
End Sub
```

# Things to consider

1. Comments in the procedure's header space should not affect the understanding of the macro if omitted.
2. After the procedure's title, keep the important information about the procedure and actionable comments.
3. Prefer to put comments to own line for better readability
4. Keep own-line comments aligned to the column start of the line/block they refer to
5. Keep  in-line comments aligned with other comments in the same block if possible.
6. There should be no comments kept after a macro.

# Macros

## Move 

```vb
Sub testCommentsMoveToOwnLine()
    Dim code As String
        code = "First line 'comment 1" & vbNewLine & "    Second line indented 'comment 2"
    Debug.Print "---INPUT---" & String(2, vbNewLine) & code & vbNewLine
        code = CommentsMoveToOwnLine(code)
    Debug.Print "---OUTPUT---" & String(2, vbNewLine) & code
End Sub

Function CommentsMoveToOwnLine(ByVal txt As String) As String
'@INCLUDE PROCEDURE CommentsTrim
'@AssignedModule F_Vbe_Comments

    Dim var As Variant
    ReDim var(0)
    Dim str As Variant
        str = Split(txt, vbNewLine)
    
    Dim N               As Long
    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim l               As Long
    Dim LineText        As String
    Dim QUOTES          As Long
    Dim Q               As Long
    Dim StartPos        As Long
    
    For j = LBound(str) To UBound(str)
        LineText = Trim(str(j))
        StartPos = 1
retry:
        N = InStr(StartPos, LineText, "'")
        Q = InStr(StartPos, LineText, """")
        QUOTES = 0
        If Q < N Then
            For l = 1 To N
                If Mid(LineText, l, 1) = """" Then
                    QUOTES = QUOTES + 1
                End If
            Next l
        End If
        If QUOTES = Application.WorksheetFunction.Odd(QUOTES) Then
            StartPos = N + 1
            GoTo retry:
        Else
            Select Case N
                Case Is = 0
                    var(UBound(var)) = str(j)
                    ReDim Preserve var(UBound(var) + 1)
                Case Is = 1
                    var(UBound(var)) = CommentsTrim(Array(str(j)))
                    ReDim Preserve var(UBound(var) + 1)
                Case Is > 1
                    var(UBound(var)) = Space(Len(str(j)) - Len(LTrim(str(j)))) & Mid(LineText, N)
                    ReDim Preserve var(UBound(var) + 1)
                    var(UBound(var)) = Space(Len(str(j)) - Len(LTrim(str(j)))) & left(LineText, N - 1)
                    ReDim Preserve var(UBound(var) + 1)
            End Select
        End If
    Next j

     CommentsMoveToOwnLine = Join(var, vbNewLine)
     CommentsMoveToOwnLine = left(CommentsMoveToOwnLine, Len(CommentsMoveToOwnLine) - Len(vbNewLine))
End Function

Sub CommentsMoveToOwnLineProcedure( _
                                Optional TargetWorkbook As Workbook, _
                                Optional Module As VBComponent, _
                                Optional Procedure As String)
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE PROCEDURE ProcedureReplace
'@INCLUDE PROCEDURE CommentsMoveToOwnLine
'@INCLUDE PROCEDURE ProcedureCode
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    Dim code As String
        code = ProcedureCode(TargetWorkbook, Module, Procedure)
        code = CommentsMoveToOwnLine(code)
    ProcedureReplace Module, Procedure, code
End Sub

Sub CommentsMoveToOwnLineModule(Optional Module As VBComponent)
    If Module Is Nothing Then Set Module = ActiveModule
    Dim code As String
        code = ModuleCode(Module)
        code = CommentsMoveToOwnLine(code)
    With Module.CodeModule
        If .CountOfLines = 0 Then Exit Sub
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, code
    End With
End Sub



Sub testCommentsTrim()
    Dim code As String
        code = "'    comment 1" & vbNewLine & "'    comment 2"
    Debug.Print "---INPUT---" & String(2, vbNewLine) & code & vbNewLine
        code = CommentsTrim(code)
    Debug.Print "---OUTPUT---" & String(2, vbNewLine) & code
End Sub

Function CommentsTrim(ByVal txt As String) As String
'@INCLUDE PROCEDURE ArrayRemoveEmptyElements
'@AssignedModule F_Vbe_Comments
    Dim var As Variant
    ReDim var(0)
    Dim str As Variant
        str = Split(txt, vbNewLine)
    For j = LBound(str) To UBound(str)
        LineText = Trim(str(j))
        If left(LineText, 2) = "' " Then
            tmp = Mid(LineText, 2)
            dif = Len(tmp) - Len(LTrim(tmp))
            var(UBound(var)) = Space(dif) & "'" & LTrim(tmp)
            ReDim Preserve var(UBound(var) + 1)
        Else
            var(UBound(var)) = str(j)
            ReDim Preserve var(UBound(var) + 1)
        End If
    Next
    var = ArrayRemoveEmptyElements(var)
    CommentsTrim = Join(var, vbNewLine)
End Function
```

## Align

```vb
Sub testCommentsAlign()
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE AlignTextParts
    Dim code As String
    code = "DoEvents   'apiidjs afsp" & vbNewLine
    code = code & "        DoEvents 'pawo" & vbNewLine
    code = code & "DoEvents 'epsi asd"
    
    Debug.Print "---INPUT---" & String(2, vbNewLine) & code & vbNewLine
    Debug.Print "---OUTPUT---" & String(2, vbNewLine) & AlignTextParts(code, "'")
End Sub

Sub cpsFormatCommentsAlign()
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE AlignCodepaneLineElements
    AlignCodepaneLineElements "'"
End Sub

Sub AlignCodepaneLineElements(AlignString As String, Optional AlignAtColumn As Long)
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE AlignTextParts
'@INCLUDE PROCEDURE cpsLinesCode
'@INCLUDE PROCEDURE cpsLineFirst
'@INCLUDE PROCEDURE cpsLinesCount
'@INCLUDE PROCEDURE ActiveModule
    Dim code As String
        code = cpsLinesCode
        code = AlignTextParts(code, AlignString, AlignAtColumn)
    Dim LineFirst As Long
        LineFirst = cpsLineFirst
    
    ActiveModule.CodeModule.DeleteLines LineFirst, cpsLinesCount
    ActiveModule.CodeModule.InsertLines LineFirst, code
    
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 05-01-2023 14:01    Alex                (AlignTextParts)

Function AlignTextParts(txt As String, AlignString As String, Optional AlignAtColumn As Long)
'@AssignedModule F_Vbe_Comments
    Dim TextLines
        TextLines = Split(txt, vbNewLine)
    Dim elementOriginalColumn As Long
    Dim rightMostColumn As Long
    Dim LineText As String
    Dim numberOfSpacesToInsert As Long
    Dim i As Long
    For i = LBound(TextLines) To UBound(TextLines)
        LineText = TextLines(i)
        elementOriginalColumn = InStrRev(LineText, AlignString)
        If elementOriginalColumn > rightMostColumn Then rightMostColumn = elementOriginalColumn
    Next
    If AlignAtColumn = 0 Then AlignAtColumn = rightMostColumn
    For i = LBound(TextLines) To UBound(TextLines)
        LineText = TextLines(i)
        elementOriginalColumn = InStrRev(LineText, AlignString)
        If elementOriginalColumn > 0 Then
            numberOfSpacesToInsert = AlignAtColumn - elementOriginalColumn
            If numberOfSpacesToInsert > 0 Then
                elementOriginalColumn = InStrRev(LineText, AlignString)
                TextLines(i) = left(TextLines(i), elementOriginalColumn - 1) & _
                                Space(numberOfSpacesToInsert) & _
                                Mid(TextLines(i), elementOriginalColumn)
            End If
        End If
    Next
    
    AlignTextParts = Join(TextLines, vbNewLine)
    
End Function
```

## Remove 

```vb
Sub testCommentsRemove()
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE CommentsRemove
    Dim code As String
        code = "'comment" & vbNewLine & "if something then 'comment" & vbNewLine & "doevents" & vbNewLine & "'comment"
    Debug.Print "---INPUT---" & String(2, vbNewLine) & code & vbNewLine
        code = CommentsRemove(code, True)
    Debug.Print "---OUTPUT---" & String(2, vbNewLine) & code
End Sub

Function CommentsRemove(ByVal txt As String, RemoveRem As Boolean) As String
'modified from Jacob Hilderbrand's code, found at
'http://www.vbaexpress.com/kb/getarticle.php?kb_id=266
    Dim var As Variant
    ReDim var(0)
    Dim str
        str = Split(txt, vbNewLine)
        str = ArrayRemoveEmptyElements(str)
    Dim N               As Long
    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim l               As Long
    Dim LineText        As String
    Dim QUOTES          As Long
    Dim Q               As Long
    Dim StartPos        As Long
    
    For j = LBound(str) To UBound(str)
        LineText = LTrim(str(j))
        If RemoveRem Then If LineText Like "Rem *" Then GoTo SKIP
        StartPos = 1
retry:
        N = InStr(StartPos, LineText, "'")
        Q = InStr(StartPos, LineText, """")
        QUOTES = 0
        If Q < N Then
            For l = 1 To N
                If Mid(LineText, l, 1) = """" Then
                    QUOTES = QUOTES + 1
                End If
            Next l
        End If
        If QUOTES = Application.WorksheetFunction.Odd(QUOTES) Then
            StartPos = N + 1
            GoTo retry:
        Else
            Select Case N
                Case Is = 0
                    If Len(LineText) > 0 Then
                        var(UBound(var)) = str(j)
                        ReDim Preserve var(UBound(var) + 1)
                    End If
                Case Is = 1
                    '
                Case Is > 1
                    var(UBound(var)) = left(str(j), N - 1)
                    ReDim Preserve var(UBound(var) + 1)
            End Select
        End If
SKIP:
    Next j
    var = ArrayRemoveEmptyElements(var)
    CommentsRemove = Join(var, vbNewLine)
End Function

Public Sub CommentsRemoveFromProcedure( _
                            Optional TargetWorkbook As Workbook, _
                            Optional Module As VBComponent, _
                            Optional Procedure As String, _
                            Optional RemoveRem As Boolean)
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE PROCEDURE ProcedureReplace
'@INCLUDE PROCEDURE CommentsRemove
'@INCLUDE PROCEDURE ProcedureCode
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    Dim code As String
        code = ProcedureCode(TargetWorkbook, Module, Procedure)
        code = CommentsRemove(code, RemoveRem)
    ProcedureReplace Module, Procedure, code
End Sub

Public Sub CommentsRemoveFromModule( _
                                    Module As VBComponent, _
                                    RemoveRem As Boolean)
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE CommentsRemove
'@INCLUDE PROCEDURE ModuleCode
'@INCLUDE PROCEDURE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    Dim code As String
        code = ModuleCode(Module)
        code = CommentsRemove(code, RemoveRem)
    With Module.CodeModule
        If .CountOfLines = 0 Then Exit Sub
        .DeleteLines 1, .CountOfLines
        .InsertLines 1, code
    End With
End Sub

Public Sub CommentsRemoveFromWorkbook(TargetWorkbook As Workbook, RemoveRem As Boolean)
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE ActiveCodepaneWorkbook
'@INCLUDE PROCEDURE CommentsRemoveFromModule
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBIDE.VBComponent
    For Each Module In ActiveCodepaneWorkbook.VBProject.VBComponents
        CommentsRemoveFromModule Module, RemoveRem
    Next
End Sub
```

## Convert 

```vb

Public Sub ReplaceQuoteWithRemInProcedure( _
                                        Optional TargetWorkbook As Workbook, _
                                        Optional Module As VBComponent, _
                                        Optional Procedure As String)
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE PROCEDURE ActiveProcedure
'@INCLUDE PROCEDURE ProcedureLinesFirst
'@INCLUDE PROCEDURE ProcedureLinesLast
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    Dim N As Long
    Dim s As String
    With Module.CodeModule
        For N = ProcedureLinesLast(Module, ActiveProcedure) To ProcedureLinesFirst(Module, ActiveProcedure) Step -1
            s = .Lines(N, 1)
            If left(Trim(s), 1) = "'" Then
                .ReplaceLine N, Replace(s, "'", "Rem ", , 1)
            End If
        Next N
    End With
End Sub

Public Sub ReplaceQuoteWithRemInModule(Optional Module As VBComponent)
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    Dim N As Long
    Dim s As String
    With Module.CodeModule
        For N = .CountOfLines To 1 Step -1
            If .CountOfLines = 0 Then Exit For
            s = .Lines(N, 1)
            If left(Trim(s), 1) = "'" Then
                .ReplaceLine N, Replace(s, "'", "Rem ", , 1)
            End If
        Next N
    End With
End Sub

Public Sub ReplaceQuoteWithRemInWorkbook(Optional TargetWorkbook As Workbook)
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE ActiveCodepaneWorkbook
'@INCLUDE PROCEDURE ReplaceQuoteWithRemInModule
'@INCLUDE DECLARATION vbComp
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim vbComp As VBComponent
    For Each vbComp In TargetWorkbook.VBProject.VBComponents
        ReplaceQuoteWithRemInModule vbComp
    Next
End Sub
```

## Custom 

```vb
Sub CommentLines()
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE cpsLineFirst
'@INCLUDE PROCEDURE cpsLineLast
'@INCLUDE PROCEDURE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = cpsLineFirst
    Dim blockEnd As Long
    blockEnd = cpsLineLast
    Dim rowLine As String
    Dim i As Long
    For i = blockStart To blockEnd
        rowLine = Module.CodeModule.Lines(i, 1)
        Module.CodeModule.ReplaceLine i, Space(Len(rowLine) - Len(LTrim(rowLine))) & "'" & Trim(rowLine)
    Next
    Module.CodeModule.CodePane.SetSelection blockStart, 1, blockEnd, 1000
End Sub

Sub UnCommentLines()
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE cpsLineFirst
'@INCLUDE PROCEDURE cpsLineLast
'@INCLUDE PROCEDURE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = cpsLineFirst
    Dim blockEnd As Long
    blockEnd = cpsLineLast
    Dim pos As Long
    Dim i As Long
    For i = blockStart To blockEnd
        With Module.CodeModule
            If left(Trim(.Lines(i, 1)), 1) = "'" Then
                pos = InStr(1, .Lines(i, 1), "'")
                .ReplaceLine i, Replace(.Lines(i, 1), "'", "", , 1)
            End If
        End With
    Next
    Module.CodeModule.CodePane.SetSelection blockStart, 1, blockEnd, 1000
End Sub

Sub RemLines()
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE cpsLineFirst
'@INCLUDE PROCEDURE cpsLineLast
'@INCLUDE PROCEDURE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = cpsLineFirst
    Dim blockEnd As Long
    blockEnd = cpsLineLast
    Dim rowLine As String
    Dim i As Long
    For i = blockStart To blockEnd
        rowLine = Module.CodeModule.Lines(i, 1)
        Module.CodeModule.ReplaceLine i, Space(Len(rowLine) - Len(LTrim(rowLine))) & "Rem " & Trim(Module.CodeModule.Lines(i, 1))
    Next
    Module.CodeModule.CodePane.SetSelection blockStart, 1, blockEnd, 1000
End Sub

Sub UnRemLines()
'@AssignedModule F_Vbe_Comments
'@INCLUDE PROCEDURE cpsLineFirst
'@INCLUDE PROCEDURE cpsLineLast
'@INCLUDE PROCEDURE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = cpsLineFirst
    Dim blockEnd As Long
    blockEnd = cpsLineLast
    Dim i As Long
    For i = blockStart To blockEnd
        With Module.CodeModule
            If left(Trim(.Lines(i, 1)), 4) = "Rem " Then
                pos = InStr(1, .Lines(i, 1), "Rem ")
                .ReplaceLine i, Replace(.Lines(i, 1), "Rem ", "", , 1)
            End If
        End With
    Next
    Module.CodeModule.CodePane.SetSelection blockStart, 1, blockEnd, 1000
End Sub
```



## Afterword

This seems the right place to ask you to leave a comment!