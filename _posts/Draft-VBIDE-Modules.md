---
title: VBIDE - Modules
# author: 'Anastasiou Alex'
# date: 2022-12-02 12:00:00 # if missing it's taken from the filename
# last_modified_at: #2022-12-02 12:00:00 
categories: [VBIDE, Modules] # can handle 1 category and 1 subcategory eg [Category, Subcategory]
# tags: [Sort] # [Tag1, Tag2 ...]
---

# Macros

```vb

Function ModuleAddOrSet( _
                       TargetWorkbook As Workbook, _
                       TargetName As String, _
                       ModuleType As VBIDE.vbext_ComponentType) As VBComponent
'@AssignedModule F_Vbe_Modules
'@INCLUDE PROCEDURE ActiveCodepaneWorkbook

'Example
'Dim Module as vbComponent
'set Module=ModuleAddOrSet(TargetWorkbook,"NewModule",vbext_ct_StdModule)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    On Error Resume Next
    Set Module = TargetWorkbook.VBProject.VBComponents(TargetName)
    On Error GoTo 0
    If Module Is Nothing Then
        Set Module = TargetWorkbook.VBProject.VBComponents.Add(ModuleType)
        Module.name = TargetName
    End If
    Set ModuleAddOrSet = Module
End Function


Sub GoToModule(Module As VBComponent)
'@AssignedModule F_Vbe_Modules
    With Application.VBE.MainWindow
        .visible = True
        .WindowState = vbext_ws_Maximize
    End With
    With Module.CodeModule.CodePane
        .Show
        .Window.visible = True
        .Window.WindowState = vbext_ws_Maximize
        .Window.SetFocus
        .SetSelection 1, 1, 1, 1
    End With
End Sub

Function ModuleCopy( _
                   Module As VBComponent, _
                   TargetWorkbook As Workbook, _
                   OverwriteExisting As Boolean) As Boolean
'@INCLUDE PROCEDURE WorkbookOfModule
'@INCLUDE PROCEDURE ModuleExists
'@INCLUDE PROCEDURE ModuleExtension
'@AssignedModule F_Vbe_Modules
    If Module.name = "ThisWorkbook" Then Exit Function
    If Module.Type = vbext_ct_Document Then Exit Function
    If WorkbookOfModule(Module).name = TargetWorkbook.name Then Exit Function
    Dim TempModule As VBIDE.VBComponent
    
    If ModuleExists(Module.name, TargetWorkbook) Then
        If OverwriteExisting = True Then
            With TargetWorkbook.VBProject
                .VBComponents.Remove .VBComponents(Module.name)
            End With
        Else
            Exit Function
        End If
    End If
    
    Dim ext As String
        ext = ModuleExtension(Module)
    Dim FName As String
        FName = Environ("Temp") & "\" & Module.name & ext
    Module.Export fileName:=FName
    
    TargetWorkbook.VBProject.VBComponents.Import fileName:=FName
    Kill FName
    ModuleCopy = True
End Function


Sub ModuleRemove( _
                TargetWorkbook As Workbook, _
                Module As VBComponent)
'@INCLUDE PROCEDURE GetSheetByCodeName
'@INCLUDE PROCEDURE WorkbookOfModule
'@INCLUDE PROCEDURE ModuleIgnore
'@AssignedModule F_Vbe_Modules
    If ModuleIgnore(Module) Then Exit Sub
    Application.DisplayAlerts = False
    If Module.Type = vbext_ct_Document Then
        If Module.name = "ThisWorkbook" Then
            Module.CodeModule.DeleteLines 1, Module.CodeModule.CountOfLines
        Else
            If TargetWorkbook.SHEETS.Count > 1 Then
                GetSheetByCodeName(TargetWorkbook, Module.name).Delete
            Else
                Dim TaragetWorksheet As Worksheet
                Set TaragetWorksheet = TargetWorkbook.SHEETS.Add
                TaragetWorksheet.name = "LastSheet"
                GetSheetByCodeName(TargetWorkbook, Module.name).Delete
            End If
        End If
    Else
        WorkbookOfModule(Module).VBProject.VBComponents.Remove Module
    End If
    Application.DisplayAlerts = True
End Sub


Sub ModuleRemoveEmpty(Optional TargetWorkbook As Workbook)
'@INCLUDE PROCEDURE ActiveCodepaneWorkbook
'@INCLUDE PROCEDURE ProceduresOfModule
'@INCLUDE PROCEDURE ModuleIgnore
'@AssignedModule F_Vbe_Modules
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Type = vbext_ct_StdModule Then
            If Not ModuleIgnore(Module) Then
                If ProceduresOfModule(Module).Count = 0 And Module.CodeModule.CountOfLines < 3 Then TargetWorkbook.VBProject.VBComponents.Remove Module
            End If
        End If
    Next
End Sub


Sub ModuleCodeRemove(Module As VBComponent)
'@INCLUDE PROCEDURE ModuleIgnore
'@AssignedModule F_Vbe_Modules
    If ModuleIgnore(Module) Then Exit Sub
    If Module.CodeModule.CountOfLines = 0 Then Exit Sub
    Module.CodeModule.DeleteLines 1, Module.CodeModule.CountOfLines '+ 1
End Sub


Sub ModuleCodeMove( _
                  FromModule As VBComponent, _
                  TargetModule As VBComponent)
'@INCLUDE PROCEDURE ModuleCode
'@INCLUDE PROCEDURE ModuleIgnore
'@INCLUDE PROCEDURE ModuleCodeRemove
'@AssignedModule F_Vbe_Modules
    If ModuleIgnore(FromModule) Then Exit Sub
    Dim ModuleDeclarations As String
    Dim ModuleCode As String
    Dim counter As Long
    If FromModule.CodeModule.CountOfDeclarationLines > 0 Then
        For counter = 1 To FromModule.CodeModule.CountOfDeclarationLines
            ModuleDeclarations = ModuleDeclarations & vbNewLine & FromModule.CodeModule.Lines(counter, 1)
        Next
    End If
    If FromModule.CodeModule.CountOfLines - FromModule.CodeModule.CountOfDeclarationLines > 0 Then
        For counter = FromModule.CodeModule.CountOfDeclarationLines + 1 To FromModule.CodeModule.CountOfLines
            ModuleCode = ModuleCode & vbNewLine & FromModule.CodeModule.Lines(counter, 1)
        Next
    End If
    With TargetModule.CodeModule
        .InsertLines 1, ModuleDeclarations
        .InsertLines .CountOfLines + 1, ModuleCode
    End With
    ModuleCodeRemove FromModule
End Sub


Sub ModulesMerge( _
               TargetWorkbook As Workbook, _
               TargetModule As VBComponent)
'@INCLUDE PROCEDURE ModuleCodeMove
'@AssignedModule F_Vbe_Modules
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Type = vbext_ct_StdModule Then
            If Module.name <> TargetModule.name Then
                ModuleCodeMove Module, TargetModule
            End If
        End If
    Next
End Sub



Sub ModuleImport( _
                TargetWorkbook As Workbook, _
                ReplaceExisting As Boolean)
'@AssignedModule F_Vbe_Modules
'@INCLUDE PROCEDURE ArrayAllocated
'@INCLUDE PROCEDURE DataFilePartFolder
'@INCLUDE PROCEDURE DataFilePartExtension
'@INCLUDE PROCEDURE DataFilePartName
'@INCLUDE PROCEDURE DataFilePicker
'@INCLUDE PROCEDURE WorkbookExists
'@INCLUDE PROCEDURE WorksheetExists
'@INCLUDE PROCEDURE Toast
'@INCLUDE PROCEDURE ModuleExists
'@INCLUDE PROCEDURE ModuleRemove
    Dim SelectedModules
        SelectedModules = DataFilePicker(Array("bas", "frm", "cls", "doccls"), True)
    If Not ArrayAllocated(SelectedModules) Then Exit Sub

    Dim basePath As String
        basePath = DataFilePartFolder(SelectedModules(1), True)
    Dim SourceWorkbook As Workbook
    Dim SourceWorkbookName As String
        SourceWorkbookName = Dir(basePath & "*.xl*")
    Dim wasOpen As Boolean
    If SourceWorkbookName <> "" Then
        wasOpen = WorkbookExists(SourceWorkbookName)
    End If
    Dim extension As String
    Dim TargetName As String
    Dim element
    For Each element In SelectedModules
        TargetName = DataFilePartName(CStr(element), False)
        extension = DataFilePartExtension(CStr(element))
        If extension <> "doccls" Then
            If ModuleExists(TargetName, TargetWorkbook) Then
                If ReplaceExisting Then
                    ModuleRemove TargetWorkbook, TargetWorkbook.VBProject.VBComponents(TargetName)
                Else
                    GoTo NextElement
                End If
            End If
            TargetWorkbook.VBProject.VBComponents.Import element
        ElseIf extension = "doccls" And SourceWorkbookName <> "" Then
            If WorksheetExists(TargetName, TargetWorkbook) Then
                If ReplaceExisting Then
                    TargetWorkbook.Worksheets(TargetName).name=TargetName & "_old"
                    'TargetWorkbook.Worksheets(TargetName).Delete
                Else
                    GoTo NextElement
                End If

                If wasOpen = False Then
                    Application.EnableEvents = False
                    Set SourceWorkbook = Workbooks.Open(basePath & SourceWorkbookName)
                Else
                    Set SourceWorkbook = Workbooks(SourceWorkbookName)
                End If
                SourceWorkbook.SHEETS(TargetName).Copy Before:=TargetWorkbook.SHEETS(1)
                Application.EnableEvents = True
            End If
        End If
NextElement:
    Next element

    If wasOpen = False And WorkbookExists(SourceWorkbookName) Then SourceWorkbook.Close False
    'https://github.com/rfl808/Notify
    Toast , "Import successful"
End Sub


```


# Helpers

```vb

Function DataFilePartFolder(fileNameWithExtension, Optional IncludeSlash As Boolean) As String
    DataFilePartFolder = left(fileNameWithExtension, InStrRev(fileNameWithExtension, "\") - 1 - IncludeSlash)
End Function

Function DataFilePartName(fileNameWithExtension As String, Optional IncludeExtension As Boolean) As String
    If InStr(1, fileNameWithExtension, "\") > 0 Then
        DataFilePartName = Right(fileNameWithExtension, Len(fileNameWithExtension) - InStrRev(fileNameWithExtension, "\"))
    ElseIf InStr(1, fileNameWithExtension, "/") > 0 Then
        DataFilePartName = Right(fileNameWithExtension, Len(fileNameWithExtension) - InStrRev(fileNameWithExtension, "/"))
    Else
        DataFilePartName = fileNameWithExtension
    End If
    If IncludeExtension = False Then DataFilePartName = left(DataFilePartName, InStr(1, DataFilePartName, ".") - 1)
End Function

Function DataFilePartExtension(str As String)
    DataFilePartNameExtension = Mid(str, InStrRev(str, ".") + 1)
End Function


Public Function DataFilePicker(Optional fileType As Variant, Optional multiSelect As Boolean) As Variant
'@AssignedModule F_FileFolder
    Dim blArray As Boolean
    Dim i As Long
    Dim strErrMsg As String, strTitle As String
    Dim varItem As Variant
    If Not IsMissing(fileType) Then
        blArray = IsArray(fileType)
        If Not blArray Then strErrMsg = "Please pass an array in the first parameter of this function!"
    End If
    If strErrMsg = vbNullString Then
        If multiSelect Then strTitle = "Choose one or more files" Else strTitle = "Choose file"
        With Application.FileDialog(msoFileDialogFilePicker)
            .initialFileName = Environ("USERprofile") & "\Desktop\"
            .AllowMultiSelect = multiSelect
            .Filters.clear
            If blArray Then .Filters.Add "File type", Replace("*." & Join(fileType, ", *."), "..", ".")
            .title = strTitle
            If .Show <> 0 Then
                ReDim arrResults(1 To .SelectedItems.Count) As Variant
                If blArray Then
                    For Each varItem In .SelectedItems
                        i = i + 1
                        arrResults(i) = varItem
                    Next varItem
                Else
                    arrResults(1) = .SelectedItems(1)
                End If
                DataFilePicker = arrResults
            End If
        End With
    Else
        MsgBox strErrMsg, vbCritical, "Error!"
    End If
End Function


Public Function ArrayAllocated(ByVal arr As Variant) As Boolean
 '@AssignedModule F_Arrays   
    On Error Resume Next
    ArrayAllocated = IsArray(arr) And (Not IsError(LBound(arr, 1))) And LBound(arr, 1) <= UBound(arr, 1)
End Function

Function WorksheetExists(SheetName As String, TargetWorkbook As Workbook) As Boolean
    Dim TargetWorksheet  As Worksheet
    On Error Resume Next
    Set TargetWorksheet = TargetWorkbook.SHEETS(SheetName)
    On Error GoTo 0
    WorksheetExists = Not TargetWorksheet Is Nothing
End Function


```