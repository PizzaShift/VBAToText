Attribute VB_Name = "VBAToText"
Sub ExcelExport(path As String, wb As Workbook)
    Dim VBProj As VBProject
    Set VBProj = wb.VBProject
    
    Dim components As VBComponents
    Set components = VBProj.VBComponents
    
    Dim component As VBComponent
    
    For Each component In components
        Dim tp As String
        Dim extension As String
        Dim to_be_exported As Boolean
             
        to_be_exported = True
        
        Select Case component.Type
            Case vbext_ComponentType.vbext_ct_ClassModule
                tp = "class module"
                extension = ".cls"
            Case vbext_ComponentType.vbext_ct_StdModule
                tp = "module"
                extension = ".bas"
            Case vbext_ComponentType.vbext_ct_MSForm
                tp = "form"
                extension = ".frm"
            Case Else
                to_be_exported = False
        End Select
        
        If to_be_exported Then
            Call ExportComponent(tp, path, extension, component)
        Else
            Debug.Print ("Skipping component " & component.name)
        End If
    Next component
End Sub

Sub ExportComponent(tp As String, path As String, extension As String, component As VBComponent)
    Debug.Print ("Exporting " & tp & " " & component.name)
    component.Export (path & "/" & component.name & extension)
End Sub

Sub ImportComponent(tp As String, path As String, file As String, name As String, wb As Workbook)
    Debug.Print ("Importing " & tp & " " & name)
    wb.VBProject.VBComponents.Import (path & file)
End Sub

Sub ExcelImport(path As String)
    Dim wb As Workbook
    Set wb = Application.Workbooks.Add()
    
    Dim file As String
    file = Dir(path)
    
    While (file <> vbNullString)
        Dim tp As String
        Dim name As String
        Dim to_be_imported As Boolean
        
        to_be_imported = True
        
        Select Case Right(file, 4)
            Case ".cls"
                tp = "class module"
                name = Left(file, Len(file) - 4)
            Case ".bas"
                tp = "module"
                name = Left(file, Len(file) - 4)
            Case ".frm"
                tp = "form"
                name = Left(file, Len(file) - 4)
            Case Else
                to_be_imported = False
        End Select
    
        If to_be_imported Then
            Call ImportComponent(tp, path, file, name, wb)
        Else
            Debug.Print ("Skipping file " & file)
        End If
        
        file = Dir
    Wend
End Sub

