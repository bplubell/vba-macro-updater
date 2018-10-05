Attribute VB_Name = "Entry"

Const sharedCodeFolderName = "Shared"

Sub main()
    Dim currentMacroPath As String
    currentMacroPath = Application.VBE.ActiveVBProject.FileName
    Call DeleteModules(currentMacroPath)
    Pause
    Call AddModules(currentMacroPath)
End Sub

Private Sub AddModules(ByVal currentMacroPath As String)
    Dim projectName As String
    Dim repositoryPath As String
    Dim projects() As String
    projects = GetProjectFolderNames(currentMacroPath)
    repositoryPath = GetFolderFromPath(currentMacroPath)
    For Each project In Application.VBE.VBProjects
        projectName = Replace(project.Name,"Project", "")
        If ArrayContains(projectName, projects) Then
            Call AddModulesInFolder(repositoryPath & "\" & projectName & "\", project)
            Call AddModulesInFolder(repositoryPath & "\" & sharedCodeFolderName & "\", project)
        End If
    Next
End Sub

Private Sub DeleteModules(ByVal currentMacroPath As String)
    Dim projects() As String
    projects = GetProjectFolderNames(currentMacroPath)
    For Each project In Application.VBE.VBProjects
        If ArrayContains(Replace(project.Name,"Project", ""), projects) Then
            For Each module In project.VBComponents
                If module.Name <> "ThisLibrary" Then
                    Call project.VBComponents.Remove(project.VBComponents(module.Name))
                End If
            Next
        End If
    Next
End Sub

' Helper functions
Private Function AddModulesInFolder(ByVal Path As String, project As Variant) As Integer
    Dim moduleName As String
    Dim output As Integer
    output = 0
    moduleName = Dir(Path & "*")
    Do While Trim(moduleName) <> ""
        If Right(moduleName,4) = ".cls" _
            Or Right(moduleName,4) = ".bas" _
            Or Right(moduleName,4) = ".frm" _
            Then
            
            Call project.VBComponents.Import(Path & "\" & moduleName)
        End If
        moduleName = Dir()
    Loop
    AddModulesInFolder = output
End Function

Private Function ArrayContains(ByVal text As String, list() As String) As Boolean
    Dim output As Boolean
    output = False
    For i = LBound(list) To UBound(list)
        If (text = list(i)) Then
            output = True
        End If
    Next i
    ArrayContains = output
End Function

Private Function GetFilenameFromPath(ByVal path As String, ByVal includeExtension As Boolean) As String
    Dim splitter() As String
    Dim name As String
    
    If (path <> "") Then
        splitter = Split(path, "\")
        name = splitter(UBound(splitter))
        If includeExtension = False Then
            Erase splitter
            splitter = Split(name, ".")
            If UBound(splitter) - LBound(splitter) <> 0 Then
                ReDim Preserve splitter(UBound(splitter) - 1)
                name = Join(splitter, ".")
            End If
        End If
    Else
        name = ""
    End If
    
    GetFilenameFromPath = name
End Function

Private Function GetFolderFromPath(ByVal path As String) As String
    Dim output As String
    Dim splitter() As String
    output = ""
    If Trim(path) <> "" And InStr(1, path, "\", vbTextCompare) <> 0 Then
        splitter = Split(path, "\")
        For i = LBound(splitter) To UBound(splitter) - 1 Step 1
            If i = LBound(splitter) Then
                output = splitter(i)
            Else
                output = output & "\" & splitter(i)
            End If
        Next i
    End If
    GetFolderFromPath = output
End Function

Private Function GetProjectFolderNames(ByVal currentProjectPath As String) As String()
    Dim output() As String
    Dim currentFolder As String
    Dim currentProjectName As String
    Dim projectFolder As String
    Dim index As Integer
    index = 0
    currentProjectName = GetFilenameFromPath(currentProjectPath, False)
    currentFolder = GetFolderFromPath(currentProjectPath)
    projectFolder = Dir(currentFolder & "\", vbDirectory)
    Do While Trim(projectFolder) <> ""
        If GetAttr(currentFolder & "\" & projectFolder) = vbDirectory _
            And InStr(1, projectFolder, "Common", vbTextCompare) = 0 _
            And Trim(Replace(projectFolder, ".", "")) <> "" _
            And InStr(1, projectFolder, currentProjectName, vbTextCompare) = 0 _
            Then
            
            ReDim Preserve output(0 To index) As String
            output(UBound(output)) = projectFolder
            index = index + 1
        End If
        projectFolder = Dir()
    Loop
    GetProjectFolderNames = output
End Function

Private Function Pause()
    Dim start As Single
    start = Timer
    Do While Timer < start + 0.5
        DoEvents
    Loop
End Function
