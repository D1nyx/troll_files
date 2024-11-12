' Ao executar, este código cria uma pasta "AposentadorDeDiscos" dentro do "C:\Temp\" que no interior estaria a gerar pastas infinitas- Para parar reinicie o PC
Dim fileSystem, rootFolder, i
Set fileSystem = CreateObject("Scripting.FileSystemObject")

rootFolder = "C:\Temp\AposentadorDeDiscos"

If Not fileSystem.FolderExists(rootFolder) Then
    fileSystem.CreateFolder(rootFolder)
End If

For i = 1 To 1000000
    Dim novaPasta
    novaPasta = rootFolder & "\Pasta_" & i
    
    fileSystem.CreateFolder(novaPasta)
    
    Dim j
    For j = 1 To 100
        fileSystem.CreateFolder(novaPasta & "\Subpasta_" & j)
    Next
Next
