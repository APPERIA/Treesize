' VBScript per elencare i file nella directory corrente e nelle sotto-cartelle e salvarli in ordine dal più grande al più piccolo
' VBScript to list files in the current directory and subdirectories and save them in order from largest to smallest
On Error Resume Next

' Creazione oggetto FileSystemObject
' Create FileSystemObject object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Ottenere il percorso della directory corrente
' Get the current directory path
strCurrentDirectory = objFSO.GetAbsolutePathName(".")

' Sottoprocedura ricorsiva per ottenere tutti i file nelle sotto-cartelle
' Recursive subroutine to get all files in subdirectories
Sub GetFilesInSubfolders(ByVal objFolder, ByVal strParentPath)
    Set colFiles = objFolder.Files

    ' Aggiungere i file e le dimensioni all'oggetto Dictionary
    ' Create Dictionary object to store files and their sizes
    For Each objFile In colFiles
        strFilePath = strParentPath & "\" & objFile.Name
        objDict.Add strFilePath, objFile.Size
    Next

    ' Ricorsivamente ottenere i file nelle sotto-cartelle
    ' Recursive subroutine to get all files in subdirectories
    For Each objSubFolder In objFolder.SubFolders
        GetFilesInSubfolders objSubFolder, strParentPath & "\" & objSubFolder.Name
    Next
End Sub

' Creazione oggetto Dictionary per memorizzare i file e le loro dimensioni
' Create Dictionary object to store files and their sizes
Set objDict = CreateObject("Scripting.Dictionary")

' Ottenere la lista di tutti i file nella directory corrente e nelle sotto-cartelle
' Get the list of all files in the current directory and subdirectories
Set objFolder = objFSO.GetFolder(strCurrentDirectory)
GetFilesInSubfolders objFolder, strCurrentDirectory

' Creazione oggetto FileSystemObject per scrivere nel file di testo
' Create FileSystemObject object to write to the text file
Set objTextFile = objFSO.CreateTextFile(strCurrentDirectory & "\elenco_file.txt")

' Ordinare l'oggetto Dictionary in base alle dimensioni dei file (dal più grande al più piccolo)
' Sort the Dictionary object based on file sizes (from largest to smallest)
Set objDictSorted = CreateObject("Scripting.Dictionary")
For Each key In objDict.Keys
    objDictSorted.Add objDict(key), key
Next
arrKeys = objDictSorted.Keys
arrKeysArray = Array()
For Each item In arrKeys
    ReDim Preserve arrKeysArray(UBound(arrKeysArray) + 1)
    arrKeysArray(UBound(arrKeysArray)) = item
Next
arrKeysArray = BubbleSort(arrKeysArray)

' Funzione Bubble Sort per ordinare l'array
' Bubble Sort function to sort the array
Function BubbleSort(arr)
    For i = 0 To UBound(arr) - 1
        For j = 0 To UBound(arr) - 1 - i
            If arr(j) < arr(j + 1) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next
    Next
    BubbleSort = arr
End Function

' Creazione oggetto FileSystemObject per scrivere nel file di log
' Create FileSystemObject object to write to the log file
Set objLogFile = objFSO.CreateTextFile(strCurrentDirectory & "\error_log.txt")

' Scrivere i file ordinati nel file di testo insieme alle dimensioni e alla cartella di appartenenza
' Write the sorted files to the text file along with their sizes and parent 
For Each key In arrKeysArray
    filePath = objDictSorted(key)
    fileSize = objDict(filePath)
    If fileSize >= 1024 * 1024 * 1024 Then
        ' Dimensione superiore a 1 gigabyte
        ' Size is larger than 1 gigabyte
        fileSize = fileSize / 1024 / 1024 / 1024 ' Conversione da byte a gigabyte
        objTextFile.WriteLine filePath & " (" & FormatNumber(fileSize, 2) & " GB)"
    ElseIf fileSize >= 1024 * 1024 Then
        ' Dimensione compresa tra 1 megabyte e 1 gigabyte
        ' Size is between 1 megabyte and 1 gigabyte
        fileSize = fileSize / 1024 / 1024 ' Conversione da byte a megabyte
        objTextFile.WriteLine filePath & " (" & FormatNumber(fileSize, 2) & " MB)"
    Else
        ' Dimensione inferiore a 1 megabyte
        ' Size is smaller than 1 megabyte
        fileSize = fileSize / 1024 ' Conversione da byte a kilobyte
        objTextFile.WriteLine filePath & " (" & FormatNumber(fileSize, 2) & " KB)"
    End If
Next

' Chiudere il file di testo
' Close the text file
objTextFile.Close

' Rilasciare gli oggetti
' Release objects
Set objTextFile = Nothing
Set objDictSorted = Nothing
Set objDict = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

' Gestione degli errori
' Error handling
If Err.Number <> 0 Then
    ' Scrivere l'errore nel file di log
    objLogFile.WriteLine "Errore: " & Err.Number
    objLogFile.WriteLine "Descrizione: " & Err.Description
    objLogFile.WriteLine "Posizione: " & Err.Source
    objLogFile.WriteLine "Data/ora: " & Now()
End If

' Chiudere il file di log
' Close the log file
objLogFile.Close

' Rilasciare l'oggetto
' Release the object
Set objLogFile = Nothing
