Set objExplorer = CreateObject("InternetExplorer.Application")
With objExplorer
    .Navigate "about:blank"
    .Visible = 1
    .Document.Title = "You bozo"
    .Toolbar = False
    .Statusbar = False
    .Top = 100
    .Left = 400
    .Height = 200
    .Width = 20
    
    ' Retrieve a random word from the API
    Dim randomWord
    randomWordURL = "https://random-word-api.herokuapp.com/word"
    
    Dim xhr
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    
    xhr.Open "GET", randomWordURL, False
    xhr.send
    
    If xhr.Status = 200 Then
        randomWord = xhr.responseText
    End If
    
    'Get a list of files in the directory
    Dim fso, folder, files, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder("C:\Users\damian.swan\Documents\PhotoWheel\")
    Set files = folder.Files
    
    'Pick a random file
    Randomize
    Dim fileIndex
    fileIndex = Int(files.Count * Rnd)
    Set file = Nothing
    For Each f In files
        If fileIndex = 0 Then
            Set file = f
            Exit For
        End If
        fileIndex = fileIndex - 1
    Next
    
    'Set the image source and random word display
    .Document.Body.innerHTML = "<p>" & randomWord & "</p>"
    .Document.Body.innerHTML = .Document.Body.innerHTML & "<img src='" & file.Path & "'>"
End With
