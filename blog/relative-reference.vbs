Const ForReading = 1, ForWriting = 2, ForAppending = 8

pub_dir = "public"
levelPaths = Array("","../","../../","../../../")

'Command line argument for websitre url'
if WScript.Arguments.Count = 0 then
    badPath = "http://localhost:1318/"
else
    badPath = WScript.Arguments(0)
end if

'File system object for file manipulation'
Set objFSO = CreateObject("Scripting.FileSystemObject")

objStartFolder =  objFSO.GetAbsolutePathName(".") & "/" & pub_dir


ShowSubfolders objFSO.GetFolder(objStartFolder), 0 'Start off in the public directory, level 0'

Sub ShowSubFolders(Folder, level)
	Set colFiles = Folder.Files
	For Each objFile in colFiles
		if isHTML(objFile.Path) then 'check if file is HTML'
			localize objFile,level 'make local references in file'
		end if
    Next
    For Each Subfolder in Folder.SubFolders
        'Wscript.Echo Subfolder.Path
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        ShowSubFolders Subfolder, level +1 'Recursive call, increment the level'
    Next
End Sub

Sub localize(file, level)
  'Wscript.echo "test"
	'create object file for reading, this is the HTML file hugo created'
	set myFile = objFSO.OpenTextFile(file.Path,ForReading,True)	

	newFile = myFile.ReadAll 'read the html file and store as stirng'
	myFile.Close 'close the file'


	'THe following uses regular expressions to find a match for links to other posts (<a href)'
	Dim regEx 'Regular Expression object'
	Dim colMatches 'will contain matches for links'

  ' Create regular expression.
	Set regEx = New RegExp
	regEx.Pattern = badPath & "(.*?"")" 'regular expression that captures website url and internal link inside
	regEx.IgnoreCase = True
	regEx.global = true

  Set colMatches = regEx.Execute(newFile)   ' Execute search.
  For Each objMatch In colMatches   ' Iterate Matches collection.

  	fullMatch = objMatch.Value 'This is the full match, includes the full  tag'
  	subMatch = objMatch.SubMatches(0) 'This is the link inside of the tag'

  	replaceMatch = replace(fullMatch,badPath,levelPaths(level)) 'puts relative reference inside the tag'

    if(isFile(subMatch)) then 'Checks if this rerference is to a css/js/html file'
      newFile = replace(newFile,fullMatch,replaceMatch)
    elseif right(subMatch,2) = "/""" then 'check if there is an ending /, if not then add it with index.html'     
         newFile = Replace(newFile,fullMatch, left(replaceMatch,len(replaceMatch)-1) & "index.html""")
      else 'There is no ending / so you must add it here'  
         newFile = Replace(newFile,fullMatch, left(replaceMatch,len(replaceMatch)-1) & "/index.html""")
    end if
  Next
  'Take care of home page references'
  bad_index = "href=""" & badPath
  good_index = "href=""" & levelPaths(level) & "index.html"
  newFile = replace(newFile,left(bad_index,len(bad_index)-1),good_index)
  

	set myFile = objFSO.OpenTextFile(file.Path,ForWriting,True)	
	myFile.Write newFile 'write the html file
	myFile.Close

end Sub

Function isHTML (file)
	if right(file,4) = "html" then
		isHTML = True
	else 
		isHTML = false
	end if
End Function

Function isFile(myString)
  Dim regEx, retVal

  ' Create regular expression.
  Set regEx = New RegExp
  regEx.Pattern = ".*\.[a-zA-Z]{3,4}"
  regEx.IgnoreCase = False

  isFile = regEx.Test(myString)

 
end Function
	

	
