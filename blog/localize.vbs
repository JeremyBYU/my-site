Const ForReading = 1, ForWriting = 2, ForAppending = 8
dim pub_dir
Dim levelPaths
pub_dir = "public"
levelPaths = Array("","../","../../")

reg_link = "<link href=""/"
reg_link_g = "<link href="""
reg_script = "<script src=""/"
reg_script_g = "<script src="""
reg_href = "<a href=""/"
reg_href_g = "<a href="""

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder =  objFSO.GetAbsolutePathName(".") & "/" & pub_dir






ShowSubfolders objFSO.GetFolder(objStartFolder), 0

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
        ShowSubFolders Subfolder, level +1
    Next
End Sub

Sub localize(file, level)

	'create temportary file, this will eventually become the final file'
	set tmpFile = objFSO.OpenTextFile(fileAddOn(file.Path,"_tmp"),ForWriting,True)
	'create object for reading, this is the HTML file hugo created'
	set myFile = objFSO.OpenTextFile(file.Path,ForReading,True)	

	strFile = myFile.ReadAll 'read the html file and store as stirng'
	myFile.Close

	newFile = Replace(strFile,reg_link,reg_link_g & levelPaths(level))
	newFile = Replace(newFile,reg_script,reg_script_g & levelPaths(level))
	newFile = Replace(newFile,reg_href,reg_href_g & levelPaths(level))


	Dim regEx, str1
	Dim colMatches
  	str1 = "The quick brown fox jumps over the lazy dog."

  ' Create regular expression.
  	Set regEx = New RegExp
  	regEx.Pattern = "(a href=""(?!http).*\/"")"
  	regEx.IgnoreCase = True
  	regEx.global = true

  ' Make replacement.
  	   'Get the matches.
    Set colMatches = regEx.Execute(newFile)   ' Execute search.

    For Each objMatch In colMatches   ' Iterate Matches collection.
      newFile = Replace(newFile,objMatch.Value,left(objMatch.Value,len(objMatch.value)-1) & "index.html""")
      RetStr = RetStr & "Match found at position "
      RetStr = RetStr & objMatch.FirstIndex & ". Match Value is '"
      RetStr = RetStr & objMatch.Value & "'." & vbCrLf
    Next
    Wscript.Echo RetStr	

	myFile.Write newFile 
	myFile.Close


end Sub

function fileAddOn(file,toAdd)
	position = InStrRev(file,".")
	result = Left(file, position -1) & toAdd & Mid( file, position)	
	fileAddOn = result
end Function


Function isHTML (file)
	if right(file,4) = "html" AND InStr(file,"_tmp") = 0 then
		isHTML = True
		'Wscript.Echo "true"
	else 
		checkHTML = false
	end if
End Function
	

	
