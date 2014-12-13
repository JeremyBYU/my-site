Const ForReading = 1, ForWriting = 2, ForAppending = 8

pub_dir = "public"
levelPaths = Array("","../","../../","../../../")

'3 Groups of strings to search for relative path references
'<link> Tags -- This includes css Files
'<script> Tags -- This include .js'
'<a href=..>' - This includes link to other posts 
' THe _g represents the good link I want it to replace it with (deprecates last '/' '
reg_link = "<link href=""/"
reg_link_g = "<link href="""
reg_script = "<script src=""/"
reg_script_g = "<script src="""
'reg_href = "<a href=""/"
'reg_href_g = "<a href="""

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

	'create object file for reading, this is the HTML file hugo created'
	set myFile = objFSO.OpenTextFile(file.Path,ForReading,True)	

	strFile = myFile.ReadAll 'read the html file and store as stirng'
	myFile.Close 'close the file'

	newFile = Replace(strFile,reg_link,reg_link_g & levelPaths(level)) 'This provides relative reference for <link>'
	newFile = Replace(newFile,reg_script,reg_script_g & levelPaths(level)) 'This provides relative reference for scripts'
	'newFile = Replace(newFile,reg_href,reg_href_g & levelPaths(level)) 'This provies relative reference for links to other pages'

	'THe following uses regular expressions to find a match for links to other posts (<a href)'
	Dim regEx 'Regular Expression object'
	Dim colMatches 'will contain matches for links'

  ' Create regular expression.
  	Set regEx = New RegExp
  	regEx.Pattern = "<a.*?href=""((?!http).*?\/?"")" 'regular expression that capture href but excludeds http'
  	regEx.IgnoreCase = True
  	regEx.global = true


    Set colMatches = regEx.Execute(newFile)   ' Execute search.

    For Each objMatch In colMatches   ' Iterate Matches collection.
    	fullMatch = objMatch.Value 'This is the full match, includes the full anchor tag'
    	subMatch = objMatch.SubMatches(0) 'This is the link inside of the anchor tag'

    	replaceSubMatch = levelPaths(level) & Mid(subMatch,2) 'adds relative refrence to the link inside'
    	replaceMatch = replace(fullMatch,subMatch,replaceSubMatch) 'puts relative reference inside the anchortag'

    	'Wscript.echo "fullMatch: " & fullMatch
    	'Wscript.echo "subMatch: " & subMatch
    	'Wscript.echo "replacesubMatch: " & replaceSubMatch
    	'Wscript.echo "New Replace: " & replaceMatch
    	'Wscript.Echo right(myMatch,2)
    	'If all files are index.html instead of 'slug'.html then you have to manually add the index.html reference
    	if right(replaceMatch,2) = "/""" then 'check if there is an ending /, if not then add it with index.html'
			newFile = Replace(newFile,fullMatch, left(replaceMatch,len(replaceMatch)-1) & "index.html""")
    	elseif len(subMatch) <= 1 then 'This is only true for the hompage, it captures the link that is empty'
    		replaceMatch = Replace(fullMatch,"href=""""","href=""" & levelPaths(level) & "index.html""")
    		newFile = Replace(newFile,fullMatch,replaceMatch)
    	else 'There is no ending / so you must add it here'    		
    		newFile = Replace(newFile,fullMatch, left(replaceMatch,len(replaceMatch)-1) & "/index.html""")
    	end if
  		
      'RetStr = RetStr & "Match found at position "
      'RetStr = RetStr & objMatch.FirstIndex & ". Match Value is '"
      'RetStr = RetStr & objMatch.Value & "'." & vbCrLf
    Next
    'Wscript.Echo RetStr

    'Overwrite the old html file with the new one'
	set myFile = objFSO.OpenTextFile(file.Path,ForWriting,True)	
	myFile.Write newFile 'write the html file
	myFile.Close

end Sub

Function isHTML (file)
	if right(file,4) = "html" then
		isHTML = True
		'Wscript.Echo "true"
	else 
		checkHTML = false
	end if
End Function
	

	
