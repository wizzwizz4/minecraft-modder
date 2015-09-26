dim mcPath, modPath, unmodPath, fsobj, modFolder, unmodFolder, wshobj, appdata

set wshobj = CreateObject("WScript.Shell")
appdata = wshobj.ExpandEnvironmentStrings("%appdata%")

mcPath = appdata & "\.minecraft"
modPath = mcPath & "\mods"
unmodPath = mcPath & "\.unmods"

set fsobj = CreateObject("Scripting.FileSystemObject")
modFolder = fsobj.GetFolder(modPath)
unmodFolder = fsobj.GetFolder(unmodPath)

Function EnableMod(modFile)
	fsobj.MoveFile unmodPath & modFile, modPath
End Function

Function DisableMod(modFile)
	fsobj.MoveFile modPath & modFile, unmodPath
End Function

Function getMods()
	redim returnarr(CInt(modFolder.Files.count))
	dim loopCounter
	loopCounter = 0
	For each g in modFolder.Files
		returnarr(loopCounter) = g
		loopCounter = loopCounter + 1
	Next
	getMods = returnarr
End Function

Function getUnMods()
	redim returnarr(CInt(unmodFolder.Files.count))
	dim loopCounter
	loopCounter = 0
	For each g in unmodFolder.Files
		returnarr(loopCounter) = g
		loopCounter = loopCounter + 1
	Next
	getUnMods = returnarr
End Function