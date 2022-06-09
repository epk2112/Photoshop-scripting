' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' This script creates an action, which is equivalent to the Mosaic Tiles action
' and executes it.

Option Explicit

Dim appRef
Dim filterDescriptor
Dim retDescriptor
Dim keyTileSizeID
Dim keyGroutWidthID
Dim keyLightenGroutID
Dim eventMosaicID
Dim adesc
Dim actionRef

Dim strSamples
Dim strVanishingPoint
Dim strLocString
Dim strArg

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

If appRef.Documents.Count <= 0 Then
	Dim fileName

	strSamples = "$$$/LocalizedFilenames.xml/SourceDirectoryName/id/Extras/[LOCALE]/[LOCALE]_Samples/value=Samples"
	strArg = Array(strSamples)
	Call getLocString(strSamples)

	strVanishingPoint = "$$$/LocalizedFilenames.xml/SourceFileName/id/Extras/[LOCALE]/[LOCALE]_Samples/Vanishing_Point.psd/value=Vanishing Point.psd"
	strArg = Array(strVanishingPoint)
	Call getLocString(strVanishingPoint)

	fileName = appRef.Path & "\" & strSamples & "\" & strVanishingPoint
	appRef.Open ( fileName )
End If

' create an action and execute it.
keyTileSizeID = appRef.CharIDToTypeID( "TlSz" )
keyGroutWidthID = appRef.CharIDToTypeID( "GrtW" )
keyLightenGroutID = appRef.CharIDToTypeID( "LghG" )
eventMosaicID = appRef.CharIDToTypeID( "MscT" )

Set filterDescriptor = CreateObject( "Photoshop.ActionDescriptor" )
filterDescriptor.PutInteger keyTileSizeID, 12
filterDescriptor.PutInteger keyGroutWidthID, 3
filterDescriptor.PutInteger keyLightenGroutID, 9

Set retDescriptor = appRef.ExecuteAction( eventMosaicID, filterDescriptor, 3 ) ' 3 = psDisplayNoDialogs

MsgBox "Create And Execute Action complete"

' ===============================================
' getLocString functions
' ===============================================
' on localized builds we pull the $$$/Strings from a .dat file, see documentation for more details
Function getLocString(strLocString)

	Dim objWshShell
	Dim myScriptPath
	Dim myFSO
	Dim strJSXFile

	Set objWshShell = WScript.CreateObject("WScript.Shell")
	myScriptPath = objWshShell.CurrentDirectory
	Set myFSO = CreateObject("Scripting.FileSystemObject")
	strJSXFile = myScriptPath + "\localize.jsx"

	strLocString =  appRef.DoJavaScriptFile(strJSXFile,Array(strLocString),1)

End Function
