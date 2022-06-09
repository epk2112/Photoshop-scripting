' Copyright 2002-2008.  Adobe Systems, Incorporated.  All rights reserved.
' This script will create a new target document which is 4 inch x 4 inch document,
' select the contents of the source document and copy it to the clipboard,
' and then paste the contents of the clipboard into the target document.
' Notice that the script sets the active document prior to doing the cut
' and paste because these operations only work on the active document.

Option Explicit

Dim appRef
Dim docRef
Dim docRef2
Dim newLayerRef
Dim fileName

Dim strSamples
Dim strVanishingPoint
Dim strLocString
Dim strArg

Set appRef = CreateObject( "Photoshop.Application" )

appRef.BringToFront

' now create a new document and copy and paste.
If appRef.Documents.Count > 0 Then
    Set docRef = appRef.ActiveDocument
Else ' open sample file

	strSamples = "$$$/LocalizedFilenames.xml/SourceDirectoryName/id/Extras/[LOCALE]/[LOCALE]_Samples/value=Samples"
	strArg = Array(strSamples)
	Call getLocString(strSamples)

	strVanishingPoint = "$$$/LocalizedFilenames.xml/SourceFileName/id/Extras/[LOCALE]/[LOCALE]_Samples/Vanishing_Point.psd/value=Vanishing Point.psd"
	strArg = Array(strVanishingPoint)
	Call getLocString(strVanishingPoint)

	fileName = appRef.Path & "\" & strSamples & "\" & strVanishingPoint
	Set docRef = appRef.Open( fileName )
End If

appRef.Preferences.RulerUnits = 2 ' psInches

Set docRef2 = appRef.Documents.Add( 4, 4, 72, "The New Document" )

appRef.ActiveDocument = docRef

docRef.Selection.SelectAll

docRef.Selection.Copy

appRef.ActiveDocument = docRef2

Set newLayerRef = docRef2.Paste

MsgBox "Clipboard Interaction script complete"

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
