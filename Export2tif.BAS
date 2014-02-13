'#Language "WWB-COM"

Sub Main
Debug.Clear
''''''''''''' User Variables ''''''''''''''''''

file_extension	= "srf"
file_directory	= "\B_grid\Znm_srf\"
export_format = "tif"

''''''''''''''''''''''''''''''''''''''''''''''
	Set surf = CreateObject("surfer.application")
	surf.Visible = True 'Progress for each file can be seen in the status bar of the application.

	'Make sure the file extension has no extra . and the data directory has a trailing \
	file_extension	= LCase(Right(file_extension,(Len(file_extension) - InStrRev(file_extension,"."))))
	If  Len(file_directory)-InStrRev(file_directory,"\") <> 0 Then file_directory = file_directory + "\"

	srf_file = Dir( file_directory  + "*." + file_extension)

	On Error GoTo FileError
	While srf_file <> ""
		'Define output grid file directory & name
		exptJPG_file = file_directory + Left(srf_file, Len(srf_file)-(Len(srf_file)-InStrRev(srf_file,".")+1) ) _
			+ "." +export_format

		'Declares Plot As an Object and opens an SRF file
		Dim Plot As Object
		Set Plot = surf.Documents.Open(file_directory + srf_file)
        'Export srf into jpg
		Plot.Export2(FileName:= exptJPG_file, Options:="Defaults =1", FilterId:=export_format)
		Debug.Print srf_file
		srf_file = Dir()
	Wend
    surf.Quit
	Exit Sub

	'Print a meaningful error message
	FileError:
	Debug.Print  "Error:	" + data_file + "						" + Err.Description
	Resume Next

End Sub
