' This grids All of the specified Type of data file In the specified directory.

Sub Main
Debug.Clear
'''''''''''''  Variables ''''''''''''''''''

file_extension	= "xls"

file_directory	= "..\ELJ_forpaper\xls\"
file_directory_grid	= "..\ELJ_forpaper\xls\grid\"

BLN_file = file_directory + "blnk4paper.bln"

'''''''''''''''''''''''''''''''''''''''''''''''
	Set surf = CreateObject("surfer.application")
	surf.Visible = True 'Progress for each file can be seen in the status bar of the application.

	'Make sure the file extension has no extra . and the data directory has a trailing \
	file_extension	= LCase(Right(file_extension,(Len(file_extension) - InStrRev(file_extension,"."))))
	If  Len(file_directory)-InStrRev(file_directory,"\") <> 0 Then file_directory = file_directory + "\"

	data_file = Dir( file_directory  + "*." + file_extension)

	On Error GoTo FileError
	While data_file <> ""
		'Define output grid file directory & name
		grid_file	= file_directory_grid + Left(data_file, Len(data_file)-(Len(data_file)-InStrRev(data_file,".")+1) ) + ".grd"

		'Grid the data file with the current Surfer defaults (but do not fill the screen with grid reports)
		surf.GridData (DataFile:= file_directory + data_file, _
		          xCol:=1, yCol:=2, zCol:=3, Algorithm:=srfKriging, _
		          xMin:=0, xMax:=1, yMin:=0, yMax:=1, _
		          ShowReport:=False, OutGrid:=grid_file)

		blanked_file  = file_directory_grid + "Blnkd_" + Left(data_file, Len(data_file)-(Len(data_file)-InStrRev(data_file,".")+1) ) + ".grd"

		surf.GridBlank(grid_file, BLN_file, blanked_file)

		data_file = Dir() 'get next file
	Wend

	surf.Quit
	Exit Sub

	'Print a meaningful error message for each file that did not grid correctly
	FileError:
	Debug.Print  "Error:	" + data_file + "						" + Err.Description
	Resume Next
End Sub
'
