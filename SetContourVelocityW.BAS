'#Language "WWB-COM"
Sub Main
	Dim SurferApp As Object
	Set SurferApp = CreateObject("Surfer.Application")
	SurferApp.Visible = True
	Dim Plot As Object
	Set Plot = SurferApp.Documents.Add
	Dim Shapes As Object
	Set Shapes = Plot.Shapes
	Dim MapFrame(9) As Object

'Setup the page
	With Plot.PageSetup
		.LeftMargin = 0.2
        .RightMargin = 0.2
	    .TopMargin = 0.2
	    .BottomMargin = 0.2
	    .Height = 9.5
	    .Width = 8.8
	End With

'Read grid file, generate contourmaps
	file_directory	= "..\ELJ_data\Automation_SrfScript\VelocityW\"
	file_extension	= "grd"

	file_extension	= LCase(Right(file_extension,(Len(file_extension) - InStrRev(file_extension,"."))))
	If  Len(file_directory)-InStrRev(file_directory,"\") <> 0 Then file_directory = file_directory + "\"

	data_file = Dir( file_directory  + "*." + file_extension)

'Set up Map's Left and Top
    Set DsMapY = 3.2
    Set UsMapY = 5.4    'WsY = UsMapY + 3.0
    Set TOneX = 0.6
    Set TTwoX = 5.0
'Set Plot Title
	Set TextTitle = Shapes.AddText(x:=TTwoX + 3.2, y:=UsMapY + 3.5, _
		Text:=" W Velocity Component")
	TextTitle.Font.Size = 24
	TextTitle.Font.Bold = True
	TextTitle.Font.Italic = True
	TextTitle.Font.HAlign = srfTARight

'Contour Grid files
    Set i =0
	While data_file <>""
		i =i +1
		Set MapFrame(i) = Shapes.AddContourMap(file_directory + data_file)
        Set Contour1 = MapFrame(i).Overlays("Contours")
        'Set Level1 = Contour1.Levels
        With MapFrame(i)
        	.xLength = 3
        	.yLength = 2
        End With
'Do not show level labels
        'Dim Level As Object
        'Set Level = Contour1.Levels
        'Level.ShowLabel = False
        Set ContourLevels = Contour1.Levels
        ContourLevels.AutoGenerate(MinLevel:=-0.3, MaxLevel:=1.5,Interval:=0.3)
'Fill contour maps with Geology CLR
        Contour1.FillContours = True
        Contour1.FillForegroundColorMap.LoadFile(SurferApp.Path + "\ColorScales\Geology.clr")
        Contour1.FillForegroundColorMap.SetDataLimits(-0.3,1.5)
        Contour1.ApplyFillToLevels(1,1,0)
'Show one color scale
        If i <> 1 Then
        	Contour1.ShowColorScale = False
        	Else
        		Contour1.ShowColorScale = True
        		Set Scl =Contour1.ColorScale
        		Scl.LabelFrequency = 1
        		Scl.FirstLabel = 1
        		Scl.Rotation = -90
        		Scl.LabelAngle = 90
        		Scl.Top = UsMapY +2.0
        		Scl.Left = TTwoX -0.8
        End If
'Arrange maps show Left and bottom labels and titles
        If i = 1 Or i = 3 Or i =5 Then
        		MapFrame(i).Left = TOneX
    	        MapFrame(i).Axes("Left axis").ShowLabels = True
    	        MapFrame(i).Axes("Left axis").SetScale(MajorInterval:= 0.2)
        		MapFrame(i).Axes("Left Axis").Title = "y/d"
    	        MapFrame(i).Axes("Left Axis").TitleFont.Size = 18
    	        Else
        			MapFrame(i).Left = TTwoX
        			MapFrame(i).Axes("Left axis").ShowLabels = False

        End If
        If i = 1 Or i = 2 Then
        	MapFrame(i).Top = DsMapY
        	MapFrame(i).Axes("Bottom axis").ShowLabels = True
        	MapFrame(i).Axes("Bottom axis").Title = "x/B"
        	MapFrame(i).Axes("Bottom Axis").TitleFont.Size = 16
        	ElseIf i = 3 Or i =4 Then
        		MapFrame(i).Top = UsMapY
        		MapFrame(i).Axes("Bottom axis").ShowLabels = False
        		Else
        			MapFrame(i).Top = UsMapY + 3.0
        			MapFrame(i).Axes("Bottom axis").ShowLabels = False
        End If
        data_file = Dir()
        'i = i+1
	Wend
'Add lines, texts to indicate the location of ELJs
    Set TextWs = Shapes.AddText(x:=TOneX + 1.6, y:=UsMapY +3.25,Text:="Without Structure")
        	'TextWs.Font.ForeColorRGBA.Color = srfColorBlue
        	TextWs.Font.Size = 15
        	TextWs.Font.Bold = True
        	TextWs.Font.Italic = True
    Set TextDs = Shapes.AddText(x:=TOneX + 3.4, y:=DsMapY - 0.8,Text:="Downstream")
        	'TextDs.Font.ForeColorRGBA.Color = srfColorBlue
        	TextDs.Font.Size = 15
        	TextDs.Font.Bold = True
        	TextDs.Font.Italic = True
    Set TextUs = Shapes.AddText(x:=TOneX + 3.5, y:=UsMapY - 0.8,Text:="Upstream")
        	'TextUs.Font.ForeColorRGBA.Color = srfColorBlue
        	TextUs.Font.Size = 15
        	TextUs.Font.Bold = True
        	TextUs.Font.Italic = True

'indicate Type One Location  ---LEFT
    Set Line3 = Shapes.AddLine(xBeg:=TOneX + 0.28, yBeg:=UsMapY - 0.07, _
    	xEnd:=TOneX + 0.28, yEnd:=UsMapY + 0.45)
    Line3.Line.Width = 0.03                    'short line LEFT
    Set Line1 = Shapes.AddLine(xBeg:=TOneX + 0.90, yBeg:=UsMapY - 4.288, _
    	xEnd:=TOneX + 0.90, yEnd:=UsMapY + 0.15)
    Line1.Line.Width = 0.03                     'long line LEFT
    Set Text1 = Shapes.AddText(x:= TOneX + 0.5, y:= UsMapY + 0.52, _
    	Text:="ELJ1")
    Text1.Font.Size = 21
    Text1.Font.Bold = True
    Text1.Font.Italic = True
    Set Line5 = Shapes.AddLine(xBeg:=TOneX + 0.43, yBeg:=UsMapY + 0.15, _
    	xEnd:=TOneX + 0.28, yEnd:=UsMapY + 0.15)
    Line5.Line.Width = 0.015
    Line5.EndArrow = srfASSimple              'T1 Arrow LEFT
    Set Line6 = Shapes.AddLine(xBeg:=TOneX + 0.75, yBeg:=UsMapY + 0.15, _
    	xEnd:=TOneX + 0.90, yEnd:=UsMapY + 0.15)
    Line6.Line.Width = 0.015
    Line6.EndArrow = srfASSimple               'T1 Arrow RIGHT

    Line3.Line.ForeColorRGBA.Color = srfColorRed
    Line1.Line.ForeColorRGBA.Color = srfColorRed
    Line5.Line.ForeColorRGBA.Color = srfColorRed
    Line6.Line.ForeColorRGBA.Color = srfColorRed
    'Text1.Font.ForeColorRGBA.Color = srfColorDeepNavyBlue

'indicate Type Two Location  ---Right
    Set Line2 = Shapes.AddLine(xBeg:=TTwoX + 0.28, yBeg:=UsMapY - 0.07, _
    	xEnd:=TTwoX + 0.28, yEnd:=UsMapY + 0.45)
    Line2.Line.Width = 0.03                    'short line Right
    Set Line4 = Shapes.AddLine(xBeg:=TTwoX + 0.75, yBeg:=UsMapY - 4.288, _
    	xEnd:=TTwoX + 0.75, yEnd:=UsMapY + 0.15)
    Line4.Line.Width = 0.03                     'long line Right
    Set Text2 = Shapes.AddText(x:= TTwoX + 0.5, y:= UsMapY + 0.52, _
    	Text:="ELJ2")
    Text2.Font.Size = 21
    Text2.Font.Bold = True
    Text2.Font.Italic = True
    Set Line7 = Shapes.AddLine(xBeg:=TTwoX + 0.43, yBeg:=UsMapY + 0.15, _
    	xEnd:=TTwoX + 0.28, yEnd:=UsMapY + 0.15)
    Line7.Line.Width = 0.015
    Line7.EndArrow = srfASSimple              'T2 Arrow LEFT
    Set Line8 = Shapes.AddLine(xBeg:=TTwoX + 0.60, yBeg:=UsMapY + 0.15, _
    	xEnd:=TTwoX + 0.75, yEnd:=UsMapY + 0.15)
    Line8.Line.Width = 0.015
    Line8.EndArrow = srfASSimple               'T2 Arrow RIGHT

    Line2.Line.ForeColorRGBA.Color = srfColorRed
    Line4.Line.ForeColorRGBA.Color = srfColorRed
    Line7.Line.ForeColorRGBA.Color = srfColorRed
    Line8.Line.ForeColorRGBA.Color = srfColorRed

		'Print a meaningful error message for each file that did not grid correctly
'	FileError:
'	Debug.Print  "Error:	" + data_file + "						" + Err.Description
'	Resume Next
End Sub


