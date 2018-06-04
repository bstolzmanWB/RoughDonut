'This macro is used to create rough donut lathe programs
'Created 5/31/18 - Barrett Stolzman

Option Explicit
#Region "Variable Declarations"

	Dim doc As FMDocument
	Dim app As Application
	Dim feature As FMFeature
	Dim stock As FMStock
	Dim partdoc As FMPartDoc
	Dim config As FMConfiguration

	Dim program_number As String
	Dim part_number As String
	Dim revision As String
	Dim operation As String
	Dim operation2 As String
	Dim operation3 As String
	Dim writer As String
	Dim rc As Variant

	'rgh stock dimensions
	Dim od As Double
	Dim id As Double
	Dim oal As Double
	Dim stock_od As Double
	Dim stock_id As Double
	Dim face_1st As Double
	Dim face_2nd As Double
	Dim face_doc As Double
	Dim od_doc As Double
	Dim id_doc As Double

	Dim dlgmaterial As String
	Dim dlgmachine As String
	Dim optfeatures As Double
	Dim chkforging As Double
	Dim material As String

	Dim trepan_od As Double
	Dim trepan_id As Double
	Dim trepan_depth As Double

	Dim fin_allow As Double 'finish allowance on o.d. & i.d.
	Dim turn_clearance As Double 'clearance for o.d. & face
	Dim bore_clearance As Double 'clearance for bore
	Dim bore_allow As Double
	Dim bore_chfr As Double

	'tools
	Dim turnsfm As Double
	Dim turnipr As Double
	Dim boresfm As Double
	Dim boreipr As Double
	Dim groovesfm As Double
	Dim grooveipr As Double
	Dim toolrad1 As Double 'tool radius of turning tool
	Dim turntool As String 'name of facing/turning tool
	Dim toolrad2 As Double 'tool radius of boring bar
	Dim boretool As String 'name of boring bar
	Dim toolrad3 As Double 'tool radius for rough trepan tool
	Dim trepantool As String 'name for rough trepan tool
	Dim trepanwidth As Double 'width of trepan tool

	'flags
	Dim trepanFlag As Boolean

	'constants
	Const pi = Atn(1)*4
#End Region

'*********************************************************************************
'*-----------------------------------Subs----------------------------------------*
'*********************************************************************************

Sub Main
	RoughDonut
End Sub

Private Sub AddIn_OnConnect(ByVal flags As FeatureCAM.tagFMAddInFlags)
	Application.CommandBars.CreateButton("Utilities", "RoughDonut", eMBFID_Hammer1)
End Sub

Private Sub AddIn_OnDisConnect(ByVal flags As FeatureCAM.tagFMAddInFlags)
	Application.CommandBars.DeleteButton("Utilities", "RoughDonut")
End Sub

Sub RoughDonut

	If Application.Version < "24.3.4.016" Then

		If MsgBox("This macro may not be supported in your current version of FeatureCAM. It's written For 2018 release 24.3.4.016 Or later.", vbOkCancel + vbCritical + vbDefaultButton2, "Version Error") = vbOK Then

			trepanFlag = False
			dlgmaterial = "Select Material"
			dlgmachine = "Select Machine"
			face_1st = .1
			face_2nd = .1

			CreateDialogMachine 'call sub for machine selection
			CreateDialogMaterial 'call sub for material selection
			InitializeVars 'call initialize variable routine
			CreateDialogDonut 'call main dimmension input form
			If trepanFlag Then
				CreateDialogTrepan
			End If
			InitializeDoc 'call initialize doc objects routine
			CreateFeature 'call to create features
			End
		End If
	End If
End Sub

Private Sub CreateDialogMachine

	Begin Dialog UserDialog 640,84,"Machine Selection" ' %GRID:10,7,1,1
		CancelButton 440,28,10,7
		PushButton 30,21,130,42,"5222",.PushButton1
		PushButton 180,21,130,42,"5219",.PushButton2
		PushButton 330,21,130,42,"5230",.PushButton3
		PushButton 480,21,130,42,"5506",.PushButton4

	End Dialog
	Dim dlg As UserDialog

		Dim rc As Integer

            rc = Dialog( dlg)

            If (rc = 0) Then
				End

            ElseIf rc = 1 Then
            dlgmachine = "5222"
            fin_allow = .05

            ElseIf rc = 2 Then
            dlgmachine = "5219"
            fin_allow = .05

            ElseIf rc = 3 Then
            dlgmachine = "5230"
            fin_allow = .1

            ElseIf rc = 4 Then
            dlgmachine = "5506"
            fin_allow = .1

			End If

End Sub

Private Sub CreateDialogMaterial

	Begin Dialog UserDialog 620,84,"Material Selection" ' %GRID:10,7,1,1
		CancelButton 580,28,10,7
		PushButton 20,21,130,42,"1018 / Bronze",.PushButton1
		PushButton 170,21,130,42,"4140",.PushButton2
		PushButton 320,21,130,42,"HT Steel",.PushButton3
		PushButton 470,21,130,42,"Copper",.PushButton4

	End Dialog
	Dim dlg As UserDialog

		Dim rc As Integer

            rc = Dialog(dlg)

            If (rc = 0) Then
				End

            ElseIf rc = 1 Then
            dlgmaterial = "1018 / Bronze"

			ElseIf rc = 2 Then
            dlgmaterial = "4140"

            ElseIf rc = 3 Then
            dlgmaterial = "HT Steel"

            ElseIf rc = 4 Then
            dlgmaterial = "Copper"

			End If


End Sub

Private Sub CreateDialogDonut

	Begin Dialog UserDialog 870,224,"Stock Dimensions & Feature Selection" ' %GRID:10,7,1,1
		TextBox 140,14,70,21,.program_number
		Text 10,21,120,14,"Program Number",.text1,1
		TextBox 140,42,120,21,.part_number
		Text 40,49,90,14,"Part Number",.text2,1
		TextBox 140,70,70,21,.Revision
		Text 20,77,110,14,"Revision Level",.text3,1
		TextBox 140,98,35,21,.operation
		TextBox 185,98,35,21,.operation2
		TextBox 230,98,35,21,.operation3
		Text 10,105,120,14,"Op. Number",.text4,1
		TextBox 140,126,70,21,.writer
		Text 40,133,90,14,"Entered By",.text5,1

		TextBox 360,14,70,21,.od
		Text 290,21,60,14,"O.D.",.text6,1
		TextBox 360,42,70,21,.ID
		Text 290,49,60,14,"I.D.",.text7,1
		TextBox 360,70,70,21,.oal
		Text 270,77,80,14,"Length",.text8,1
		TextBox 360,98,70,21,.stock_od
		Text 270,105,80,14,"Stock O.D.",.text9,1
		TextBox 360,126,70,21,.stock_id
		Text 270,133,80,14,"Stock I.D.",.text10,1

		TextBox 590,14,70,21,.face_1st
		Text 450,21,130,14,"Face Stock 1st Side",.Text11,1
		TextBox 590,42,70,21,.face_2nd
		Text 440,49,140,14,"Face Stock 2nd Side",.Text12,1
		TextBox 590,70,70,21,.face_doc
		Text 470,77,110,14,"Face DOC",.text13,1
		TextBox 590,98,70,21,.od_doc
		Text 460,105,120,14,"O.D. DOC",.text14,1
		TextBox 590,126,70,21,.id_doc
		Text 470,133,110,14,"Bore DOC",.text15,1

		Text 670,133,140,14,"Stock Is Forged Ring",.text25,1

		GroupBox 690,7,160,63,"Bore Features",.GroupBox1
		GroupBox 690,84,160,35,"Other Features",.GroupBox2
		Text 700,21,110,14,"Standard Bore",.text22,1
		Text 700,35,110,14,"Tight Tol. Bore",.text23,1
		Text 700,49,110,14,"Eccentric Bore",.text24,1
		Text 700,98,110,14,"Trepan",.text26,1
		OptionGroup .optfeatures
			OptionButton 820,21,20,14,"",.optstandardbore
			OptionButton 820,35,20,14,"",.opttightbore
			OptionButton 820,49,20,14,"",.opteccbore
			OptionButton 820,98,20,14,"",.opttrepan

		CheckBox 820,133,14,14,"chkforging",.chkforging
		PushButton 690,161,160,21,dlgmaterial,.dlgmaterial
		PushButton 690,189,160,21,dlgmachine,.dlgmachine

		OKButton 290,182,120,28
		CancelButton 450,182,120,28



	End Dialog


	ShowDialogP:


	'Dialog dlg

	Dim dlg As UserDialog
	dlg.program_number = CStr(program_number)
	dlg.part_number = CStr(part_number)
	dlg.Revision = CStr(revision)
	dlg.operation = CStr(operation)
	dlg.operation2 = CStr(operation2)
	dlg.operation3 = CStr(operation3)
	dlg.writer = CStr(writer)
	dlg.od = CStr(od)
	dlg.ID = CStr(id)
	dlg.oal = CStr(oal)
	dlg.stock_od = CStr(stock_od)
	dlg.stock_id = CStr(stock_id)
	dlg.face_1st = CStr(face_1st)
	dlg.face_2nd = CStr(face_2nd)
	dlg.face_doc = CStr(face_doc)
	dlg.od_doc = CStr(od_doc)
	dlg.id_doc = CStr(id_doc)
	dlg.optfeatures = CStr(optfeatures)
	dlg.chkforging = CStr(chkforging)

	rc = Dialog( dlg)

	If (rc = 0) Then
		End
	ElseIf rc = 1 Then
	program_number = CStr(dlg.program_number)
	part_number = CStr(dlg.part_number)
	revision = CStr(dlg.Revision)
	operation = CStr(dlg.operation)
	operation2 = CStr(dlg.operation2)
	operation3 = CStr(dlg.operation3)
	writer = CStr(dlg.writer)
	od = CDbl(dlg.od)
	id = CDbl(dlg.ID)
	oal = CDbl(dlg.oal)
	stock_od = CDbl(dlg.stock_od)
	stock_id = CDbl(dlg.stock_id)
	face_1st = CDbl(dlg.face_1st)
	face_2nd = CDbl(dlg.face_2nd)
	face_doc = CDbl(dlg.face_doc)
	od_doc = CDbl(dlg.od_doc)
	id_doc = CDbl(dlg.id_doc)
	optfeatures = CDbl(dlg.optfeatures)
	chkforging = CDbl(dlg.chkforging)

    	CreateDialogMaterial
    	End

    ElseIf rc = 2 Then
    program_number = CStr(dlg.program_number)
	part_number = CStr(dlg.part_number)
	revision = CStr(dlg.Revision)
	operation = CStr(dlg.operation)
	operation2 = CStr(dlg.operation2)
	operation3 = CStr(dlg.operation3)
	writer = CStr(dlg.writer)
	od = CDbl(dlg.od)
	id = CDbl(dlg.ID)
	oal = CDbl(dlg.oal)
	stock_od = CDbl(dlg.stock_od)
	stock_id = CDbl(dlg.stock_id)
	face_1st = CDbl(dlg.face_1st)
	face_2nd = CDbl(dlg.face_2nd)
	face_doc = CDbl(dlg.face_doc)
	od_doc = CDbl(dlg.od_doc)
	id_doc = CDbl(dlg.id_doc)
	optfeatures = CDbl(dlg.optfeatures)
	chkforging = CDbl(dlg.chkforging)

    	CreateDialogMachine
    	End

	ElseIf( rc = -1) Then

	program_number = CStr(dlg.program_number)
	part_number = CStr(dlg.part_number)
	revision = CStr(dlg.Revision)
	operation = CStr(dlg.operation)
	operation2 = CStr(dlg.operation2)
	operation3 = CStr(dlg.operation3)
	writer = CStr(dlg.writer)
	od = CDbl(dlg.od)
	id = CDbl(dlg.ID)
	oal = CDbl(dlg.oal)
	stock_od = CDbl(dlg.stock_od)
	stock_id = CDbl(dlg.stock_id)
	face_1st = CDbl(dlg.face_1st)
	face_2nd = CDbl(dlg.face_2nd)
	face_doc = CDbl(dlg.face_doc)
	od_doc = CDbl(dlg.od_doc)
	id_doc = CDbl(dlg.id_doc)
	optfeatures = CDbl(dlg.optfeatures)
	chkforging = CDbl(dlg.chkforging)


	If ValidInputDonut = False Then
	GoTo ShowDialogP
	End If

	If optfeatures = 3 Then
		trepanFlag = True
	End If

	End If
End Sub

Private Sub CreateDialogTrepan

	'trepan_od = od-.25
	'trepan_id = id+.25


	Begin Dialog UserDialog 580,119,"Trepan Dimensions" ' %GRID:10,7,1,1
		TextBox 110,14,70,21,.trepan_od
		TextBox 290,14,70,21,.trepan_id
		TextBox 480,14,70,21,.trepan_depth
		Text 10,21,90,14,"Trepan O.D.",.text16,1
		Text 200,21,80,14,"Trepan I.D.",.text17,1
		Text 380,21,90,14,"Trepan Depth",.text18,1


		OKButton 150,70,120,28
		CancelButton 310,70,120,28

	End Dialog


	ShowDialogTrepan:


	'Dialog dlg

	Dim dlg As UserDialog
	dlg.trepan_od = CStr(trepan_od)
	dlg.trepan_id = CStr(trepan_id)
	dlg.trepan_depth = CStr(trepan_depth)


	rc = Dialog( dlg)

	If (rc = 0) Then
		End

	ElseIf( rc = -1) Then

	trepan_od = CDbl(dlg.trepan_od)
	trepan_id = CDbl(dlg.trepan_id)
	trepan_depth = CDbl(dlg.trepan_depth)


	If ValidInputTrepan = False Then
	GoTo ShowDialogTrepan
	End If
	End If



End Sub

Private Sub CreateTrepan

'geometry for trepan
Dim ODLine, backLine, IDLine, UhhLine As FMLine

Set ODLine  = doc.Geometry.AddLine2Points(trepan_od/2,0,0,trepan_od/2,0,-trepan_depth)
Set backLine = doc.Geometry.AddLine2Points(trepan_od/2,0,-trepan_depth,trepan_id/2,0,-trepan_depth)
Set IDLine = doc.Geometry.AddLine2Points(trepan_id/2,0,-trepan_depth,trepan_id/2,0,0)

'trepan features
If trepan_od > 4.999 And trepan_od/2-trepan_id/2-(trepan_depth/Tan(28*pi/180))-.1-Sin(45*pi/180)*(.1+toolrad2-Sin(45*pi/180)*(toolrad2))+toolrad2>.8   Then

Set UhhLine = doc.Geometry.AddLine2Points(trepan_od/2,0,0,trepan_od/2-trepan_depth/(Tan(28*pi/180)),0,-trepan_depth)

toolrad3 = .046
If dlgmachine = "5230" Then
    trepantool = "DDJNR (trepan)"
Else
    trepantool = "DDJNL (trepan)"
End If

Dim trepandoc

If dlgmaterial = "Copper" Then
    trepandoc = .05
Else
    trepandoc = .12
End If

'rough trepan
Dim rghCurveL As FMCurvePtList
Dim rghCurve As FMCurve

Set rghCurveL = doc.CreateCurvePtList
rghCurveL.AddPt(trepan_od/2,0,0)
rghCurveL.AddPt(trepan_od/2-trepan_depth/(Tan(28*pi/180)),0,-trepan_depth)
rghCurveL.AddPt(trepan_id/2,0,-trepan_depth)
rghCurveL.AddPt(trepan_id/2,0,0)
Set rghCurve = rghCurveL.AddCurveFromPtList(False)

Dim rghTurn As FMTurn

Set rghTurn = doc.Features.AddTurn(rghCurve,,,)
rghTurn.SetAttribute(eAID_DoFinish,,False,False,)
rghTurn.SetAttribute(eAID_TurnCycleType,,1,False,)
rghTurn.Operations(1).SetAttribute(eAID_TurnRoughDOC,,trepandoc,False,)
rghTurn.Operations(1).SetAttribute(eAID_MaxZBound,,-oal,False,)
rghTurn.Operations(1).SetAttribute(eAID_TurnFinishXAllow,,.03,False,)
rghTurn.Operations(1).SetAttribute(eAID_TurnFinishZAllow,,.03,False,)
rghTurn.Operations(1).SetAttribute(eAID_SkipWallPass,,True,False,)
rghTurn.Operations(1).SetAttribute(eAID_TurnRoughEngageWindrawAtRapid,,False,False,)
'rghTurn.Operations(1).SetAttribute(eAID_TurnRapidFeedRate,,.009,False,)
rghTurn.Operations(1).OverrideSpeed(turnsfm,)
rghTurn.Operations(1).OverrideFeed(.012,)
rghTurn.Operations(1).OverrideTool(trepantool,)
rghTurn.Name = "rough_trepan"

'finish trepan bottom
Dim bottomStartpt, bottomEndpt As String

bottomStartpt = "pt(" & trepan_od/2-Sin(28*pi/180)*(toolrad3)+Cos(28*pi/180)*(.1+toolrad3) & ",0," & Cos(28*pi/180)*(toolrad3)+Sin(28*pi/180)*(.1+toolrad3)+(.1+toolrad3) &")"
bottomEndpt = "pt(" & trepan_id/2+.03+toolrad3+Sin(45*pi/180)*(.1+toolrad3) & ",0," & Cos(45*pi/180)*(.1+toolrad3*2) &")"

Dim bottomCurveL As FMCurvePtList
Dim bottomCurve As FMCurve

Set bottomCurveL = doc.CreateCurvePtList
bottomCurveL.AddPt(trepan_od/2,0,0)
bottomCurveL.AddPt(trepan_od/2-trepan_depth/(Tan(28*pi/180)),0,-trepan_depth)
bottomCurveL.AddPt(trepan_id/2+.03,0,-trepan_depth)
bottomCurveL.AddPt(trepan_id/2+.03,0,-trepan_depth+.05)
Set bottomCurve = bottomCurveL.AddCurveFromPtList(False)

Dim bottomTurn As FMTurn

Set bottomTurn = doc.Features.AddTurn(bottomCurve,,,)
bottomTurn.SetAttribute(eAID_DoRough,,False,False,)
bottomTurn.SetAttribute(eAID_TurnCycleType,,1,False,)
bottomTurn.Operations(1).OverrideSpeed(turnsfm,)
bottomTurn.Operations(1).OverrideFeed(.012,)
bottomTurn.Operations(1).SetAttribute(eAID_TurnStartPt,,bottomStartpt,False,)
bottomTurn.Operations(1).SetAttribute(eAID_TurnEndPt,,bottomEndpt,False,)
bottomTurn.Operations(1).OverrideTool(trepantool,)
bottomTurn.Name = "fin_trepan_bottom"

'finish hub dia

Dim hubStartpt, hubEndpt As String

hubStartpt = "pt(" & trepan_id/2+toolrad3-.02-Sin(45*pi/180)*(.1)-toolrad3 & ",0," & Cos(45*pi/180)*(.1+toolrad3*2) &")"
hubEndpt = "pt(" & trepan_id/2+.09+toolrad3+Sin(45*pi/180)*(.1+toolrad3) & ",0," & Cos(45*pi/180)*(.1+toolrad3*2) &")"

Dim hubCurveL As FMCurvePtList
Dim hubCurve As FMCurve

Set hubCurveL = doc.CreateCurvePtList
hubCurveL.AddPt(trepan_id/2-.02,0,0)
hubCurveL.AddPt(trepan_id/2,0,-.02)
hubCurveL.AddPt(trepan_id/2,0,-trepan_depth)
hubCurveL.AddPt(trepan_id/2+.09,0,-trepan_depth)
Set hubCurve = hubCurveL.AddCurveFromPtList(False)

Dim hubTurn As FMTurn

Set hubTurn = doc.Features.AddTurn(hubCurve,,,)
hubTurn.SetAttribute(eAID_DoRough,,False,False,)
hubTurn.Operations(1).OverrideSpeed(turnsfm,)
hubTurn.Operations(1).OverrideFeed(.012,)
hubTurn.Operations(1).SetAttribute(eAID_TurnStartPt,,hubStartpt,False,)
hubTurn.Operations(1).SetAttribute(eAID_TurnEndPt,,hubEndpt,False,)
hubTurn.Operations(1).OverrideTool(trepantool,)
hubTurn.Name = "fin_hub_dia"

'rough match trepan
Dim rghMatchEndpt

rghMatchEndpt = "pt(" & trepan_od/2-.03-toolrad2-Sin(45*pi/180)*.025 & ",0," & -trepan_depth+.03+toolrad2+Cos(45*pi/180)*.025 &")"

Dim rghMatchCurveL As FMCurvePtList
Dim rghMatchCurve As FMCurve

Set rghMatchCurveL = doc.CreateCurvePtList
    rghMatchCurveL.AddPt(trepan_od/2,0,0)
    rghMatchCurveL.AddPt(trepan_od/2,0,-trepan_depth)
    rghMatchCurveL.AddPt(trepan_od/2-trepan_depth/(Tan(28*pi/180)),0,-trepan_depth)
Set rghMatchCurve = rghMatchCurveL.AddCurveFromPtList(False)

Dim rghMatchBore As FMBore

Set rghMatchBore = doc.Features.AddBore(rghMatchCurve,UhhLine,,)
    rghMatchBore.SetAttribute(eAID_DoFinish,,False,False,)
    rghMatchBore.Operations(1).SetAttribute(eAID_TurnRoughDOC,,trepandoc,False,)
    rghMatchBore.Operations(1).SetAttribute(eAID_MinXBound,,trepan_od/2-trepan_depth/(Tan(28*pi/180)),False,)
    rghMatchBore.Operations(1).SetAttribute(eAID_TurnFinishXAllow,,.03,False,)
    rghMatchBore.Operations(1).SetAttribute(eAID_TurnFinishZAllow,,.03,False,)
    rghMatchBore.Operations(1).SetAttribute(eAID_SkipWallPass,,True,False,)
    rghMatchBore.Operations(1).SetAttribute(eAID_TurnRoughEngageWindrawAtRapid,,True,False,)
    rghMatchBore.Operations(1).SetAttribute(eAID_TurnEndPt,,rghMatchEndpt,False,)
    rghMatchBore.Operations(1).OverrideSpeed(boresfm,)
    rghMatchBore.Operations(1).OverrideFeed(.012,)
    rghMatchBore.Operations(1).OverrideTool(boretool,)
    rghMatchBore.Name = "rough_rghMatch_trepan"

'finish match trepan
Dim finMatchStartpt, finMatchEndpt As String

finMatchStartpt = "pt(" & trepan_od/2-trepan_depth/(Tan(28*pi/180))-.1-Sin(45*pi/180)*(.1+toolrad2-Sin(45*pi/180)*(toolrad2)) & ",0," & toolrad2-.05+Cos(45*pi/180)*(.1+toolrad2-Cos(45*pi/180)*(toolrad2)) &")"
finMatchEndpt = "pt(" & trepan_od/2-.03-toolrad2-Sin(45*pi/180)*(toolrad2+.1) & ",0," & trepan_depth-.05+Cos(45*pi/180)*(.1+toolrad3*2) &")"

Dim finMatchCurveL As FMCurvePtList
Dim finMatchCurve As FMCurve

Set finMatchCurveL = doc.CreateCurvePtList
    finMatchCurveL.AddPt(trepan_od/2-.03,0,-trepan_depth + .05)
    finMatchCurveL.AddPt(trepan_od/2-.03,0,-trepan_depth)
    finMatchCurveL.AddPt(trepan_od/2-trepan_depth/(Tan(28*pi/180))-.1,0,-trepan_depth)
Set finMatchCurve = finMatchCurveL.AddCurveFromPtList(False)

Dim finMatchBore As FMBore

Set finMatchBore = doc.Features.AddBore(finMatchCurve,,,)
    finMatchBore.SetAttribute(eAID_DoRough,,False,False,)
    finMatchBore.SetAttribute(eAID_TurnCycleType,,1,False,) 'change cycle from bore to face
    finMatchBore.SetAttribute(eAID_TurnReverseFinish,,True,,) 'change feed direction from negative to positive
    finMatchBore.Operations(1).SetAttribute(eAID_TurnEngageAng,,45,False,)
    finMatchBore.Operations(1).SetAttribute(eAID_TurnStartPt,,finMatchStartpt,False,)
    finMatchBore.Operations(1).SetAttribute(eAID_TurnEndPt,,finMatchEndpt,False,)
    finMatchBore.Operations(1).OverrideSpeed(boresfm,)
    finMatchBore.Operations(1).OverrideFeed(.012,)
    finMatchBore.Operations(1).OverrideTool(boretool,)
    finMatchBore.Name = "fin_match_trepan"

'finish outer trepan dia
Dim finOuterStartpt, finOuterEndpt

finOuterStartpt = "pt(" & trepan_od/2+.02+Sin(45*pi/180)*(.1) & ",0," & Cos(45*pi/180)*(.1+toolrad2*2) &")"
finOuterEndpt = "pt(" & trepan_od/2-.09-Sin(45*pi/180)*(toolrad2+.1) & ",0," & Cos(45*pi/180)*(.1+toolrad2*2) &")"

Dim finOuterCurveL As FMCurvePtList
Dim finOuterCurve As FMCurve

Set finOuterCurveL = doc.CreateCurvePtList
    finOuterCurveL.AddPt(trepan_od/2+.02,0,0)
    finOuterCurveL.AddPt(trepan_od/2,0,-.02)
    finOuterCurveL.AddPt(trepan_od/2,0,-trepan_depth)
    finOuterCurveL.AddPt(trepan_od/2-.09,0,-trepan_depth)
Set finOuterCurve = finOuterCurveL.AddCurveFromPtList(False)

Dim finBore As FMBore

Set finBore = doc.Features.AddBore(finOuterCurve,,,)
    finBore.SetAttribute(eAID_DoRough,,False,False,)
    finBore.Operations(1).SetAttribute(eAID_TurnStartPt,,finOuterStartpt,False,)
    finBore.Operations(1).SetAttribute(eAID_TurnEndPt,,finOuterEndpt,False,)
    finBore.Operations(1).OverrideSpeed(boresfm,)
    finBore.Operations(1).OverrideFeed(.012,)
    finBore.Operations(1).OverrideTool(boretool,)
    finBore.Name = "fin_outer_trepan_dia"

Else

toolrad3 = .03
trepantool = ".250 w (.03 tnr) M/Cr Face"
trepanwidth = .25

'rough trepan
Dim rghEndpt As String

rghEndpt = "pt(" & trepan_id/2+.05+toolrad3 & ",0," & .1+toolrad3 &")"

Dim curveL As FMCurvePtList
Dim curve As FMCurve

Set curveL = doc.CreateCurvePtList
    curveL.AddPt(trepan_od/2,0,0)
    curveL.AddPt(trepan_od/2,0,-trepan_depth)
    curveL.AddPt(trepan_id/2,0,-trepan_depth)
    curveL.AddPt(trepan_id/2,0,0)
Set curve = curveL.AddCurveFromPtList(False)

Dim rghGroove As FMTurnGroove

Set rghGroove = doc.Features.AddProfileTurnGroove(rghCurve,,eGO_Face,,)
    rghGroove.SetAttribute(eAID_DoFinish,,False,False,)
    rghGroove.Operations(1).SetAttribute(eAID_TurnFinishXAllow,,.05,False,)
    rghGroove.Operations(1).SetAttribute(eAID_TurnFinishZAllow,,.01,False,)
    rghGroove.Operations(1).SetAttribute(eAID_TurnGrvRoughDOC,,5,,)
    rghGroove.Operations(1).OverrideTool(trepantool,)
    rghGroove.Operations(1).SetAttribute(eAID_TurnEndPt,,rghEndpt,False,)
    rghGroove.Name = "rough_trepan"

'finish trepan
Dim finStartpt, finEndpt As String

finStartpt = "pt(" & trepan_id/2-.02+Sin(45*pi/180)*toolrad3 & ",0," & .1+toolrad3 &")"
finEndpt = "pt(" & trepan_id/2+.02+toolrad3 & ",0," & .1+toolrad3 &")"

Dim finCurveL As FMCurvePtList
Dim finCurve As FMCurve

Set finCurveL = doc.CreateCurvePtList
    finCurveL.AddPt(trepan_od/2+.02,0,0)
    finCurveL.AddPt(trepan_od/2,0,-.02)
    finCurveL.AddPt(trepan_od/2,0,-trepan_depth)
    finCurveL.AddPt(trepan_id/2,0,-trepan_depth)
    finCurveL.AddPt(trepan_id/2,0,-.02)
    finCurveL.AddPt(trepan_id/2-.02,0,0)
Set finCurve = finCurveL.AddCurveFromPtList(False)

Dim finGroove As FMTurnGroove

Set finGroove = doc.Features.AddProfileTurnGroove(finCurve,,eGO_Face,,)
    finGroove.SetAttribute(eAID_DoRough,,False,False,)
    finGroove.Operations(1).SetAttribute(eAID_TurnFinishLiftoffDist,,.02,,)
    finGroove.Operations(1).OverrideSpeed(groovesfm,)
    finGroove.Operations(1).OverrideFeed(grooveipr,)
    finGroove.Operations(1).OverrideTool(trepantool,)
    finGroove.Operations(1).SetAttribute(eAID_TurnStartPt,,finStartpt,False,)
    finGroove.Operations(1).SetAttribute(eAID_TurnEndPt,,finEndpt,False,)
    finGroove.Name = "finish_trepan"

End If


End Sub

Private Sub InitializeVars
	If dlgmachine = "5222" Then
		toolrad1 = .046
		turntool = "CNMG 433"
		If dlgmaterial = "Copper" Then
			face_doc =.05
			od_doc = .1
			id_doc = .1
			turnsfm = 1000
            turnipr = .012
            boresfm = 1000
            boreipr = .012
            groovesfm = 400
            grooveipr = .004
		ElseIf dlgmaterial = "HT Steel" Then
			face_doc =.05
			od_doc = .15
			id_doc = .15
			turnsfm = 350
            turnipr = .012
            boresfm = 350
            boreipr = .012
            groovesfm = 250
            grooveipr = .004
        Else
			face_doc =.05
			od_doc = .15
			id_doc = .15
			turnsfm = 600
            turnipr = .012
            boresfm = 600
            boreipr = .012
            groovesfm = 400
            grooveipr = .004
		End If

	ElseIf dlgmachine = "5219" Then
		toolrad1 = .046
		turntool = "CNMG 433"
		If dlgmaterial = "Copper" Then
			face_doc =.05
			od_doc = .1
			id_doc = .1
			turnsfm = 1000
            turnipr = .012
            boresfm = 1000
            boreipr = .012
            groovesfm = 400
            grooveipr = .004
		ElseIf dlgmaterial = "HT Steel" Then
			face_doc =.05
			od_doc = .18
			id_doc = .18
			turnsfm = 350
            turnipr = .012
            boresfm = 350
            boreipr = .012
            groovesfm = 250
            grooveipr = .004
        Else
			face_doc =.05
			od_doc = .18
			id_doc = .18
			turnsfm = 600
            turnipr = .015
            boresfm = 600
            boreipr = .012
            groovesfm = 400
            grooveipr = .004
		End If

	ElseIf dlgmachine = "5230" Then
		toolrad1 = .046
		turntool = "CNMG 433"
		If dlgmaterial = "Copper" Then
			face_doc =.08
			od_doc = .09
			id_doc = .08
			turnsfm = 1200
            turnipr = .012
            boresfm = 1200
            boreipr = .012
            groovesfm = 600
            grooveipr = .004
		ElseIf dlgmaterial = "HT Steel" Then
			face_doc =.1
			od_doc = .2
			id_doc = .18
			turnsfm = 350
            turnipr = .015
            boresfm = 350
            boreipr = .015
            groovesfm = 300
            grooveipr = .003
		ElseIf dlgmaterial = "4140" Then
			face_doc =.1
			od_doc = .2
			id_doc = .18
			turnsfm = 600
            turnipr = .015
            boresfm = 600
            boreipr = .015
            groovesfm = 400
            grooveipr = .004
        Else
			face_doc =.1
			od_doc = .2
			id_doc = .18
			turnsfm = 600
            turnipr = .015
            boresfm = 500
            boreipr = .015
            groovesfm = 500
            grooveipr = .004

		End If

	ElseIf dlgmachine = "5506" Then
		toolrad1 = .046
		turntool = "CNMG 543"
		toolrad2 = .046
		boretool = "CNMG 543 Bore"
		If dlgmaterial = "Copper" Then
			face_doc =.08
			od_doc = .09
			id_doc = .08
			turnsfm = 1200
            turnipr = .012
            boresfm = 1200
            boreipr = .012
            groovesfm = 600
            grooveipr = .004
		ElseIf dlgmaterial = "HT Steel" Then
			face_doc =.15
			od_doc = .15
			id_doc = .15
			turnsfm = 350
            turnipr = .015
            boresfm = 350
            boreipr = .015
            groovesfm = 300
            grooveipr = .003
		ElseIf dlgmaterial = "4140" Then
			face_doc =.15
			od_doc = .15
			id_doc = .15
			turnsfm = 500
            turnipr = .015
            boresfm = 500
            boreipr = .015
            groovesfm = 400
            grooveipr = .004
        Else
			face_doc =.15
			od_doc = .15
			id_doc = .15
			turnsfm = 600
            turnipr = .018
            boresfm = 600
            boreipr = .018
            groovesfm = 400
            grooveipr = .004

		End If
	End If

	If stock_id<3 Then
    toolrad2 = .031
    boretool = "BB-CNMG 432"
	ElseIf optfeatures = 3 Then
	    toolrad2 = .031
	    boretool = "BB-CNMG 432"
	Else
	    toolrad2 = .046
	    boretool = "BB-CNMG 433"
	End If

	If optfeatures = 1 Then
	    bore_allow = .015
	Else
	    bore_allow = 0
	End If

	If optfeatures = 2 Then
	    bore_chfr = .08
	Else
	    bore_chfr = .02
	End If

End Sub

Private Sub InitializeDoc

	Set app = ActiveDocument.Application
	Set doc = Application.Documents.AddFM(eST_Turning,True)
	Set doc = Application.ActiveDocument
	Set stock = doc.Stock

	doc.SetAttribute(eAID_TurnMinToolChange,,0,,)

	'set config files, tool crib, and post options for each material and machine
	If dlgmachine = "5219" And dlgmaterial = "1018 / Bronze" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5219\5219.cnc" )
	doc.ActiveToolCrib = "5214"
	app.SetTurnPostOptions(,10,5,30,10,2000)
	Set config = Application.Configurations.Item("5214-1018")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5214\5214.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5214 has been copied to local directory"
		End If
		stock.Material = "5214-1018"

	ElseIf dlgmachine = "5219" And dlgmaterial = "4140" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5219\5219.cnc" )
	doc.ActiveToolCrib = "5214"
	app.SetTurnPostOptions(,10,5,30,10,2000)
	Set config = Application.Configurations.Item("5214-4140")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5214\5214.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5214 has been copied to local directory"
		End If
		stock.Material = "5214-4140"

	ElseIf dlgmachine = "5219" And dlgmaterial = "HT Steel" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5219\5219.cnc" )
	doc.ActiveToolCrib = "5214"
	app.SetTurnPostOptions(,10,5,30,10,2000)
	Set config = Application.Configurations.Item("5214-32-38RC")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5214\5214.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5214 has been copied to local directory"
		End If
		stock.Material = "5214-32-38RC"

	ElseIf dlgmachine = "5219" And dlgmaterial = "Copper" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5219\5219.cnc" )
	doc.ActiveToolCrib = "5214"
	app.SetTurnPostOptions(,10,5,30,10,2000)
	Set config = Application.Configurations.Item("5214-COP")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5214\5214.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5214 has been copied to local directory"
		End If
		stock.Material = "5214-COP"

	ElseIf dlgmachine = "5222" And dlgmaterial = "1018 / Bronze" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5222\5222.cnc" )
	doc.ActiveToolCrib = "5222"
	app.SetTurnPostOptions(,10,5,30,10,2500)
	Set config = Application.Configurations.Item("5222-1018")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5222\5222.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5222 has been copied to local directory"
		End If
		stock.Material = "5222-1018"

	ElseIf dlgmachine = "5222" And dlgmaterial = "4140" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5222\5222.cnc" )
	doc.ActiveToolCrib = "5222"
	app.SetTurnPostOptions(,10,5,30,10,2500)
	Set config = Application.Configurations.Item("5222-4140")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5222\5222.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5222 has been copied to local directory"
		End If
		stock.Material = "5222-4140"

	ElseIf dlgmachine = "5222" And dlgmaterial = "HT Steel" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5222\5222.cnc" )
	doc.ActiveToolCrib = "5222"
	app.SetTurnPostOptions(,10,5,30,10,2500)
	Set config = Application.Configurations.Item("5222-32-38RC")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5222\5222.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5222 has been copied to local directory"
		End If
		stock.Material = "5222-32-38RC"

	ElseIf dlgmachine = "5222" And dlgmaterial = "Copper" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5222\5222.cnc" )
	doc.ActiveToolCrib = "5222"
	app.SetTurnPostOptions(,10,5,30,10,2500)
	Set config = Application.Configurations.Item("5222-COP")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5222\5222.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5222 has been copied to local directory"
		End If
		stock.Material = "5222-COP"



	ElseIf dlgmachine = "5230" And dlgmaterial = "1018 / Bronze" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5230\5230.cnc" )
	doc.ActiveToolCrib = "5230"
	app.SetTurnPostOptions(,10,5,32,10,350)
	Set config = Application.Configurations.Item("5230-1018")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5230\5230.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5230 has been copied to local directory"
		End If
		stock.Material = "5230-1018"

	ElseIf dlgmachine = "5230" And dlgmaterial = "4140" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5230\5230.cnc" )
	doc.ActiveToolCrib = "5230"
	app.SetTurnPostOptions(,10,5,32,10,350)
	Set config = Application.Configurations.Item("5230-4140")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5230\5230.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5230 has been copied to local directory"
		End If
		stock.Material = "5230-4140"

	ElseIf dlgmachine = "5230" And dlgmaterial = "HT Steel" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5230\5230.cnc" )
	doc.ActiveToolCrib = "5230"
	app.SetTurnPostOptions(,10,5,32,10,350)
	Set config = Application.Configurations.Item("5230-4140")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5230\5230.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5230 has been copied to local directory"
		End If
		stock.Material = "5230-4140"

	ElseIf dlgmachine = "5230" And dlgmaterial = "Copper" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5230\5230.cnc" )
	doc.ActiveToolCrib = "5230"
	app.SetTurnPostOptions(,10,5,32,10,350)
	Set config = Application.Configurations.Item("5230-COP")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5230\5230.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5230 has been copied to local directory"
		End If
		stock.Material = "5230-COP"



	ElseIf dlgmachine = "5506" And dlgmaterial = "1018 / Bronze" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5506\5506.cnc" )
	doc.ActiveToolCrib = "5506"
	app.SetTurnPostOptions(,10,5,32,10,350)
	Set config = Application.Configurations.Item("5506-STEEL")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5506\5506.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5506 has been copied to local directory"
		End If
		stock.Material = "5506-1018"

	ElseIf dlgmachine = "5506" And dlgmaterial = "4140" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5506\5506.cnc" )
	doc.ActiveToolCrib = "5506"
	app.SetTurnPostOptions(,10,5,32,10,350)
	Set config = Application.Configurations.Item("5506-STEEL")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5506\5506.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5506 has been copied to local directory"
		End If
		stock.Material = "5506-4140"

	ElseIf dlgmachine = "5506" And dlgmaterial = "HT Steel" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5506\5506.cnc" )
	doc.ActiveToolCrib = "5506"
	app.SetTurnPostOptions(,10,5,32,10,350)
	Set config = Application.Configurations.Item("5506-STEEL")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5506\5506.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5506 has been copied to local directory"
		End If
		stock.Material = "5506-4140"

	ElseIf dlgmachine = "5506" And dlgmaterial = "Copper" Then
	app.SetTurnPostOptions( "\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5506\5506.cnc" )
	doc.ActiveToolCrib = "5506"
	app.SetTurnPostOptions(,10,5,32,10,350)
	Set config = Application.Configurations.Item("5506-COP")
	If Not config Is Nothing Then
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		Else
		Application.Configurations.Import("\\ant-fs-01\departments\mfg-eng\cnc\FEATURECAM\5506\5506.cdb")
		Application.Configurations.Item( doc.Name ).CopyConfiguration( config.Name )
		MsgBox "Configuration file for 5506 has been copied to local directory"
		End If
		stock.Material = "5506-COP"

	End If


End Sub

Private Sub CreateFeature

	Call Setup(1)
'geometry for profile
	CreateRoughGeo

'turn scale off forgings
	If chkforging = 1 Then

		Call CleanOD_Op("front")

	End If

'face 1st side
	Call Face_Op("front")


'rough turn
	Call RoughOD_Op("front")


'fin turn
	Call FinishOD_Op("front")


'rough bore
	RoughBore_Op


'finish bore
	Call FinishBore_Op("front")


'***************************************** 2nd side **********************************************

 'setup 2
 	Call Setup(2)

'turn scale off forgings
	Call CleanOD_Op("back")

'face 2nd side
	Call Face_Op("back")

'************************************************************************************************

'rough turn o.d. 2nd side
	Call RoughOD_Op("back")

'fin turn
	Call FinishOD_Op("back")

'***************************************** trepan **********************************************


	If optfeatures = 3 Then
		CreateTrepan
	End If

'***************************************** fin bore **********************************************

'finish bore or chfr bore
	Call FinishBore_Op("back")

'***************************************** ecc bore **********************************************

	If optfeatures = 2 Then
'setup 3
	Call Setup(3)

'ecc bore
	Eccentric_Op

	End If

'Finish Document
	FinishUp

End Sub

Private Sub FinishUp

	doc.PartDocumentation.Title = part_number
	doc.PartDocumentation.Author = writer
	doc.PartDocumentation.Note1 = revision
	doc.PartDocumentation.Note2 = operation

	Dim Setup As FMSetup

	For Each Setup In doc.Setups
		Setup.PartName = program_number
	Next Setup

	doc.InvalidateAll

	Setup.Item(1).Activate
	'ActiveDocument.SetView(eVT_Top)
	ActiveDocument.SetView(eVT_Isometric)
	'ActiveDocument.Sim3DActiveSetups
	'ActiveDocument.SaveAs( "\\ant-fs-01\departments\mfg-eng\compactII\" & dlgmachine & "\" & program_number & ".FM")


	If optfeatures = 3 And toolrad3 = .046 Then
		If dlgmaterial = "Copper" Then
			MsgBox("For Feature 'rough_trepan' set Engage & Withdraw Feed to .005",,"Set Feed")
		Else
			MsgBox("For Feature 'rough_trepan' set Engage & Withdraw Feed to .009",,"Set Feed")
		End If
	End If

End Sub

'*********************************************************************************
'*--------------------------------Operation Subs---------------------------------*
'*********************************************************************************

Private Sub Setup(opt As Integer)

    If opt = 1 Then

        Dim setup1 As FMSetup

        Set setup1 = doc.ActiveSetup
        setup1.Name = part_number & " Rev." & revision & " Op. " & operation
        setup1.Activate
        setup1.ucs.SetLocation(0,0,0)

        'set stock size
        If chkforging = 1 Then
            stock_od = stock_od-.2
        End If

        Dim stock As FMStock

        Set stock = doc.Stock
        stock.SetDimensions(eST_Round,oal+face_1st+face_2nd,,,stock_od+.2,stock_id,eAT_AxisZ,,)
        stock.SetLocation(0,0,face_1st)
        stock.SingleProgram = False
        ActiveDocument.SetView(eVT_CenterAll)

    ElseIf opt = 2 Then

        Dim setup2 As FMSetup

        Set setup2 = doc.AddSetup("side2",eST_Turning,,,)
        setup2.ucs.Align(eUA_RoundStockBack,,,)
        'setup2.Spindle = eSST_Main
        setup2.Name = part_number & " Rev." & revision & " Op. " & operation2
        setup2.FixtureID = "54"
        setup2.ucs.SetLocation(0,0,-oal,)
        setup2.Activate
        ActiveDocument.SetView(eVT_Top)

    Else

        Dim setup3 As FMSetup

        Set setup3 = doc.AddSetup("side3",eST_Turning,,,)
        setup3.ucs.Align(eUA_RoundStockBack,,,)
        'setup3.Spindle = eSST_Main
        setup3.Name = part_number & " Rev." & revision & " Op. " & operation3
        setup3.FixtureID = "54"
        setup3.ucs.SetLocation(0,0,-oal,)
        setup3.Activate
        ActiveDocument.SetView(eVT_Top)

    End If
End Sub

Private Sub CreateRoughGeo

    Dim FrontLine As FMLine
    Dim IDLine As FMLine
    Dim backLine As FMLine
    Dim ODLine As FMLine

    Set FrontLine = doc.Geometry.AddLine2Points(od/2,0,0,id/2,0,0)
    Set IDLine = doc.Geometry.AddLine2Points(id/2,0,0,id/2,0,-oal)
    Set backLine = doc.Geometry.AddLine2Points(id/2,0,-oal,od/2,0,-oal)
    Set ODLine = doc.Geometry.AddLine2Points(od/2,0,-oal,od/2,0,0)

End Sub

Private Sub Face_Op(opt As String)

If opt = "front" Then

	Dim frontPasses As Integer
	Dim endpt,startpt, semiendpt As String
	Dim frontFace1 As FMTurnFace
	Dim frontFace2 As FMTurnFace

	If (face_1st/face_doc) - Round((face_1st/face_doc),0) > .00000001 Then
	    frontPasses = Round((face_1st/face_doc),0)+1
	Else
	    frontPasses = Round((face_1st/face_doc),0)
	End If

	If frontPasses = 1 Then

	    'endpt = "pt(" & stock_id/2-.1+Sin(45*pi/180)*(toolrad1+.1) & ",0," & toolrad1+Cos(45*pi/180)*(toolrad1+.1) &")"
	    endpt = "pt(" & stock_od/2 + fin_allow + toolrad1 + .1 & ",0,.1)"

	    Set frontFace1 = doc.Features.AddTurnFace(face_1st,True,stock_od+.2,stock_id-.2,,False)
	        frontFace1.SetAttribute(eAID_DoRough,,False)
	        frontFace1.Operations(1).SetAttribute(eAID_TurnEndPt,,endpt,False,)
	        frontFace1.Operations(1).OverrideSpeed(turnsfm,)
	        frontFace1.Operations(1).OverrideFeed(turnipr,)
	        frontFace1.Operations(1).OverrideTool(turntool,)
	        frontFace1.Name = "fin_face"

	ElseIf frontPasses = 2 Then

	    semiendpt = "pt(" & stock_id/2-.1+Sin(45*pi/180)*(toolrad1+.1) & ",0," & toolrad1+Cos(45*pi/180)*(toolrad1+.1) &")"
	    Set frontFace1 = doc.Features.AddTurnFace(face_1st/2,True,stock_od+.2,stock_id-.2,,False)
	        frontFace1.SetFeatureLocation(eFLT_xyz,0,0,face_1st/2,0,0,0,0,,,,,)
	        frontFace1.SetAttribute(eAID_DoRough,,False)
	        'frontFace1.Operations(1).SetAttribute(eAID_TurnEndPt,,semiendpt,False,)
	        frontFace1.Operations(1).OverrideSpeed(turnsfm,)
	        frontFace1.Operations(1).OverrideFeed(turnipr,)
	        frontFace1.Operations(1).OverrideTool(turntool,)
	        frontFace1.Name = "semi_face"

	    endpt = "pt(" & stock_id/2-.1+Sin(45*pi/180)*(toolrad1+.1) & ",0," & toolrad1+Cos(45*pi/180)*(toolrad1+.1) &")"
	    Set frontFace2 = doc.Features.AddTurnFace(face_1st/2,True,stock_od+.2,stock_id-.2,,False)
	        frontFace2.SetAttribute(eAID_TurnDoRough,,False,False,)
	        'frontFace2.Operations(1).SetAttribute(eAID_TurnEndPt,,endpt,False,)
	        frontFace2.Operations(1).OverrideSpeed(turnsfm,)
	        frontFace2.Operations(1).OverrideFeed(turnipr,)
	        frontFace2.Operations(1).OverrideTool(turntool,)
	        frontFace2.Name = "fin_face"

	Else

	    endpt = "pt(" & stock_id/2-.1+Sin(45*pi/180)*(.025) & ",0," & toolrad1+face_1st/frontPasses+Cos(45*pi/180)*(.025) &")"
	    Set  frontFace1 = doc.Features.AddTurnFace(face_1st,True,stock_od+.2,stock_id-.2,,False)
	        frontFace1.SetAttribute(eAID_DoRough,,False)
	        frontFace1.Operations(1).SetAttribute(eAID_TurnFinishZAllow,,0,False,)
	        frontFace1.Operations(1).SetAttribute(eAID_TurnRoughDOC,,face_doc,False,)
	        frontFace1.Operations(1).SetAttribute(eAID_TurnRoughEngageWindrawAtRapid,,True,False,)
	        'frontFace1.Operations(1).SetAttribute(eAID_TurnEndPt,,endpt,False,)
	        frontFace1.Operations(1).OverrideSpeed(turnsfm,)
	        frontFace1.Operations(1).OverrideFeed(turnipr,)
	        frontFace1.Operations(1).OverrideTool(turntool,)
	        frontFace1.Name = "fin_face"

	End If

	If chkforging = 1 Then

	    frontFace1.SetOuterDiameter(stock_od+.2-.25,,False)
	    If frontPasses=2 Then
	        frontFace2.SetOuterDiameter(stock_od+.2-.25,,False)
	    End If
	    If frontPasses <3 Then
	        startpt = "pt(" & stock_od/2+.1-.125+toolrad1+.1 & ",0," & toolrad1 &")"
	    Else
	        startpt = "pt(" & stock_od/2+.1-.125+.1 & ",0," & face_1st-face_1st/frontPasses+toolrad1 &")"

	    End If
	    'frontFace.Operations(1).SetAttribute(eAID_TurnStartPt,,startpt,False,)
	    If frontPasses=2 Then
	        frontFace2.Order = 3
	    End If
	    frontFace1.Order = 2

	End If

ElseIf opt = "back" Then

	Dim backPasses As Integer

	If (face_2nd/face_doc) - Round((face_2nd/face_doc),0) > .00000001 Then
	    backPasses = Round((face_2nd/face_doc),0)+1
	Else
	    backPasses = Round((face_2nd/face_doc),0)
	End If

	Dim backsemiendpt, backendpt As String
	Dim semiFace, finFace As FMTurnFace

	If backPasses = 1 Then

		backendpt = "pt(" & id/2-.05+Sin(45*pi/180)*(toolrad1+.1) & ",0," & toolrad1+Cos(45*pi/180)*(toolrad1+.1) &")"

		Set finFace = doc.Features.AddTurnFace(face_2nd,True,stock_od+.2,id-.1,,False)
		    finFace.SetAttribute(eAID_DoRough,,False)
		    finFace.Operations(1).SetAttribute(eAID_TurnEndPt,,backendpt,False,)
		    finFace.Operations(1).OverrideSpeed(turnsfm,)
		    finFace.Operations(1).OverrideFeed(turnipr,)
		    finFace.Operations(1).OverrideTool(turntool,)
		    finFace.Name = "fin_face_"

	ElseIf backPasses = 2 Then

		backsemiendpt = "pt(" & id/2-.05+Sin(45*pi/180)*(toolrad1+.1) & ",0," & toolrad1+Cos(45*pi/180)*(toolrad1+.1) &")"

		Set semiFace = doc.Features.AddTurnFace(face_2nd/2,True,stock_od+.2,id-.1,,False)
		    semiFace.SetFeatureLocation(eFLT_xyz,0,0,face_2nd/2,0,0,0,0,,,,,)
		    semiFace.SetAttribute(eAID_DoRough,,False,False,)
		    semiFace.Operations(1).SetAttribute(eAID_TurnEndPt,,backsemiendpt,False,)
		    semiFace.Operations(1).OverrideSpeed(turnsfm,)
		    semiFace.Operations(1).OverrideFeed(turnipr,)
		    semiFace.Operations(1).OverrideTool(turntool,)
		    semiFace.Name = "semi_face_"

		backendpt = "pt(" & id/2-.05+Sin(45*pi/180)*(toolrad1+.1) & ",0," & toolrad1+Cos(45*pi/180)*(toolrad1+.1) &")"

		Set finFace = doc.Features.AddTurnFace(face_2nd/2,True,stock_od+.2,id-.1,,False)
		    finFace.SetAttribute(eAID_DoRough,,False,False,)
		    'finFace.Operations(1).SetAttribute(eAID_TurnEndPt,,backendpt,False,)
		    finFace.Operations(1).OverrideSpeed(turnsfm,)
		    finFace.Operations(1).OverrideFeed(turnipr,)
		    finFace.Operations(1).OverrideTool(turntool,)
		    finFace.Name = "fin_face_"

	Else

		backendpt = "pt(" & id/2-.05+Sin(45*pi/180)*(.025) & ",0," & toolrad1+face_2nd/backPasses+Cos(45*pi/180)*(.025) &")"

		Set finFace = doc.Features.AddTurnFace(face_2nd,True,stock_od+.2,id-.1,,False)
		    finFace.SetAttribute(eAID_DoFinish,,False,False,)
		    finFace.Operations(1).SetAttribute(eAID_TurnFinishZAllow,,0,False,)
		    finFace.Operations(1).SetAttribute(eAID_TurnRoughDOC,,face_doc,False,)
		    finFace.Operations(1).SetAttribute(eAID_TurnRoughEngageWindrawAtRapid,,True,False,)
		    finFace.Operations(1).SetAttribute(eAID_TurnEndPt,,backendpt,False,)
		    finFace.Operations(1).OverrideSpeed(turnsfm,)
		    finFace.Operations(1).OverrideFeed(turnipr,)
		    finFace.Operations(1).OverrideTool(turntool,)
		    finFace.Name = "fin_face_"

	End If

	If chkforging = 1 Then
	    finFace.SetOuterDiameter(stock_od+.2-.25,,False)
	    If backPasses=2 Then
	        semiFace.SetOuterDiameter(stock_od+.2-.25,,False)
	    End If
	    Dim backFinStartpt As String
	    If backPasses <3 Then
	        backFinStartpt = "pt(" & stock_od/2+.1-.125+toolrad1+.1 & ",0," & toolrad1 &")"
	    Else
	        backFinStartpt = "pt(" & stock_od/2+.1-.125+.1 & ",0," & face_2nd-face_2nd/backPasses+toolrad1 &")"

	    End If
	    finFace.Operations(1).SetAttribute(eAID_TurnStartPt,,backFinStartpt,False,)
	    If backPasses=2 Then
	        semiFace.Order = 3
	    End If
	    finFace.Order = 2

	End If

End If
End Sub

Private Sub RoughOD_Op(opt As String)

	If opt = "front" Then

	    Dim passes As Integer
	    'figures rough turn depth of cut for setting start point on 1st side                      >.000000001 is used because even # of passes was rounding up
	    If chkforging = 1 Then
	        If stock_od/2+.1-od/2-fin_allow > .125 Then
	            If ((stock_od/2+.1-od/2-.125-fin_allow)/od_doc) - Round(((stock_od/2+.1-od/2-.125-fin_allow)/od_doc),0) > .00000001 Then
	                passes = Round(((stock_od/2+.1-od/2-.125-fin_allow)/od_doc),0)+1
	            Else
	                passes = Round(((stock_od/2+.1-od/2-.125-fin_allow)/od_doc),0)
	            End If
	        End If

	    Else

            If ((stock_od/2+.1-od/2-fin_allow)/od_doc) - Round(((stock_od/2+.1-od/2-fin_allow)/od_doc),0) > .00000001 Then
                passes = Round(((stock_od/2+.1-od/2-fin_allow)/od_doc),0)+1
            Else
                passes = Round(((stock_od/2+.1-od/2-fin_allow)/od_doc),0)
            End If

	    End If

	    Dim turndoc As Double

	    If chkforging = 1 Then
	        If stock_od/2+.1-od/2-fin_allow > .125 Then
	            turndoc = (stock_od/2+.1-od/2-fin_allow-.125) / passes
	        End If
	    Else
	            turndoc = (stock_od/2+.1-od/2-fin_allow) / passes
	    End If

	    Dim startpt, endpt As String

	    startpt = "pt(" & od/2+fin_allow+(passes-1)*turndoc & ",0,.1)"
	    endpt = "pt(" & od/2+fin_allow+toolrad1+Sin(45*pi/180)*.025 & ",0," & Cos(45*pi/180)*(.1+toolrad1*2) &")"

	    'turnendpt = "pt(" & od/2+fin_allow+toolrad1 + .018 +Sin(45*pi/180)*.025 & ",0," & (.1+toolrad1*2) &")"

	    'If chkforging = 1 And stock_od/2+.1-od/2-fin_allow < .125001 Then

	    'Else

	    Dim curveL As FMCurvePtList
	    Dim curve As FMCurve

	    Set curveL = doc.CreateCurvePtList
	        curveL.AddPt(od/2,0,0)
	        curveL.AddPt(od/2,0,-oal/2-.014-.05)
	    Set curve = curveL.AddCurveFromPtList(False)

	    Dim turn As FMTurn

	    Set turn = doc.Features.AddTurn(curve,,,)

	        turn.SetAttribute(eAID_TurnDoFinish,,False)
	        turn.Operations(1).SetAttribute(eAID_TurnFinishXAllow,,fin_allow,False,)
	        turn.Operations(1).SetAttribute(eAID_TurnFinishZAllow,,0,False,)
	        turn.Operations(1).SetAttribute(eAID_TurnRoughDOC,,od_doc,False,)
	        turn.Operations(1).OverrideTool(turntool,)
	        turn.Operations(1).SetAttribute(eAID_SkipWallPass,,True,False,)
	        turn.Operations(1).SetAttribute(eAID_TurnRoughEngageWindrawAtRapid,,True,False,)
	        turn.Operations(1).SetAttribute(eAID_TurnStartPt,,startpt,False,)
	        turn.Operations(1).SetAttribute(eAID_TurnEndPt,,endpt,False,)
	        turn.Operations(1).OverrideSpeed(turnsfm,)
	        turn.Operations(1).OverrideFeed(turnipr,)
	        If chkforging = 1 Then
	        turn.Operations(1).SetAttribute(eAID_MaxXBound,,stock_od/2+.1-.125,False,)
	        End If

	        turn.Name = "rough_od"

	    'End if

    Else

'figures rough turn depth of cut for setting start point on 2nd side                      >.000000001 is used because even # of passes was rounding up
	    Dim backpasses As Integer

	    If chkforging = 1 Then
			If stock_od/2+.1-od/2-fin_allow > .125 Then
				If ((stock_od/2+.1-od/2-.125-fin_allow)/od_doc) - Round(((stock_od/2+.1-od/2-.125-fin_allow)/od_doc),0) > .00000001 Then
					backpasses = Round(((stock_od/2+.1-od/2-.125-fin_allow)/od_doc),0)+1
				Else
					backpasses = Round(((stock_od/2+.1-od/2-.125-fin_allow)/od_doc),0)
				End If
			End If

		Else

			If ((stock_od/2+.1-od/2-fin_allow)/od_doc) - Round(((stock_od/2+.1-od/2-fin_allow)/od_doc),0) > .00000001 Then
				backpasses = Round(((stock_od/2+.1-od/2-fin_allow)/od_doc),0)+1
			Else
				backpasses = Round(((stock_od/2+.1-od/2-fin_allow)/od_doc),0)
			End If

		End If

	    Dim backturndoc As Double

		If chkforging = 1 Then
			If stock_od/2+.1-od/2-fin_allow > .125 Then
				backturndoc = (stock_od/2+.1-od/2-fin_allow-.125) / backpasses
			End If
		Else
			backturndoc = (stock_od/2+.1-od/2-fin_allow) / backpasses
		End If

		startpt = "pt(" & od/2+fin_allow+toolrad1+(passes-1)*backturndoc & ",0,.1)"
		endpt = "pt(" & od/2+fin_allow+toolrad1+Sin(45*pi/180)*.025 & ",0," & Cos(45*pi/180)*(.1+toolrad1*2) &")"

		'If chkforging = 1 And stock_od/2+.1-od/2-fin_allow < .125001 Then

			'Else
	    Dim backcurveL As FMCurvePtList
	    Dim backcurve As FMCurve

		Set backcurveL = doc.CreateCurvePtList
			backcurveL.AddPt(od/2,0,0)
			backcurveL.AddPt(od/2,0,-oal/2-.014+.05)
		Set backcurve = backcurveL.AddCurveFromPtList(False)

	    Dim backTurn As FMTurn

		Set backTurn = doc.Features.AddTurn(backcurve,,,)
			backTurn.SetAttribute(eAID_TurnDoFinish,,False)
			backTurn.Operations(1).SetAttribute(eAID_TurnFinishXAllow,,fin_allow,False,)
			backTurn.Operations(1).SetAttribute(eAID_TurnFinishZAllow,,0,False,)
			backTurn.Operations(1).SetAttribute(eAID_TurnRoughDOC,,od_doc,False,)
			backTurn.Operations(1).OverrideTool(turntool)
			backTurn.Operations(1).SetAttribute(eAID_SkipWallPass,,True,False,)
			backTurn.Operations(1).SetAttribute(eAID_TurnRoughEngageWindrawAtRapid,,True,False,)
			'backTurn.Operations(1).SetAttribute(eAID_TurnStartPt,,startpt,False,)
			'backTurn.Operations(1).SetAttribute(eAID_TurnEndPt,,endpt,False,)
			backTurn.Operations(1).OverrideSpeed(turnsfm,)
			backTurn.Operations(1).OverrideFeed(turnipr,)
			If chkforging = 1 Then
			backTurn.Operations(1).SetAttribute(eAID_MaxXBound,,stock_od/2+.1-.125,False,)
			End If
			backTurn.Name = "rough_od_"
		'End If

    End If

End Sub

Private Sub FinishOD_Op(opt As String)

	Dim startpt, endpt As String
	Dim curveL As FMCurvePtList
	Dim curve As FMCurve
	Dim turn As FMTurn

	If opt = "front" Then



	    startpt = "pt(" & od/2+toolrad1-.02-Sin(45*pi/180)*(.1)-toolrad1 & ",0," & Cos(45*pi/180)*(.1+toolrad1*2) &")"
	    endpt = "pt(" & od/2+toolrad1+Sin(45*pi/180)*(.1+toolrad1) & ",0," & -oal/2-.014-.05+Cos(45*pi/180)*(.1+toolrad1) &")"

	    Set curveL = doc.CreateCurvePtList
	    curveL.AddPt(od/2-.02,0,0)
	    curveL.AddPt(od/2,0,-.02)
	    curveL.AddPt(od/2,0,-oal/2-.014-.05)
	    Set curve = curveL.AddCurveFromPtList(False)

	    Set turn = doc.Features.AddTurn(curve,,,)
	    turn.SetAttribute(eAID_TurnDoRough,,False)
	    turn.Operations(1).OverrideTool(turntool,)
	    turn.Operations(1).OverrideSpeed(turnsfm,)
	    turn.Operations(1).OverrideFeed(turnipr,)
	    turn.Operations(1).SetAttribute(eAID_TurnStartPt,,startpt,False,)
	    'turn.Operations(1).SetAttribute(eAID_TurnEndPt,,endpt,False,)
	    turn.Name = "fin_od"
	Else

		startpt = "pt(" & od/2+toolrad1-.02-Sin(45*pi/180)*(.1)-toolrad1 & ",0," & Cos(45*pi/180)*(.1+toolrad1*2) &")"
		endpt = "pt(" & od/2+toolrad1+Sin(45*pi/180)*(.1+toolrad1) & ",0," & -oal/2-.014+.05+Cos(45*pi/180)*(.1+toolrad1) &")"

		Set curveL = doc.CreateCurvePtList
		curveL.AddPt(od/2-.02,0,0)
		curveL.AddPt(od/2,0,-.02)
		curveL.AddPt(od/2,0,-oal/2-.014+.05)
		Set curve = curveL.AddCurveFromPtList(False)

		Set turn = doc.Features.AddTurn(curve,,,)
		turn.SetAttribute(eAID_TurnDoRough,,False,False,)
		'turn.Operations(1).SetAttribute(eAID_TurnStartPt,,startpt,False,)
		'turn.Operations(1).SetAttribute(eAID_TurnEndPt,,endpt,False,)
		turn.Operations(1).OverrideSpeed(turnsfm,)
		turn.Operations(1).OverrideFeed(turnipr,)
		turn.Operations(1).OverrideTool(turntool)
		turn.Name = "fin_od_"

	End If

End Sub

Private Sub RoughBore_Op

    Dim endpt As String

    endpt = "pt(" & id/2-fin_allow-bore_allow-toolrad2-Sin(45*pi/180)*.025 & ",0," & Cos(45*pi/180)*(.1+toolrad2*2) &")"

    Dim curveL As FMCurvePtList
    Dim curve As FMCurve

    Set curveL = doc.CreateCurvePtList
        curveL.AddPt(id/2-bore_allow,0,0)
        curveL.AddPt(id/2-bore_allow,0,-oal-face_2nd-.05)
    Set curve = curveL.AddCurveFromPtList(False)

    Dim bore As FMBore

    Set bore = doc.Features.AddBore(curve,,,)
        bore.SetAttribute(eAID_TurnDoFinish,,False,False,)

        If dlgmachine = "5506" Then
            bore.SetAttribute(eAID_BelowCenterline,,True)
        End If

    bore.Operations(1).SetAttribute(eAID_TurnFinishXAllow,,fin_allow,False,)
    bore.Operations(1).SetAttribute(eAID_TurnFinishZAllow,,0,False,)
    bore.Operations(1).SetAttribute(eAID_SkipWallPass,,True,False,)
    bore.Operations(1).SetAttribute(eAID_TurnRoughDOC,,id_doc,False,)
    bore.Operations(1).SetAttribute(eAID_TurnRoughEngageWindrawAtRapid,,True,False,)
    bore.Operations(1).SetAttribute(eAID_MinZBound,,-oal-face_2nd-.05,False,)
    bore.Operations(1).SetAttribute(eAID_TurnEndPt,,endpt,False,)
    bore.Operations(1).OverrideSpeed(boresfm,)
    bore.Operations(1).OverrideFeed(boreipr,)
    bore.Operations(1).OverrideTool(boretool,)
    bore.Name = "rough_bore"

End Sub

Private Sub FinishBore_Op(opt As String)

	Dim startpt As String
	Dim curveL As FMCurvePtList
	Dim curve As FMCurve
	Dim bore As FMBore

	If opt = "front" Then

	    Set curveL = doc.CreateCurvePtList
	        curveL.AddPt(id/2+bore_chfr,0,0)
	        curveL.AddPt(id/2-bore_allow,0,-bore_chfr-bore_allow)
	        curveL.AddPt(id/2-bore_allow,0,-oal-face_2nd-.05)
	    Set curve = curveL.AddCurveFromPtList(False)

	    startpt = "pt(" & id/2+bore_chfr+Sin(45*pi/180)*.1 & ",0," & Cos(45*pi/180)*(.1+toolrad2*2) &")"

	    Set bore = doc.Features.AddBore(curve,,,)

	    If dlgmachine = "5506" Then
	        bore.SetAttribute(eAID_BelowCenterline,,True)
	    End If

	    bore.SetAttribute(eAID_TurnDoRough,,False,False,)
	    bore.Operations(1).SetAttribute(eAID_MinZBound,,-oal-face_2nd-.05,False,)
	    bore.Operations(1).SetAttribute(eAID_TurnStartPt,,startpt,False,)
	    bore.Operations(1).OverrideSpeed(boresfm,)
	    bore.Operations(1).OverrideFeed(boreipr,)
	    bore.Operations(1).OverrideTool(boretool,)
	    If optfeatures = 1 Then
	        bore.Name = "semi_bore"
	    Else
	        bore.Name = "finish_bore"
	    End If

	Else

		Set curveL = doc.CreateCurvePtList
		    curveL.AddPt(id/2+.02,0,0)
		    curveL.AddPt(id/2,0,-.02)
		If optfeatures = 1 Then
		    curveL.AddPt(id/2,0,-oal-.05)
		Else
		    curveL.AddPt(id/2,0,-.03)
		End If
		Set curve = curveL.AddCurveFromPtList(False)

		Set bore = doc.Features.AddBore(curve,,,)

		    If dlgmachine = "5506" Then
		        bore.SetAttribute(eAID_BelowCenterline,,True)
		    End If

		    bore.SetAttribute(eAID_TurnDoRough,,False,False,)
		    bore.Operations(1).OverrideSpeed(boresfm,)
		    bore.Operations(1).OverrideFeed(boreipr,)
		    bore.Operations(1).OverrideTool(boretool,)
		If optfeatures = 1 Then
		    bore.Name = "finish_bore"
		    bore.Operations(1).SetAttribute(eAID_MinZBound,,.05,False,)
		Else
		    startpt = "pt(" & id/2+.02+Sin(45*pi/180)*(.1) & ",0," & Cos(45*pi/180)*(.1+toolrad1*2) &")"
		    bore.Operations(1).SetAttribute(eAID_TurnStartPt,,startpt,False,)
		    bore.Name = "chfr_bore"
		End If

	End If

End Sub

Private Sub CleanOD_Op(opt As String)

	If opt = "front" Then
    	Dim frontCurveL As FMCurvePtList
        Dim frontCurve As FMCurve

    	Set frontCurveL = doc.CreateCurvePtList
    	  frontCurveL.AddPt(stock_od/2+.1-.125,0,face_1st)
    	  frontCurveL.AddPt(stock_od/2+.1-.125,0,-oal/2-.014-.05)
    	Set frontCurve = frontCurveL.AddCurveFromPtList(False)

        Dim endptF As String

    	If stock_od/2+.1-od/2-fin_allow > .125 Then
    	  endptF = "pt(" & stock_od/2+.1-.125+toolrad1+Sin(45*pi/180)*.025 & ",0," & -oal/2-.014-.05-face_1st+Cos(45*pi/180)*.025 &")"
    	Else
    	  endptF = "pt(" & stock_od/2+.1-.125+fin_allow+toolrad1+Sin(45*pi/180)*.025 & ",0," & -oal/2-.014-.05-face_1st+Cos(45*pi/180)*.025 &")"
    	End If

        Dim frontTurn As FMTurn

    	Set frontTurn = doc.Features.AddTurn(frontCurveL,,,)
    	  frontTurn.SetAttribute(eAID_TurnDoFinish, ,False)
    	If stock_od/2+.1-od/2-fin_allow > .125 Then
            frontTurn.Operations(1).SetAttribute(eAID_TurnFinishXAllow,,0,False,)
    	Else
    	  frontTurn.Operations(1).SetAttribute(eAID_TurnFinishXAllow,,fin_allow,False,)
    	End If
    	frontTurn.Operations(1).SetAttribute(eAID_TurnRoughDOC,,.125,False,)
    	frontTurn.Operations(1).SetAttribute(eAID_TurnFinishZAllow,,0,False,)
    	frontTurn.Operations(1).SetAttribute(eAID_TurnClearance,,.2,False,)
        frontTurn.Operations(1).SetAttribute(eAID_SkipWallPass,,True,False,)
        frontTurn.Operations(1).SetAttribute(eAID_TurnRoughEngageWindrawAtRapid,,True,False,)
        frontTurn.Operations(1).SetAttribute(eAID_TurnEndPt,,endptF,False,)
    	frontTurn.Operations(1).OverrideSpeed(turnsfm,)
        frontTurn.Operations(1).OverrideFeed(turnipr,)
        frontTurn.Operations(1).OverrideTool(turntool,)
        frontTurn.Name = "clean_od"
    Else
        Dim backcurveL As FMCurvePtList
        Dim backcurve As FMCurve

        Set backcurveL = doc.CreateCurvePtList
        backcurveL.AddPt(stock_od/2+.1-.125,0,face_1st)
        backcurveL.AddPt(stock_od/2+.1-.125,0,-oal/2-.014+.05)
        Set backcurve = backcurveL.AddCurveFromPtList(False)

        Dim endptB As String

        If stock_od/2+.1-od/2-fin_allow > .125 Then
            endptB = "pt(" & stock_od/2+.1-.125+toolrad1+Sin(45*pi/180)*.025 & ",0," & -oal/2-.014+.05-face_1st+Cos(45*pi/180)*.025 &")"
        Else
            endptB = "pt(" & stock_od/2+.1-.125+fin_allow+toolrad1+Sin(45*pi/180)*.025 & ",0," & -oal/2-.014+.05-face_1st+Cos(45*pi/180)*.025 &")"
        End If

        Dim backTurn As FMTurn

        Set backTurn = doc.Features.AddTurn(backcurve,,,)
        backTurn.SetAttribute(eAID_TurnDoFinish,,False)
        If stock_od/2+.1-od/2-fin_allow > .125 Then
            backTurn.Operations(1).SetAttribute(eAID_TurnFinishXAllow,,0,False,)
        Else
            backTurn.Operations(1).SetAttribute(eAID_TurnFinishXAllow,,fin_allow,False,)
        End If
        backTurn.Operations(1).SetAttribute(eAID_TurnFinishZAllow,,0,False,)
        backTurn.Operations(1).SetAttribute(eAID_TurnClearance,,.2,False,)
        backTurn.Operations(1).SetAttribute(eAID_SkipWallPass,,True,False,)
        backTurn.Operations(1).SetAttribute(eAID_TurnRoughEngageWindrawAtRapid,,True,False,)
        backTurn.Operations(1).SetAttribute(eAID_TurnEndPt,,endptB,False,)
        backTurn.Operations(1).OverrideSpeed(turnsfm,)
        backTurn.Operations(1).OverrideFeed(turnipr,)
        backTurn.Operations(1).OverrideTool(turntool,)
        backTurn.Name = "clean_od_"
    End If

End Sub

Private Sub Eccentric_Op

Dim curveL As FMCurvePtList
Dim curve As FMCurve

Set curveL = doc.CreateCurvePtList
    curveL.AddPt(id/2+.02,0,0)
    curveL.AddPt(id/2,0,-.02)
    curveL.AddPt(id/2,0,-oal-.05)
Set curve = curveL.AddCurveFromPtList(False)

Dim bore As FMBore

Set bore = doc.Features.AddBore(curve,,,)

    If dlgmachine = "5506" Then
        bore.SetAttribute(eAID_BelowCenterline,,True)
    End If

    bore.SetAttribute(eAID_TurnDoRough,,False,False,)
    bore.Operations(1).SetAttribute(eAID_MinZBound,,.05,False,)
    bore.Operations(1).OverrideSpeed(boresfm,)
    bore.Operations(1).OverrideFeed(boreipr,)
    bore.Operations(1).OverrideTool(boretool,)
    bore.Name = "ecc_bore"

End Sub

'*********************************************************************************
'*----------------------------------Functions------------------------------------*
'*********************************************************************************

Function ValidInputDonut() As Boolean

	If program_number = "" Then
		MsgBox("Please Enter Program Number",,"Error")
	ElseIf part_number = "" Then
		MsgBox("Please Enter Part Number",,"Error")
	ElseIf revision = "" Then
		MsgBox("Please Enter Revision Level",,"Error")
	ElseIf operation = "" Then
		MsgBox("Please Enter Operation Number",,"Error")
	ElseIf operation2 = "" Then
		MsgBox("Please Enter 2nd Operation Number",,"Error")
	ElseIf writer = "" Then
		MsgBox("Please Enter Programmer Initials",,"Error")
	ElseIf od = 0 Then
		MsgBox("Please Enter O.D.",,"Error")
	ElseIf id = 0 Then
		MsgBox("Please Enter I.D.",,"Error")
	ElseIf oal = 0 Then
		MsgBox("Please Enter OAL",,"Error")
	ElseIf stock_od = 0 Then
		MsgBox("Please Enter Stock O.D.",,"Error")
	ElseIf stock_id = 0 Then
		MsgBox("Please Enter Stock I.D.",,"Error")
	ElseIf face_1st = 0 Then
		MsgBox("Please Enter 1st Side Face Stock",,"Error")
	ElseIf face_2nd = 0 Then
		MsgBox("Please Enter 2nd Side Face Stock",,"Error")
	ElseIf face_doc = 0 Then
		MsgBox("Please Enter Face Depth of Cut",,"Error")
	ElseIf od_doc = 0 Then
		MsgBox("Please Enter O.D. Depth of Cut",,"Error")
	ElseIf id_doc = 0 Then
		MsgBox("Please Enter I.D. Depth of Cut",,"Error")
	ElseIf Dir("\\ant-fs-01\departments\mfg-eng\compactII\" & dlgmachine & "\" & program_number & ".FM") <> "" Then
		MsgBox("File " & program_number & ".FM Already Exists, File Must Be Renamed.",,"Error")


		ValidInputDonut = False
		Else
		ValidInputDonut = True
	End If

End Function

Function ValidInputTrepan() As Boolean

	If trepan_od = 0 Then
		MsgBox("Please Enter Trepan O.D.",,"Error")
	ElseIf trepan_id = 0 Then
		MsgBox("Please Enter Trepan I.D.",,"Error")
	ElseIf trepan_depth = 0 Then
		MsgBox("Please Enter Trepan Depth",,"Error")
	ElseIf trepan_od>od Then
		MsgBox("Trepan O.D. Must Be Less Than Part O.D.",,"Error")
	ElseIf trepan_id<id Then
		MsgBox("Trepan I.D. Must Be Greater Than Part I.D.",,"Error")


		ValidInputTrepan = False
		Else
		ValidInputTrepan = True
	End If

End Function
