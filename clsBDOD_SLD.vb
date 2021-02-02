Imports System
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Imports BDOD_CAD.clsBDOD_CADEntities
Imports BDOD_CAD.clsBDOD_CADSetting
Imports BDOD_CAD.clsBlock_Jigs
Imports BDOD_CAD.clsSLD_Table

Imports Excel = Microsoft.Office.Interop.Excel
'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.DatabaseServices

Public Class clsBDOD_SLD
	Private Shared SLDdtaTable1, SLDdtaTable2 As DataTable
	Private Shared FolderPath As String
	Private Shared FileName As String = String.Empty
	Private Shared dicPilotJoint As New Dictionary(Of String, List(Of String))
	Private Shared dicPilotJointCable_Table As New Dictionary(Of String, List(Of String))

	Private Shared DSS_SearchValue As String
	Private Shared DSS_SequenceValue As String
	Private Shared PON_Patch_DSS_SearchValue As String
	Private Shared PON_Patch_DSS_SequenceValue As String
	Private Shared MPT_SearchValue As String
	Private Shared MPTAssociation As String = String.Empty
	Private Shared CTLAssociation As String = String.Empty

	Private Shared MPTAssocTempList As New List(Of String)
	Private Shared dicMPTAssociation As New Dictionary(Of String, List(Of String))
	Private Shared dicMPTCable_Table As New Dictionary(Of String, List(Of String))
	Private Shared FANAssociation As String = String.Empty
	Private Shared dicDJL_Cable_IN As New Dictionary(Of String, List(Of String))
	Private Shared PrimaryTraceFile As String = String.Empty
	Private Shared dicCable_Table As New Dictionary(Of String, List(Of String))
	Private Shared dctSLDFAN_AttTAG_AttVal As New Dictionary(Of String, String)
	Private Shared dctSLDDJL_AttTAG_AttVal As New Dictionary(Of String, String)
	Private Shared dctSLDBMPT_AttTAG_AttVal As New Dictionary(Of String, String)
	Private Shared dctSLDFJL_AttTAG_AttVal As New Dictionary(Of String, String)
	Private Shared blnLTCFound As Boolean = False
	Private Shared PilotJoint As String = String.Empty
	Private Shared FJLSplitter As String = String.Empty
	Private Shared ValueAtSequence As String = String.Empty
	Private Shared PortAtSequence As String = String.Empty
	Private Shared dicPONPatch As New Dictionary(Of String, String)
	'****************************LFN VARIABLES***********************
	Private Shared LFN_FJL_Name As String
	Private Shared LFN_FJL_ADDRESS As String
	Private Shared LFN_FJL_SUBURB_STATE As String

	Private Shared LFN_BJL_Name As Dictionary(Of String, List(Of String))
	Private Shared LFN_BJL_ADDRESS As String
	Private Shared LFN_BJL_SUBURB_STATE As String

	Private Shared LFN_PCD_Name As String
	Private Shared LFN_PCD_ADDRESS As String
	Private Shared LFN_PCD_SUBURB_STATE As String

	Private Shared LFN_SMP_Name As String
	Private Shared LFN_SMP_ADDRESS As String
	Private Shared LFN_SMP_SUBURB_STATE As String
	Private Shared LFN_SMP_TYPE As String
	Private Shared LFN_SMP_RATIO As String
	Private Shared LFN_SMP_SPLITTER_IN As String
	Private Shared LFN_SMP_SPLITTER_OUT As String

	Private Shared LFN_CTL_NAME As String
	Private Shared LFN_CTL_ADDRESS As String
	Private Shared LFN_CTL_SUBURB_STATE As String

	Private Shared LFN_FSL_NAME As String
	Private Shared LFN_FSL_Fibre As String
	Private Shared LFN_FSL_Fibre_Sequence As String
	Private Shared LFN_FSL_LENGTH As String

	Private Shared LFN_SSS_NAME As String
	Private Shared LFN_SSS_Fibre As String
	Private Shared LFN_SSS_Fibre_Sequence As String
	Private Shared LFN_SSS_LENGTH As String

	Private Shared LFN_SDS_NAME As String
	Private Shared LFN_SDS_Fibre As String
	Private Shared LFN_SDS_Fibre_Sequence As String
	Private Shared LFN_SDS_LENGTH As String

	Private Shared LFN_SPLITTER_Name As String
	Private Shared LFN_SPLITTER_Branch As String

	Private Shared LFN_PIC_NAME As String
	Private Shared LFN_PIC_Fibre As String
	Private Shared LFN_PIC_LENGTH As String
	Private Shared LFN_PIC_SEQUENCE As String

	Private Shared LFN_ICD_Name As String
	Private Shared LFN_ICD_ADDRESS As String
	Private Shared LFN_ICD_SUBURB_STATE As String

	Private Shared LFN_PDC_NAME As String
	Private Shared LFN_PDC_Fibre As String
	Private Shared LFN_PDC_LENGTH As String
	Private Shared LFN_PDC_SEQUENCE As String

	Private Shared LFN_NTD_Name As String
	Private Shared LFN_NTD_ADDRESS As String
	Private Shared LFN_NTD_SUBURB_STATE As String

	Private Shared LFN_CTL_AT_Cable_Side_CSV As String
	Private Shared LFN_SMP_AT_Cable_Side_CSV As String
	Private Shared LFN_SPLITTER_AT_Cable_Side_CSV As String
	Private Shared LFN_CTL_AT_DROP_Side_CSV As String
	Private Shared LFN_SMP_AT_DROP_Side_CSV As String
	Private Shared dicLFN_Splice As List(Of String)
	Private Shared dctLFN_SLD_CTL_AttTAG_AttVal As New Dictionary(Of String, String)
	Private Shared dctLFN_SLD_ICD_AttTAG_AttVal As New Dictionary(Of String, String)
	Private Shared dctLFN_SLD_NTD_AttTAG_AttVal As New Dictionary(Of String, String)
	Private Shared dctLFN_SLD_BJL_AttTAG_AttVal As New Dictionary(Of String, String)
	Private Shared dctLFN_SLD_SMP_AttTAG_AttVal As New Dictionary(Of String, String)
	Private Shared dctLFN_SLD_PCD_AttTAG_AttVal As New Dictionary(Of String, String)
	Private Shared dicLFN_CTL_Splice As Dictionary(Of String, List(Of String))
	Private Shared dicLFN_ICD_Splice As Dictionary(Of String, List(Of String))
	Private Shared dicLFN_BJL_Splice As Dictionary(Of String, List(Of String))
	Private Shared dicLFN_SMP_Splice As Dictionary(Of String, List(Of String))
	Private Shared dicLFN_PCD_Splice As Dictionary(Of String, List(Of String))

	Private Shared blkLFCTLInsPt As Point3d
	Private Shared RepeatCount As String

	Private Shared Sub UpdatePONPatchTemplate()
		getAutoCADActiveDocument()
		Dim xlSheetName As String = "Template"
		Try
			xlApp = New Microsoft.Office.Interop.Excel.Application()
			xlWorkBook = xlApp.Workbooks.Open(BDOD_Temp_Path + PONTemplate)
			xlApp.DisplayAlerts = False
			xlApp.Visible = False
			xlWorkSheet = xlWorkBook.Sheets.Item(xlSheetName)
			xlWorkSheet.Select()
			xlCells = xlWorkSheet.Cells(7, 3)
			xlCells.Value = "NBN-Select"
			xlCells = xlWorkSheet.Cells(8, 3)
			xlCells.Value = ApplicationName.ToString
			xlCells = xlWorkSheet.Cells(9, 3)
			'If ProjectInfo.Count = 24 Then
			xlCells.Value = DesignProjectID.ToString
			'Else
			'	xlCells.Value = ProjectInfo(16).ToString
			'End If
			xlCells = xlWorkSheet.Cells(10, 3)
			xlCells.Value = "FTTP"
			xlCells = xlWorkSheet.Cells(11, 3)
			xlCells.Value = "New design"
			xlCells = xlWorkSheet.Cells(12, 3)
			xlCells.Value = DesignFSA.ToString + "-" + DesignFSAM.ToString
			xlCells = xlWorkSheet.Cells(13, 3)
			xlCells.Value = FANAssociation
			xlCells = xlWorkSheet.Cells(20, 2)

			If CTLAssociation <> String.Empty Then
				xlCells.Value = CTLAssociation
			ElseIf PilotJoint <> String.Empty Then
				xlCells.Value = PilotJoint

			End If
			xlCells = xlWorkSheet.Cells(20, 3)
			xlCells.Value = FJLSplitter
			xlCells = xlWorkSheet.Cells(20, 4)
			xlCells.Value = "N/A"
			xlCells = xlWorkSheet.Cells(20, 5)
			xlCells.Value = PON_Patch_DSS_SearchValue + Chr(32) + "Fibre" + Chr(32) + PON_Patch_DSS_SequenceValue ' DSS_SearchValue + Chr(32) + "Fibre" + Chr(32) + DSS_SequenceValue
			xlCells = xlWorkSheet.Cells(20, 6)
			xlCells.Value = ValueAtSequence + Chr(32) + "PORT" + Chr(32) + PortAtSequence
			Dim TempCurDrawingName As String = Path.GetFileNameWithoutExtension(CurDrawingName)
			Dim PONPTachFileName As String = TempCurDrawingName.Substring(0, 19) + Chr(45) + "PON PATCHING REQ" + Chr(45) + TempCurDrawingName.Substring(TempCurDrawingName.Length - 4)

			xlWorkBook.SaveAs(Path.Combine(CurDrawingPath, PONPTachFileName + ".xlsx"), FileFormat:=51)
			xlWorkBook.Close()
			xlApp.Quit()

			clsBOMBOQ.ReleaseComObject(xlCells)
			clsBOMBOQ.ReleaseComObject(xlWorkSheet)
			clsBOMBOQ.ReleaseComObject(xlWorkBook)
			clsBOMBOQ.ReleaseComObject(xlApp)

			xlCells = Nothing
			xlWorkSheet = Nothing
			xlWorkBook = Nothing
			xlApp = Nothing

			dicPONPatch.Clear()
			dicPONPatch = PopulateBlockAttributes(blkPONPATCH)
			SetBlockAttributeValues(dicPONPatch)
			If CTLAssociation <> String.Empty Then
				dicPONPatch("ENTITY_HOUSING_1") = CTLAssociation
			ElseIf PilotJoint <> String.Empty Then
				dicPONPatch("ENTITY_HOUSING_1") = PilotJoint
			End If
			dicPONPatch("SPLITTER_1") = FJLSplitter
			dicPONPatch("DFN_PORT_1") = "N/A"
			dicPONPatch("DSS_CABLE_1") = PON_Patch_DSS_SearchValue + Chr(32) + "Fibre" + Chr(32) + PON_Patch_DSS_SequenceValue
			dicPONPatch("HDODF_TERMINATION_PORT_1") = ValueAtSequence + Chr(32) + "PORT" + Chr(32) + PortAtSequence
			Dim AttTAG As New List(Of String)
			Dim AttVal As New List(Of String)
			AttTAG = dicPONPatch.Keys.ToList()
			AttVal = dicPONPatch.Values.ToList()

            Dim Prompt1 As String = "Select Insertion Point for PON Patch Table: "
            Dim Prompt2 As String = String.Empty
			Dim CurOSMODE = GETCADSystemVariable("OSMODE")
			Dim CurLayer As String = GETCADSystemVariable("CLAYER")
			SETCADSystemVariable("OSMODE", 0)
			SETCADSystemVariable("CLAYER", "0")
			Dim PONPatchInsPt As New Point3d
			Dim PONPatchRotation As New Double
			Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
			InsertBDOD_Block(blkPONPATCH, AttVal, Prompt1, Prompt2, PONPatchInsPt, PONPatchRotation)
			SETCADSystemVariable("OSMODE", CurOSMODE)
			SETCADSystemVariable("CLAYER", CurLayer)
		Catch ex As System.Exception
			MsgBox(ex.Message)
		End Try
	End Sub
	Private Shared Sub StartSLD_IN_CAD()
		Dim CurOSMODE = GETCADSystemVariable("OSMODE")
		Dim CurLayer As String = GETCADSystemVariable("CLAYER")
		SETCADSystemVariable("OSMODE", 0)
		SETCADSystemVariable("CLAYER", layOtherProjectLayer)
		Dim blkType As String = String.Empty
		Dim SLDFANInsPt As New Point3d
		Dim SLDFANRotation As New Double
		Dim SLDFANAnnoOffset As New Point3d
		Dim AttTAG As New List(Of String)
		Dim AttVal As New List(Of String)
		dctSLDFAN_AttTAG_AttVal.Clear()
		dctSLDFAN_AttTAG_AttVal = PopulateBlockAttributes(blkSLDFAN)
		SetBlockAttributeValues(dctSLDFAN_AttTAG_AttVal)

		dctSLDFAN_AttTAG_AttVal("NAME") = FANAssociation
		dctSLDFAN_AttTAG_AttVal("PROJECT_ID") = DesignProjectID
		dctSLDFAN_AttTAG_AttVal("PROJECT_NAME") = ApplicationName.ToString
		dctSLDFAN_AttTAG_AttVal("NETWORK_TYPE") = "LFN"

		AttTAG = dctSLDFAN_AttTAG_AttVal.Keys.ToList()
		AttVal = dctSLDFAN_AttTAG_AttVal.Values.ToList()

		Dim Prompt1 As String = "Select Insertion Point for SLD: "
		Dim Prompt2 As String = String.Empty
		SLDFANAnnoOffset = New Point3d(SLDFANInsPt.X, SLDFANInsPt.Y + 1.3, SLDFANInsPt.Z)
		InsertBDOD_Block(blkSLDFAN, AttVal, Prompt1, Prompt2, SLDFANInsPt, SLDFANRotation)
		Dim DJLInsPT As Point3d = New Point3d(SLDFANInsPt.X + 12.5, SLDFANInsPt.Y, SLDFANInsPt.Z)
		dctSLDDJL_AttTAG_AttVal.Clear()
		dctSLDDJL_AttTAG_AttVal = PopulateBlockAttributes(blkSLDDJL)
		SetBlockAttributeValues(dctSLDDJL_AttTAG_AttVal)
		For Each KVP As KeyValuePair(Of String, List(Of String)) In dicDJL_Cable_IN
			dctSLDDJL_AttTAG_AttVal("SLD_DJL_NAME1") = KVP.Key
			dctSLDDJL_AttTAG_AttVal("SLD_DJL_NAME2") = KVP.Key
			dctSLDDJL_AttTAG_AttVal("SLD_DJL_NAME3") = KVP.Key
			dctSLDDJL_AttTAG_AttVal("PROJECT_ID") = KVP.Value(0).ToString
			dctSLDDJL_AttTAG_AttVal("PROJECT_NAME") = KVP.Value(1).ToString
			dctSLDDJL_AttTAG_AttVal("NETWORK_TYPE") = KVP.Value(2).ToString
			dctSLDDJL_AttTAG_AttVal("ADDRESS1") = KVP.Value(3).ToString
			dctSLDDJL_AttTAG_AttVal("ADDRESS2") = KVP.Value(3).ToString
			dctSLDDJL_AttTAG_AttVal("ADDRESS3") = KVP.Value(3).ToString
			dctSLDDJL_AttTAG_AttVal("SLD_DSS_NAME1") = KVP.Value(4).ToString
			dctSLDDJL_AttTAG_AttVal("SLD_DSS_NAME2") = KVP.Value(4).ToString
			dctSLDDJL_AttTAG_AttVal("SLD_DSS_NAME3") = KVP.Value(4).ToString
			dctSLDDJL_AttTAG_AttVal("SLD_DSS_FIBRE1") = KVP.Value(5).ToString
			dctSLDDJL_AttTAG_AttVal("SLD_DSS_FIBRE2") = KVP.Value(5).ToString
			dctSLDDJL_AttTAG_AttVal("SLD_DSS_FIBRE3") = KVP.Value(5).ToString
			dctSLDDJL_AttTAG_AttVal("SLD_DSS_LENGTH1") = KVP.Value(6).ToString
			dctSLDDJL_AttTAG_AttVal("SLD_DSS_LENGTH2") = KVP.Value(6).ToString
			dctSLDDJL_AttTAG_AttVal("SLD_DSS_LENGTH3") = KVP.Value(6).ToString
			AttTAG.Clear()
			AttVal.Clear()
			AttTAG = dctSLDDJL_AttTAG_AttVal.Keys.ToList()
			AttVal = dctSLDDJL_AttTAG_AttVal.Values.ToList()
			InsertBlockWithAttributes(ProjectType, blkSLDDJL, "0", blkType, DJLInsPT, SLDFANRotation, DJLInsPT, AttTAG, AttVal)
			Start_SLD_TABLE(dicCable_Table, KVP.Key, DJLInsPT)
			DJLInsPT = New Point3d(DJLInsPT.X + 71.0, DJLInsPT.Y, DJLInsPT.Z)
		Next
		Dim CurLineType As String = GETCADSystemVariable("CELTYPE")
		If dicMPTAssociation.Count > 0 Then
			dctSLDBMPT_AttTAG_AttVal.Clear()
			dctSLDBMPT_AttTAG_AttVal = PopulateBlockAttributes(blkSLDBMPT)
			SetBlockAttributeValues(dctSLDBMPT_AttTAG_AttVal)
			'For Each KVP As KeyValuePair(Of String, List(Of String)) In dicMPTAssociation
			dctSLDBMPT_AttTAG_AttVal("SLD_BMPT_NAME1") = dicMPTAssociation.Keys(0).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_BMPT_NAME2") = dicMPTAssociation.Keys(0).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_BMPT_NAME3") = dicMPTAssociation.Keys(0).ToString
			dctSLDBMPT_AttTAG_AttVal("TYPE1") = MPTAssocTempList(1).ToString
			dctSLDBMPT_AttTAG_AttVal("TYPE2") = MPTAssocTempList(1).ToString
			dctSLDBMPT_AttTAG_AttVal("TYPE3") = MPTAssocTempList(1).ToString
			dctSLDBMPT_AttTAG_AttVal("PROJECT_ID") = MPTAssocTempList(2).ToString
			dctSLDBMPT_AttTAG_AttVal("PROJECT_NAME") = MPTAssocTempList(3).ToString
			dctSLDBMPT_AttTAG_AttVal("NETWORK_TYPE") = MPTAssocTempList(4).ToString
			dctSLDBMPT_AttTAG_AttVal("ADDRESS1") = MPTAssocTempList(5).ToString
			dctSLDBMPT_AttTAG_AttVal("ADDRESS2") = MPTAssocTempList(5).ToString
			dctSLDBMPT_AttTAG_AttVal("ADDRESS3") = MPTAssocTempList(5).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_NAME1") = MPTAssocTempList(0).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_NAME2") = MPTAssocTempList(0).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_NAME3") = MPTAssocTempList(0).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_NAME4") = MPTAssocTempList(0).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_NAME5") = MPTAssocTempList(0).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_NAME6") = MPTAssocTempList(0).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_FIBRE1") = MPTAssocTempList(6).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_FIBRE2") = MPTAssocTempList(6).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_FIBRE3") = MPTAssocTempList(6).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_FIBRE4") = MPTAssocTempList(6).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_FIBRE5") = MPTAssocTempList(6).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_FIBRE6") = MPTAssocTempList(6).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_LENGTH1") = MPTAssocTempList(7).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_LENGTH2") = MPTAssocTempList(7).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_LENGTH3") = MPTAssocTempList(7).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_LENGTH4") = MPTAssocTempList(7).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_LENGTH5") = MPTAssocTempList(7).ToString
			dctSLDBMPT_AttTAG_AttVal("SLD_DSS_LENGTH6") = MPTAssocTempList(7).ToString
			AttTAG.Clear()
			AttVal.Clear()
			AttTAG = dctSLDBMPT_AttTAG_AttVal.Keys.ToList()
			AttVal = dctSLDBMPT_AttTAG_AttVal.Values.ToList()
			InsertBlockWithAttributes(ProjectType, blkSLDBMPT, "0", blkType, DJLInsPT, SLDFANRotation, DJLInsPT, AttTAG, AttVal)
			CurLineType = GETCADSystemVariable("CELTYPE")

			SETCADSystemVariable("CLAYER", "PROPOSED")
			SETCADSystemVariable("CELTYPE", "continuous")
			Start_SLD_TABLE(dicMPTCable_Table, dicMPTCable_Table.Keys(0), DJLInsPT)

			SETCADSystemVariable("OSMODE", CurOSMODE)
			SETCADSystemVariable("CLAYER", CurLayer)
			SETCADSystemVariable("CELTYPE", CurLineType)
		End If
		dctSLDFJL_AttTAG_AttVal.Clear()
		dctSLDFJL_AttTAG_AttVal = PopulateBlockAttributes(blkSLDFJL)
		SetBlockAttributeValues(dctSLDFJL_AttTAG_AttVal)

		Dim PilotList As List(Of String) = dicPilotJoint.Values(0).ToList
		'For Each KVP As KeyValuePair(Of String, List(Of String)) In dicPilotJoint
		'	PilotList = KVP.Value.ToList
		'Next
		If dicMPTAssociation.Count > 0 Then
			DJLInsPT = New Point3d(DJLInsPT.X + 71.0, DJLInsPT.Y, DJLInsPT.Z)
		End If
		dctSLDFJL_AttTAG_AttVal("SLD_FJL_NAME1") = dicPilotJoint.Keys(0).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FJL_NAME2") = dicPilotJoint.Keys(0).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FJL_NAME3") = dicPilotJoint.Keys(0).ToString
		dctSLDFJL_AttTAG_AttVal("PROJECT_ID") = PilotList(0).ToString
		dctSLDFJL_AttTAG_AttVal("PROJECT_NAME") = PilotList(1).ToString
		dctSLDFJL_AttTAG_AttVal("NETWORK_TYPE") = PilotList(2).ToString
		dctSLDFJL_AttTAG_AttVal("ADDRESS1") = PilotList(3).ToString
		dctSLDFJL_AttTAG_AttVal("ADDRESS2") = PilotList(3).ToString
		dctSLDFJL_AttTAG_AttVal("ADDRESS3") = PilotList(3).ToString
		dctSLDFJL_AttTAG_AttVal("SUBURB_STATE1") = PilotList(4).ToString
		dctSLDFJL_AttTAG_AttVal("SUBURB_STATE2") = PilotList(4).ToString
		dctSLDFJL_AttTAG_AttVal("SUBURB_STATE3") = PilotList(4).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME1") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME2") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME3") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME4") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME5") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME6") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME7") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME8") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME9") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME10") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME11") = PilotList(5).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_NAME12") = PilotList(5).ToString

		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE1") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE2") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE3") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE4") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE5") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE6") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE7") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE8") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE9") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE10") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE11") = PilotList(6).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_FIBRE12") = PilotList(6).ToString

		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH1") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH2") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH3") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH4") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH5") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH6") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH7") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH8") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH9") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH10") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH11") = PilotList(7).ToString
		dctSLDFJL_AttTAG_AttVal("SLD_FSD_LENGTH12") = PilotList(7).ToString

		AttTAG.Clear()
		AttVal.Clear()
		AttTAG = dctSLDFJL_AttTAG_AttVal.Keys.ToList()
		AttVal = dctSLDFJL_AttTAG_AttVal.Values.ToList()
		InsertBlockWithAttributes(ProjectType, blkSLDFJL, "0", blkType, DJLInsPT, SLDFANRotation, DJLInsPT, AttTAG, AttVal)

		SETCADSystemVariable("CLAYER", "PROPOSED")
		SETCADSystemVariable("CELTYPE", "continuous")
		If CTLAssociation <> String.Empty Then
			Dim TempPilotJointTable As New Dictionary(Of String, List(Of String))
			Dim TempPilotKey As String = dicPilotJointCable_Table.Keys(0).ToString
			Dim TempPilotValue As String = String.Empty
			Dim TempPilotlist As List(Of String) = dicPilotJointCable_Table.Values(0).ToList()
			If TempPilotlist.Count > 1 Then
				TempPilotValue = TempPilotlist(0).Substring(0, TempPilotlist(0).IndexOf("=") - 1) + TempPilotlist(1).Substring(TempPilotlist(0).IndexOf("=") + 1)
			ElseIf TempPilotlist.Count = 1 Then
				TempPilotValue = TempPilotlist(0).ToString 'TempPilotlist(0).Substring(0, TempPilotlist(0).IndexOf("=") - 1) + TempPilotlist(1).Substring(TempPilotlist(0).IndexOf("=") + 1)
			End If
			dicPilotJointCable_Table.Clear()
			dicPilotJointCable_Table.Add(TempPilotKey, New List(Of String)({TempPilotValue}))
		End If
		Start_SLD_TABLE(dicPilotJointCable_Table, dicPilotJointCable_Table.Keys(0), DJLInsPT)
		SETCADSystemVariable("OSMODE", CurOSMODE)
		SETCADSystemVariable("CLAYER", CurLayer)
		SETCADSystemVariable("CELTYPE", CurLineType)
	End Sub
	Public Shared Function GetCellValue(ByVal wbPart As WorkbookPart, ByVal wsPart As WorksheetPart, ByVal addressName As String) As String
		Dim value As String = Nothing
		Dim theCell As Cell = wsPart.Worksheet.Descendants(Of Cell)().Where(Function(c) c.CellReference = addressName).FirstOrDefault()
		If theCell IsNot Nothing Then
			value = theCell.InnerText
			If theCell.DataType IsNot Nothing Then
				Select Case theCell.DataType.Value
					Case CellValues.SharedString
						Dim stringTable = wbPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault()
						If stringTable IsNot Nothing Then
							value = stringTable.SharedStringTable.ElementAt(Integer.Parse(value)).InnerText
						End If
						Exit Select
					Case CellValues.[Boolean]
						Select Case value
							Case "0"
								value = "FALSE"
								Exit Select
							Case Else
								value = "TRUE"
								Exit Select
						End Select
						Exit Select
				End Select
			End If
		End If
		If addressName = "I1" AndAlso value = String.Empty Then
			value = "Branch"
		End If
		Return value
	End Function
	Public Shared Function ConvertCSVtoDataTable(ByVal strFilePath As String) As DataTable
		Dim COLCount As New List(Of Integer)
		Dim Fields As String()
		Dim Lines As String() = File.ReadAllLines(strFilePath)
		For Each Line In Lines
			Fields = Line.Split(New Char() {","c})
			COLCount.Add(Fields.Length)
		Next
		Dim Cols As Integer = COLCount.Max
		Fields = Lines(0).Split(New Char() {","c})
		Dim dt As New DataTable()
		Dim i As Integer = 0
		While i < Cols
			If i < Fields.Length AndAlso Fields(i).ToString <> String.Empty Then
				dt.Columns.Add(Fields(i).ToLower(), GetType(String))
			ElseIf i < Fields.Length AndAlso Fields(i).ToString = String.Empty Then
				dt.Columns.Add("")
			ElseIf i >= Fields.Length Then 'AndAlso Fields(i).ToString = String.Empty Then
				dt.Columns.Add("")
			End If
			i += 1
		End While
		Using sr As New StreamReader(strFilePath)
			While Not sr.EndOfStream
				Fields = sr.ReadLine().Split(","c)
				Dim dr As DataRow = dt.NewRow()
				i = 0
				While i < Cols
					If i < Fields.Length AndAlso Fields(i).ToString <> String.Empty Then
						dr(i) = Fields(i)
					ElseIf i >= Fields.Length Then
						dr(i) = ""
					End If
					i += 1
				End While
				dt.Rows.Add(dr)
			End While
		End Using
		Return dt
	End Function

	Private Shared Sub PopulateCable_IN()
		Dim FileEntries As String() = Directory.GetFiles(FolderPath, "*.csv")

		Dim DJLNAME As String = String.Empty
		Dim DSSNAME As String = String.Empty
		Dim DJLAddress As String = String.Empty
		Dim DSSFibre As String = String.Empty
		Dim DSSLength As String = String.Empty
		Dim SheetsUsed As New List(Of String)
		Dim ReadCableIN As String = String.Empty
		Dim ReadCableOUT As String = String.Empty
		Dim TempdicDJL_Cable_IN As New Dictionary(Of String, List(Of String))
		'dicDJL_Cable_IN.Clear()
		Dim TempdicCable_Table As New Dictionary(Of String, List(Of String))
		Dim FibreValue As String = String.Empty
		dicCable_Table.Clear()
		dicMPTAssociation.Clear()
		'Dim frmShowProess As New BDOD_Progress
		'Dim ProgressValue As Integer = 2
		'frmShowProess.ColorProgressBar1.Minimum = 1
		'frmShowProess.ColorProgressBar1.Maximum = dicDJL_DSS.Count + 1
		'For Each KVP As KeyValuePair(Of String, String) In dicDJL_DSS
		'frmShowProess.TopMost = True
		'frmShowProess.ColorProgressBar1.Value = ProgressValue
		'frmShowProess.lblFilesCount.Text = "FILE COUNT: " + Chr(32) + ProgressValue.ToString + Chr(32) + "of" + Chr(32) + dicDJL_DSS.Count.ToString
		'frmShowProess.lblFileName.Text = "File Name:" + Chr(32) + KVP.Key.ToString
		'frmShowProess.Show()
		'frmShowProess.ColorProgressBar1.Refresh()
		'frmShowProess.Refresh()
		'DJLNAME = KVP.Key
		'    DSSNAME = KVP.Value
		'Dim SheetName As String = String.Empty
		'Dim CellAddress As String = String.Empty
		SLDdtaTable1 = New DataTable
		'************************TEST CSV****************
		For Each FileName As String In FileEntries
			SLDdtaTable1 = ConvertCSVtoDataTable(FileName)
			DSSNAME = String.Empty
			DJLAddress = String.Empty
			DSSFibre = String.Empty
			DSSLength = String.Empty
			Dim SpliceRow As Integer = 0
			FibreValue = String.Empty
			Dim TempList2 As New List(Of String)
			For Each Row As DataRow In SLDdtaTable1.Rows
				Select Case True
					Case Row.ItemArray(0).ToString = "Splice name:"
						DJLNAME = Row.ItemArray(1).ToString
					Case Row.ItemArray(0).ToString = "Address:"
						DJLAddress = Row.ItemArray(1).ToString
					Case Row.ItemArray(0).ToString.Contains("DSS") AndAlso DSSNAME = String.Empty
						Dim ADASearch As String = ADA_Prefix.Substring(0, 4)
						DSSNAME = Row.ItemArray(0).ToString.Substring(Row.ItemArray(0).ToString.IndexOf(ADASearch), 18)
						DSSFibre = Row.ItemArray(1).ToString + "F"
						DSSLength = String.Format("{0:F1}", Double.Parse(Row.ItemArray(6).ToString)) + "m"
						If Row.ItemArray(6).ToString = 0 Then
							DSSLength = String.Format("{0:F1}", Double.Parse(Row.ItemArray(5).ToString)) + "m"
						End If
						For Each KVP As KeyValuePair(Of String, List(Of String)) In dicDJL_Cable_IN
							If KVP.Key = DJLNAME Then
								KVP.Value.Add(DSSLength)
								Exit For
							End If
						Next

					Case Row.ItemArray(0).ToString = "Splicing summary"
						SpliceRow = SLDdtaTable1.Rows.IndexOf(Row)
					Case SpliceRow > 0 AndAlso Row.ItemArray(0).ToString <> String.Empty
						If Row.ItemArray(0).ToString = "Unconnected" Then
							ReadCableIN = "IDLE"
						ElseIf Row.ItemArray(0).ToString.Contains("DSS") Then
							ReadCableIN = Row.ItemArray(0).ToString
							FibreValue = Chr(70) + ReadCableIN.Substring(ReadCableIN.IndexOf(Chr(32)) + 1)
							ReadCableIN = ReadCableIN.Substring(0, 4) + Chr(45) + ReadCableIN.Substring(4, 2) + Chr(45) + ReadCableIN.Substring(6, 2) + Chr(45) + ReadCableIN.Substring(8, 3) + Chr(45) + ReadCableIN.Substring(11, 3) + Chr(32) + FibreValue
						End If
						If Row.ItemArray(2).ToString = "Unconnected" Then
							ReadCableOUT = "IDLE"
							'TempList2.Add("IDLE")
						ElseIf Row.ItemArray(2).ToString.Contains("DSS") Or Row.ItemArray(2).ToString.Contains("FSD") Or Row.ItemArray(2).ToString.Contains("HSD") Then
							ReadCableOUT = Row.ItemArray(2).ToString
							FibreValue = Chr(70) + ReadCableOUT.Substring(ReadCableOUT.IndexOf(Chr(32)) + 1)
							ReadCableOUT = ReadCableOUT.Substring(0, 4) + Chr(45) + ReadCableOUT.Substring(4, 2) + Chr(45) + ReadCableOUT.Substring(6, 2) + Chr(45) + ReadCableOUT.Substring(8, 3) + Chr(45) + ReadCableOUT.Substring(11, 3) + Chr(32) + FibreValue
						End If
					Case Row.ItemArray(3).ToString = MPT_SearchValue AndAlso Row.ItemArray(0).ToString.Contains(DSS_SearchValue)
						DSSLength = String.Format("{0:F1}", Double.Parse(Row.ItemArray(6).ToString)) + "m"
						MPTAssocTempList.Add(DSSLength)
						dicMPTAssociation.Add(MPTAssociation, MPTAssocTempList)
				End Select
				If ReadCableIN <> String.Empty AndAlso ReadCableOUT <> String.Empty Then
					TempList2.Add(ReadCableIN + Chr(32) + Chr(61) + Chr(32) + ReadCableOUT)
					ReadCableIN = String.Empty
					ReadCableOUT = String.Empty
				End If
			Next

			'TempdicDJL_Cable_IN.Add(DJLNAME, New List(Of String)({DJLAddress, DSSNAME, DSSFibre, DSSLength}))
			'TempdicCable_Table.Add(DJLNAME, TempList2)
			dicCable_Table.Add(DJLNAME, TempList2)
		Next
		'Dim dicDJL_CableINKeys As List(Of String) = TempdicDJL_Cable_IN.Keys.ToList
		'dicDJL_CableINKeys.Sort()
		'For Each ReadKey As String In dicDJL_CableINKeys
		'	'For Each KVP As KeyValuePair(Of String, List(Of String)) In TempdicDJL_Cable_IN
		'	'	If ReadKey = KVP.Key Then
		'	'		dicDJL_Cable_IN.Add(KVP.Key, KVP.Value)
		'	'	End If
		'	'Next
		'	For Each KVP As KeyValuePair(Of String, List(Of String)) In TempdicCable_Table
		'		If ReadKey = KVP.Key Then
		'			dicCable_Table.Add(KVP.Key, KVP.Value)
		'		End If
		'	Next
		'Next
		'**************************************
		If dicDJL_Cable_IN.Count <> dicCable_Table.Count Then
			Dim CSVKeysMissing = dicDJL_Cable_IN.Keys.Except(dicCable_Table.Keys)
			Dim DJLKeysMissing = dicCable_Table.Keys.Except(dicDJL_Cable_IN.Keys)
			If CSVKeysMissing IsNot Nothing Then
				Throw New System.Exception("CSV FILE MISSING FOR: " + CSVKeysMissing(0).ToString)
			ElseIf DJLKeysMissing IsNot Nothing Then
				Throw New System.Exception("TRACE REPORT DOES NOT HAVE: " + DJLKeysMissing(0).ToString)
			End If
		End If
	End Sub
	Private Shared Sub OpenAndReadTraceReeport()
		'dicDJL_DSS.Clear()
		dicDJL_Cable_IN.Clear()
		dicPilotJoint.Clear()
		dicPilotJointCable_Table.Clear()
		dicMPTAssociation.Clear()
		dicMPTCable_Table.Clear()
		'Dim valueToSearch As String = "FJL"
		Dim DJLName As String = String.Empty
		Dim DJLAddress As String = String.Empty
		Dim DSSName As String = String.Empty
		PilotJoint = String.Empty
		Dim PilotAddress As String = String.Empty
		Dim FSDNAME As String = String.Empty
		Dim FSDFibre As String = String.Empty
		Dim FSDFibreCount As String = String.Empty
		Dim FSDFibreLength As String = String.Empty
		Dim FJLSpliceIN As String = String.Empty
		Dim FJLSpliceOUT As String = String.Empty
		Dim FSLNAME As String = String.Empty
		Dim DSSFibre As String = String.Empty
		CTLAssociation = String.Empty
		PortAtSequence = String.Empty
		ValueAtSequence = String.Empty
		DSS_SearchValue = String.Empty
		MPTAssociation = String.Empty
		MPT_SearchValue = String.Empty
		DSS_SequenceValue = String.Empty
		PON_Patch_DSS_SearchValue = String.Empty
		PON_Patch_DSS_SequenceValue = String.Empty
		For Each Row As DataRow In SLDdtaTable1.Rows
			Dim ValueAtZero As Object = Row.ItemArray(0)
			Dim ValueAtOne As Object = Row.ItemArray(1)
			Dim RowNum As Integer = 5
			Select Case True
				Case ValueAtZero IsNot DBNull.Value AndAlso (ValueAtZero.ToString.Contains("-LTC-"))
					blnLTCFound = True
				Case ValueAtZero IsNot DBNull.Value AndAlso ((ValueAtZero.ToString.Contains("-FJL-") Or ValueAtZero.ToString.Contains("-FDH-")))
					If PilotJoint = String.Empty AndAlso PilotAddress = String.Empty Then
						PilotJoint = ValueAtZero.ToString.Substring(ValueAtZero.ToString.Length - 18)
						PilotAddress = Row.ItemArray(1).ToString
					End If
				Case ValueAtZero IsNot DBNull.Value AndAlso (ValueAtZero.ToString.Contains("-CTL-"))
					CTLAssociation = ValueAtZero.ToString.Substring(ValueAtZero.ToString.IndexOf(DesignFSA))
				Case ValueAtZero IsNot DBNull.Value AndAlso (ValueAtZero.ToString.Contains("-DJL-"))
					DJLName = ValueAtZero.ToString.Substring(ValueAtZero.ToString.Length - 18)
					If ValueAtOne IsNot DBNull.Value Then
						DJLAddress = ValueAtOne
					Else
						DJLAddress = String.Empty
					End If
				Case ValueAtZero IsNot DBNull.Value AndAlso (ValueAtZero.ToString.Contains("-MPT-"))
					MPTAssociation = ValueAtZero.ToString.Substring(ValueAtZero.length - 18) '
					MPT_SearchValue = ValueAtZero.ToString
					'ADDED Or (ValueAtOne.ToString.Contains("-FSD-"))) FOR TESTING PURPOSE ONLY
				Case ValueAtOne IsNot DBNull.Value AndAlso ValueAtOne.ToString.Contains("-DSS-") 'Or (ValueAtOne.ToString.Contains("-FSD-")))
					If MPT_SearchValue = String.Empty Then
						DSS_SearchValue = ValueAtOne.ToString.Substring(ValueAtOne.length - 18)
						DSS_SequenceValue = Row.ItemArray(2).ToString
					End If
					If FANAssociation <> String.Empty AndAlso PON_Patch_DSS_SearchValue = String.Empty AndAlso PON_Patch_DSS_SequenceValue = String.Empty Then
						PON_Patch_DSS_SearchValue = ValueAtOne.ToString.Substring(ValueAtOne.length - 18)
						PON_Patch_DSS_SequenceValue = Row.ItemArray(2).ToString
					End If
					DSSName = ValueAtOne.ToString.Substring(ValueAtOne.ToString.Length - 18)
					DSSFibre = Regex.Replace(ValueAtOne.ToString.Substring(0, ValueAtOne.ToString.Length - 20), "[A-Za-z_:]", "")
					If DSSFibre.Length >= 4 Then
						DSSFibre = DSSFibre.Substring(0, 3) + "F"
					Else
						DSSFibre = DSSFibre + "F"
					End If
				Case ValueAtOne IsNot DBNull.Value AndAlso ((ValueAtOne.ToString.Contains("-FAN-")) Or (ValueAtOne.ToString.Contains("-TAN-")) Or (ValueAtOne.ToString.Contains("-AGG-")))
					FANAssociation = ValueAtOne.ToString.Substring(ValueAtOne.ToString.Length - 19)
				Case ValueAtOne IsNot DBNull.Value AndAlso (ValueAtOne.ToString.Contains("FSD"))
					FSDNAME = ValueAtOne.ToString.Substring(ValueAtOne.ToString.Length - 18)
					FSDFibre = "F" + Row.ItemArray(2).ToString
					FSDFibreCount = ValueAtOne.ToString.Substring(0, ValueAtOne.ToString.IndexOf(Chr(32)) - 1)
					FSDFibreCount = Regex.Replace(FSDFibreCount, "[A-Za-z_:]", "") + "F"
					FSDFibreLength = String.Format("{0:F1}", Double.Parse(Row.ItemArray(7).ToString)) + "m"
				Case ValueAtOne IsNot DBNull.Value AndAlso (ValueAtOne.ToString.Contains("SPL")) AndAlso Row.ItemArray(3).ToString.Contains("IN") AndAlso FJLSpliceIN = String.Empty
					If ProjectType = "EE" Then
						FJLSpliceIN = ValueAtOne.ToString.Substring(0, 4) + Chr(45) +
						  ValueAtOne.ToString.Substring(4, 2) + Chr(45) +
						  ValueAtOne.ToString.Substring(6, 3) + Chr(45) +
						  ValueAtOne.ToString.Substring(9, 3) + Chr(45) +
						  ValueAtOne.ToString.Substring(12, 3) + Chr(32)
						'Row.ItemArray(3).ToString + Chr(32)
					ElseIf ProjectType = "GPON" Then
						FJLSpliceIN = ValueAtOne.ToString.Substring(0, 4) + Chr(45) +
						  ValueAtOne.ToString.Substring(4, 2) + Chr(45) +
						  ValueAtOne.ToString.Substring(6, 2) + Chr(45) +
						  ValueAtOne.ToString.Substring(8, 3) + Chr(45) +
						  ValueAtOne.ToString.Substring(11, 3) + Chr(32) +
						  Row.ItemArray(3).ToString + Chr(32) + Row.ItemArray(2).ToString
					End If

				Case ValueAtOne IsNot DBNull.Value AndAlso (ValueAtOne.ToString.Contains("SPL")) AndAlso Row.ItemArray(3).ToString.Contains("OUT") AndAlso FJLSpliceOUT = String.Empty
					FJLSplitter = ValueAtOne
					If ProjectType = "EE" Then
						FJLSpliceOUT = ValueAtOne.ToString.Substring(0, 4) + Chr(45) +
										  ValueAtOne.ToString.Substring(4, 2) + Chr(45) +
										  ValueAtOne.ToString.Substring(6, 3) + Chr(45) +
										  ValueAtOne.ToString.Substring(9, 3) + Chr(45) +
										  ValueAtOne.ToString.Substring(12, 4) + Chr(32)
						'Row.ItemArray(3).ToString + Chr(32)
					ElseIf ProjectType = "GPON" Then
						FJLSpliceOUT = ValueAtOne.ToString.Substring(0, 4) + Chr(45) +
						  ValueAtOne.ToString.Substring(4, 2) + Chr(45) +
						  ValueAtOne.ToString.Substring(6, 2) + Chr(45) +
						  ValueAtOne.ToString.Substring(8, 3) + Chr(45) +
						  ValueAtOne.ToString.Substring(11, 3) + Chr(32) +
						Row.ItemArray(3).ToString + Chr(32)
					End If

					' FSLNAME = String.Empty in the Line Below ADDED ON 20191024 as suggested by Amir to be specific for JOB 3MDC-22-AYCA-5LMGCA-DESIGN-V1.0
					' Any further Impact of this change needs to be monitored
				Case ValueAtOne IsNot DBNull.Value AndAlso (ValueAtOne.ToString.Contains("FSL")) AndAlso FSLNAME = String.Empty
					FSLNAME = ValueAtOne.ToString.Substring(ValueAtOne.ToString.IndexOf(Chr(32))) + Chr(32) + "F" + Row.ItemArray(2).ToString

				Case ValueAtZero.ToString.Contains("OSR") AndAlso ValueAtSequence = String.Empty
					ValueAtSequence = Row.ItemArray(2).ToString
				Case ValueAtSequence <> String.Empty AndAlso PortAtSequence = String.Empty AndAlso IsNumeric(Row.ItemArray(2).ToString)
					PortAtSequence = Row.ItemArray(2).ToString
			End Select
			If DJLName <> String.Empty AndAlso DSSName <> String.Empty Then
				dicDJL_Cable_IN.Add(DJLName, New List(Of String)({DesignProjectID, ApplicationName.ToString, "LFN", DJLAddress, DSSName, DSSFibre}))
				'dicDJL_DSS.Add(DJLName, DSSName)
				DJLName = String.Empty
				DSSName = String.Empty
			End If
			If MPTAssociation <> String.Empty AndAlso DSSName <> String.Empty Then
				MPTAssocTempList.Add(DSSName)
				If MPT_SearchValue.Contains("4 Port") Then
					MPTAssocTempList.Add("B4")
				ElseIf MPT_SearchValue.Contains("8 Port") Then
					MPTAssocTempList.Add("B8")
				ElseIf MPT_SearchValue.Contains("12 Port") Then
					MPTAssocTempList.Add("B12")
				End If
				MPTAssocTempList.Add(DesignProjectID)
				MPTAssocTempList.Add(ApplicationName.ToString)
				MPTAssocTempList.Add("LFN")
				MPTAssocTempList.Add(ValueAtOne)
				MPTAssocTempList.Add(DSSFibre)
				DSSName = String.Empty
			End If
		Next
		Dim STATE_SUBURB As String = DesignSuburb.ToString + Chr(44) + Chr(32) + DesignState.ToString + Chr(45) + DesignPostCode.ToString
		dicPilotJoint.Add(PilotJoint, New List(Of String)({DesignProjectID, ApplicationName.ToString, "LFN", PilotAddress, STATE_SUBURB, FSDNAME, FSDFibreCount, FSDFibreLength}))
		If FJLSpliceIN <> String.Empty AndAlso FJLSpliceOUT <> String.Empty Then
			dicPilotJointCable_Table.Add(PilotJoint, New List(Of String)({(FSDNAME + Chr(32) + FSDFibre) + Chr(32) + Chr(61) + Chr(32) + FJLSpliceIN, FJLSpliceOUT + Chr(32) + Chr(61) + Chr(32) + FSLNAME}))
		ElseIf FJLSpliceIN = String.Empty AndAlso FJLSpliceOUT = String.Empty Then
			dicPilotJointCable_Table.Add(PilotJoint, New List(Of String)({(FSDNAME + Chr(32) + FSDFibre) + Chr(32) + Chr(61) + Chr(32) + FSLNAME}))
		End If

		Dim TempKeyValue As String = String.Empty
		Dim tempList As New List(Of String)
		Select Case True
			Case Integer.Parse(DSS_SequenceValue) <= 4
				TempKeyValue = DSS_SearchValue + Chr(32) + "F" + DSS_SequenceValue + Chr(32) + "=" + Chr(32) + FSDNAME + Chr(32) + "F" + DSS_SequenceValue
				tempList.Add(TempKeyValue)
				dicMPTCable_Table.Add(DSS_SearchValue, tempList)
			Case Integer.Parse(DSS_SequenceValue) >= 5 AndAlso Integer.Parse(DSS_SequenceValue) <= 8
				TempKeyValue = DSS_SearchValue + Chr(32) + "F5-12" + Chr(32) + "=" + Chr(32) + FSDNAME + Chr(32) + "F1-4"
				tempList.Add(TempKeyValue)
				dicMPTCable_Table.Add(DSS_SearchValue, tempList)
			Case Integer.Parse(DSS_SequenceValue) > 8
				TempKeyValue = DSS_SearchValue + Chr(32) + "F9-12" + Chr(32) + "=" + Chr(32) + FSDNAME + Chr(32) + "F1-4"
				tempList.Add(TempKeyValue)
				dicMPTCable_Table.Add(DSS_SearchValue, tempList)
		End Select
	End Sub
    Private Shared Sub Start_LFN_SLD_AFTER_CTL_IN_CAD()
        Dim AttTAG As New List(Of String)
        Dim AttVal As New List(Of String)
        Dim blkType As String = String.Empty

        If LFN_ICD_Name <> String.Empty Then
            dctLFN_SLD_ICD_AttTAG_AttVal.Clear()
            dctLFN_SLD_ICD_AttTAG_AttVal = PopulateBlockAttributes(blkLFNSLDICD)
            SetBlockAttributeValues(dctLFN_SLD_ICD_AttTAG_AttVal)
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_ICD_NAME1") = LFN_ICD_Name
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_ICD_NAME2") = LFN_ICD_Name
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_ICD_NAME3") = LFN_ICD_Name
            dctLFN_SLD_ICD_AttTAG_AttVal("ADDRESS1") = LFN_ICD_ADDRESS
            dctLFN_SLD_ICD_AttTAG_AttVal("SUBURB_STATE1") = LFN_ICD_SUBURB_STATE
            dctLFN_SLD_ICD_AttTAG_AttVal("ADDRESS2") = LFN_ICD_ADDRESS
            dctLFN_SLD_ICD_AttTAG_AttVal("SUBURB_STATE2") = LFN_ICD_SUBURB_STATE
            dctLFN_SLD_ICD_AttTAG_AttVal("ADDRESS3") = LFN_ICD_ADDRESS
            dctLFN_SLD_ICD_AttTAG_AttVal("SUBURB_STATE3") = LFN_ICD_SUBURB_STATE
			dctLFN_SLD_ICD_AttTAG_AttVal("PROJECT_ID") = DesignProjectID
			dctLFN_SLD_ICD_AttTAG_AttVal("PROJECT_NAME") = ApplicationName.ToString
			dctLFN_SLD_ICD_AttTAG_AttVal("NETWORK_TYPE") = "LFN"
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_PIC_NAME1") = LFN_PIC_NAME
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_PIC_NAME2") = LFN_PIC_NAME
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_PIC_NAME3") = LFN_PIC_NAME
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_PIC_FIBRE1") = LFN_PIC_Fibre
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_PIC_FIBRE2") = LFN_PIC_Fibre
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_PIC_FIBRE3") = LFN_PIC_Fibre
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_PIC_LENGTH1") = LFN_PIC_LENGTH
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_PIC_LENGTH2") = LFN_PIC_LENGTH
            dctLFN_SLD_ICD_AttTAG_AttVal("SLD_PIC_LENGTH3") = LFN_PIC_LENGTH
            AttTAG.Clear()
            AttVal.Clear()
            AttTAG = dctLFN_SLD_ICD_AttTAG_AttVal.Keys.ToList()
            AttVal = dctLFN_SLD_ICD_AttTAG_AttVal.Values.ToList()
            Dim blkLFNICDInsPt As Point3d = New Point3d(blkLFCTLInsPt.X + 71.25, blkLFCTLInsPt.Y, blkLFCTLInsPt.Z)
            Dim blkLFNICDRotation As New Double
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            InsertBlockWithAttributes(ProjectType, blkLFNSLDICD, "0", blkType, blkLFNICDInsPt, blkLFNICDRotation, blkLFNICDInsPt, AttTAG, AttVal)
            SETCADSystemVariable("CLAYER", "PROPOSED")
            SETCADSystemVariable("CELTYPE", "continuous")
            Dim ICD_SPLICE_TABLE_InsPT As Point3d = New Point3d(blkLFNICDInsPt.X + 30.25, blkLFNICDInsPt.Y, blkLFNICDInsPt.Z)
            dicLFN_ICD_Splice = New Dictionary(Of String, List(Of String))
            Dim TempList As New List(Of String)
            TempList.Add(LFN_PIC_NAME + Chr(32) + Chr(70) + LFN_PIC_SEQUENCE + Chr(32) + Chr(61) + Chr(32) + LFN_PDC_NAME + Chr(32) + Chr(70) + LFN_PDC_SEQUENCE)
            dicLFN_ICD_Splice.Add(LFN_ICD_Name, TempList)
            Start_SLD_TABLE(dicLFN_ICD_Splice, LFN_ICD_Name, ICD_SPLICE_TABLE_InsPT)
            SETCADSystemVariable("CLAYER", layOtherProjectLayer)
            dctLFN_SLD_NTD_AttTAG_AttVal.Clear()

            dctLFN_SLD_NTD_AttTAG_AttVal = PopulateBlockAttributes(blkLFNSLDNTD)
            SetBlockAttributeValues(dctLFN_SLD_NTD_AttTAG_AttVal)
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_NTD_NAME1") = LFN_NTD_Name
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_NTD_NAME2") = LFN_NTD_Name
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_NTD_NAME3") = LFN_NTD_Name
            dctLFN_SLD_NTD_AttTAG_AttVal("ADDRESS1") = LFN_NTD_ADDRESS
            dctLFN_SLD_NTD_AttTAG_AttVal("SUBURB_STATE1") = LFN_NTD_SUBURB_STATE
            dctLFN_SLD_NTD_AttTAG_AttVal("ADDRESS2") = LFN_NTD_ADDRESS
            dctLFN_SLD_NTD_AttTAG_AttVal("SUBURB_STATE2") = LFN_NTD_SUBURB_STATE
            dctLFN_SLD_NTD_AttTAG_AttVal("ADDRESS3") = LFN_NTD_ADDRESS
            dctLFN_SLD_NTD_AttTAG_AttVal("SUBURB_STATE3") = LFN_NTD_SUBURB_STATE
			dctLFN_SLD_NTD_AttTAG_AttVal("PROJECT_ID") = DesignProjectID
			dctLFN_SLD_NTD_AttTAG_AttVal("PROJECT_NAME") = ApplicationName.ToString
			dctLFN_SLD_NTD_AttTAG_AttVal("NETWORK_TYPE") = "LFN"
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_NAME1") = LFN_PDC_NAME
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_NAME2") = LFN_PDC_NAME
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_NAME3") = LFN_PDC_NAME
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_FIBRE1") = LFN_PDC_Fibre
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_FIBRE2") = LFN_PDC_Fibre
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_FIBRE3") = LFN_PDC_Fibre
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_LENGTH1") = LFN_PDC_LENGTH
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_LENGTH2") = LFN_PDC_LENGTH
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_LENGTH3") = LFN_PDC_LENGTH
            AttTAG.Clear()
            AttVal.Clear()
            AttTAG = dctLFN_SLD_NTD_AttTAG_AttVal.Keys.ToList()
            AttVal = dctLFN_SLD_NTD_AttTAG_AttVal.Values.ToList()
            Dim blkLFNNTDInsPt As Point3d = New Point3d(blkLFNICDInsPt.X + 71.25, blkLFNICDInsPt.Y, blkLFNICDInsPt.Z)
            Dim blkLFNNTDRotation As New Double
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            InsertBlockWithAttributes(ProjectType, blkLFNSLDNTD, "0", blkType, blkLFNNTDInsPt, blkLFNNTDRotation, blkLFNNTDInsPt, AttTAG, AttVal)
        ElseIf LFN_ICD_Name = String.Empty Then
            dctLFN_SLD_NTD_AttTAG_AttVal = PopulateBlockAttributes(blkLFNSLDNTD)
            SetBlockAttributeValues(dctLFN_SLD_NTD_AttTAG_AttVal)
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_NTD_NAME1") = LFN_NTD_Name
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_NTD_NAME2") = LFN_NTD_Name
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_NTD_NAME3") = LFN_NTD_Name
            dctLFN_SLD_NTD_AttTAG_AttVal("ADDRESS1") = LFN_NTD_ADDRESS
            dctLFN_SLD_NTD_AttTAG_AttVal("SUBURB_STATE1") = LFN_NTD_SUBURB_STATE
            dctLFN_SLD_NTD_AttTAG_AttVal("ADDRESS2") = LFN_NTD_ADDRESS
            dctLFN_SLD_NTD_AttTAG_AttVal("SUBURB_STATE2") = LFN_NTD_SUBURB_STATE
            dctLFN_SLD_NTD_AttTAG_AttVal("ADDRESS3") = LFN_NTD_ADDRESS
            dctLFN_SLD_NTD_AttTAG_AttVal("SUBURB_STATE3") = LFN_NTD_SUBURB_STATE
			dctLFN_SLD_NTD_AttTAG_AttVal("PROJECT_ID") = DesignProjectID
			dctLFN_SLD_NTD_AttTAG_AttVal("PROJECT_NAME") = ApplicationName.ToString
			dctLFN_SLD_NTD_AttTAG_AttVal("NETWORK_TYPE") = "LFN"
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_NAME1") = LFN_PIC_NAME
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_NAME2") = LFN_PIC_NAME
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_NAME3") = LFN_PIC_NAME
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_FIBRE1") = LFN_PIC_Fibre
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_FIBRE2") = LFN_PIC_Fibre
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_FIBRE3") = LFN_PIC_Fibre
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_LENGTH1") = LFN_PIC_LENGTH
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_LENGTH2") = LFN_PIC_LENGTH
            dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_LENGTH3") = LFN_PIC_LENGTH
            AttTAG.Clear()
            AttVal.Clear()
            AttTAG = dctLFN_SLD_NTD_AttTAG_AttVal.Keys.ToList()
            AttVal = dctLFN_SLD_NTD_AttTAG_AttVal.Values.ToList()
            Dim blkLFNNTDInsPt As Point3d = New Point3d(blkLFCTLInsPt.X + 71.25, blkLFCTLInsPt.Y, blkLFCTLInsPt.Z)
            Dim blkLFNNTDRotation As New Double
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            InsertBlockWithAttributes(ProjectType, blkLFNSLDNTD, "0", blkType, blkLFNNTDInsPt, blkLFNNTDRotation, blkLFNNTDInsPt, AttTAG, AttVal)
        End If
    End Sub
    Private Shared Sub Start_LFN_SLD_IN_CAD()
		Dim CurOSMODE = GETCADSystemVariable("OSMODE")
		Dim CurLayer As String = GETCADSystemVariable("CLAYER")
		SETCADSystemVariable("OSMODE", 0)
		SETCADSystemVariable("CLAYER", layOtherProjectLayer)
		Dim blkType As String = String.Empty
		Dim AttTAG As New List(Of String)
		Dim AttVal As New List(Of String)

		dctSLDFJL_AttTAG_AttVal.Clear()
		dctSLDFJL_AttTAG_AttVal = PopulateBlockAttributes(blkLFNSLDFJL)
		SetBlockAttributeValues(dctSLDFJL_AttTAG_AttVal)
		dctSLDFJL_AttTAG_AttVal("SLD_FJL_NAME1") = LFN_FJL_Name
		dctSLDFJL_AttTAG_AttVal("SLD_FJL_NAME2") = LFN_FJL_Name
		dctSLDFJL_AttTAG_AttVal("SLD_FJL_NAME3") = LFN_FJL_Name
		dctSLDFJL_AttTAG_AttVal("PROJECT_ID") = DesignProjectID
		dctSLDFJL_AttTAG_AttVal("PROJECT_NAME") = ApplicationName.ToString
		dctSLDFJL_AttTAG_AttVal("NETWORK_TYPE") = "LFN"
		dctSLDFJL_AttTAG_AttVal("ADDRESS1") = LFN_FJL_ADDRESS
		dctSLDFJL_AttTAG_AttVal("SUBURB_STATE1") = LFN_FJL_SUBURB_STATE
		dctSLDFJL_AttTAG_AttVal("ADDRESS2") = LFN_FJL_ADDRESS
		dctSLDFJL_AttTAG_AttVal("SUBURB_STATE2") = LFN_FJL_SUBURB_STATE
		dctSLDFJL_AttTAG_AttVal("ADDRESS3") = LFN_FJL_ADDRESS
		dctSLDFJL_AttTAG_AttVal("SUBURB_STATE3") = LFN_FJL_SUBURB_STATE

		AttTAG = dctSLDFJL_AttTAG_AttVal.Keys.ToList()
		AttVal = dctSLDFJL_AttTAG_AttVal.Values.ToList()

		Dim Prompt1 As String = "Select Insertion Point For SLD: "
		Dim Prompt2 As String = String.Empty
		Dim blkLFNFJLInsPt As New Point3d
		Dim blkLFNFJLRotation As New Double

		Dim blkLFNBJLInsPt As New Point3d
		Dim blkLFNBJLRotation As New Double

		Dim blkLFNSMPInsPt As New Point3d
		Dim blkLFNSMPRotation As New Double

		Dim blkLFNPCDInsPt As New Point3d
		Dim blkLFNPCDRotation As New Double

		Dim blkLFNNTDInsPt As New Point3d
		Dim blkLFNNTDRotation As New Double


		Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
		InsertBDOD_Block(blkLFNSLDFJL, AttVal, Prompt1, Prompt2, blkLFNFJLInsPt, blkLFNFJLRotation)
		If dctLFN_SLD_CTL_AttTAG_AttVal.Count > 0 Then

			dctLFN_SLD_CTL_AttTAG_AttVal.Clear()
			dctLFN_SLD_CTL_AttTAG_AttVal = PopulateBlockAttributes(blkLFNSLDCTL)
			SetBlockAttributeValues(dctLFN_SLD_CTL_AttTAG_AttVal)
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_CTL_NAME1") = LFN_CTL_NAME
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_CTL_NAME2") = LFN_CTL_NAME
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_CTL_NAME3") = LFN_CTL_NAME
			dctLFN_SLD_CTL_AttTAG_AttVal("ADDRESS1") = LFN_CTL_ADDRESS
			dctLFN_SLD_CTL_AttTAG_AttVal("SUBURB_STATE1") = LFN_CTL_SUBURB_STATE
			dctLFN_SLD_CTL_AttTAG_AttVal("ADDRESS2") = LFN_CTL_ADDRESS
			dctLFN_SLD_CTL_AttTAG_AttVal("SUBURB_STATE2") = LFN_CTL_SUBURB_STATE
			dctLFN_SLD_CTL_AttTAG_AttVal("ADDRESS3") = LFN_CTL_ADDRESS
			dctLFN_SLD_CTL_AttTAG_AttVal("SUBURB_STATE3") = LFN_CTL_SUBURB_STATE
			dctLFN_SLD_CTL_AttTAG_AttVal("PROJECT_ID") = DesignProjectID
			dctLFN_SLD_CTL_AttTAG_AttVal("PROJECT_NAME") = ApplicationName.ToString
			dctLFN_SLD_CTL_AttTAG_AttVal("NETWORK_TYPE") = "LFN"
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_FSL_NAME1") = LFN_FSL_NAME
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_FSL_NAME2") = LFN_FSL_NAME
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_FSL_NAME3") = LFN_FSL_NAME
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_FSL_FIBRE1") = LFN_FSL_Fibre
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_FSL_FIBRE2") = LFN_FSL_Fibre
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_FSL_FIBRE3") = LFN_FSL_Fibre
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_FSL_LENGTH1") = LFN_FSL_LENGTH
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_FSL_LENGTH2") = LFN_FSL_LENGTH
			dctLFN_SLD_CTL_AttTAG_AttVal("SLD_FSL_LENGTH3") = LFN_FSL_LENGTH

			AttTAG.Clear()
			AttVal.Clear()
			AttTAG = dctLFN_SLD_CTL_AttTAG_AttVal.Keys.ToList()
			AttVal = dctLFN_SLD_CTL_AttTAG_AttVal.Values.ToList()
			blkLFCTLInsPt = New Point3d(blkLFNFJLInsPt.X + 2.95, blkLFNFJLInsPt.Y, blkLFNFJLInsPt.Z)
			Dim blkLFNCTLRotation As New Double
			Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
			InsertBlockWithAttributes(ProjectType, blkLFNSLDCTL, "0", blkType, blkLFCTLInsPt, blkLFNCTLRotation, blkLFCTLInsPt, AttTAG, AttVal)

			Dim CTL_SPLICE_TABLE_InsPT As Point3d = New Point3d(blkLFCTLInsPt.X + 30.25, blkLFCTLInsPt.Y, blkLFCTLInsPt.Z)
			dicLFN_CTL_Splice = New Dictionary(Of String, List(Of String))
			dicLFN_CTL_Splice.Add(LFN_CTL_NAME, dicLFN_Splice)
			Start_SLD_TABLE(dicLFN_CTL_Splice, LFN_CTL_NAME, CTL_SPLICE_TABLE_InsPT)

		ElseIf LFN_BJL_Name.Count > 0 Then
			Dim LFN_BJL_Count As Integer = LFN_BJL_Name.Count
			Dim LFN_BJL_Index As Integer = 0
			blkLFNBJLInsPt = New Point3d(blkLFNFJLInsPt.X + 2.95, blkLFNFJLInsPt.Y, blkLFNFJLInsPt.Z)

			Do While LFN_BJL_Count > 0

				dctLFN_SLD_BJL_AttTAG_AttVal.Clear()
				dctLFN_SLD_BJL_AttTAG_AttVal = PopulateBlockAttributes(blkLFNSLDBJL)
				SetBlockAttributeValues(dctLFN_SLD_BJL_AttTAG_AttVal)
				Dim LFN_BJL_Values As New List(Of String)
				LFN_BJL_Name.TryGetValue(LFN_BJL_Name.Keys(LFN_BJL_Index), LFN_BJL_Values)
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_BJL_NAME1") = LFN_BJL_Name.Keys(LFN_BJL_Index).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_BJL_NAME2") = LFN_BJL_Name.Keys(LFN_BJL_Index).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_BJL_NAME3") = LFN_BJL_Name.Keys(LFN_BJL_Index).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("ADDRESS1") = LFN_BJL_Values(0).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SUBURB_STATE1") = LFN_BJL_Values(1).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("ADDRESS2") = LFN_BJL_Values(0).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SUBURB_STATE2") = LFN_BJL_Values(1).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("ADDRESS3") = LFN_BJL_Values(0).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SUBURB_STATE3") = LFN_BJL_Values(1).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("PROJECT_ID") = DesignProjectID
				dctLFN_SLD_BJL_AttTAG_AttVal("PROJECT_NAME") = ApplicationName.ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("NETWORK_TYPE") = "LFN"
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_FSL_NAME1") = LFN_BJL_Values(2).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_FSL_NAME2") = LFN_BJL_Values(2).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_FSL_NAME3") = LFN_BJL_Values(2).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_FSL_FIBRE1") = LFN_BJL_Values(3).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_FSL_FIBRE2") = LFN_BJL_Values(3).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_FSL_FIBRE3") = LFN_BJL_Values(3).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_FSL_LENGTH1") = LFN_BJL_Values(5).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_FSL_LENGTH2") = LFN_BJL_Values(5).ToString
				dctLFN_SLD_BJL_AttTAG_AttVal("SLD_FSL_LENGTH3") = LFN_BJL_Values(5).ToString

				AttTAG.Clear()
				AttVal.Clear()
				AttTAG = dctLFN_SLD_BJL_AttTAG_AttVal.Keys.ToList()
				AttVal = dctLFN_SLD_BJL_AttTAG_AttVal.Values.ToList()
				Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
				InsertBlockWithAttributes(ProjectType, blkLFNSLDBJL, "0", blkType, blkLFNBJLInsPt, blkLFNBJLRotation, blkLFNBJLInsPt, AttTAG, AttVal)

				Dim BJL_SPLICE_TABLE_InsPT As Point3d = New Point3d(blkLFNBJLInsPt.X + 30.25, blkLFNBJLInsPt.Y, blkLFNBJLInsPt.Z)
				'dicLFN_BJL_Splice = New Dictionary(Of String, List(Of String))
				'dicLFN_BJL_Splice.Add(LFN_BJL_Name.Keys(LFN_BJL_Index).ToString, dicLFN_Splice)
				Dim CurBJL As String = LFN_BJL_Name.Keys(LFN_BJL_Index)
				'dicLFN_BJL_Splice.Keys(Index)
				Start_SLD_TABLE(dicLFN_BJL_Splice, CurBJL, BJL_SPLICE_TABLE_InsPT)

				LFN_BJL_Count -= 1
				LFN_BJL_Index += 1
				If LFN_BJL_Count > 0 Then
					blkLFNBJLInsPt = New Point3d(blkLFNBJLInsPt.X + 71.25, blkLFNBJLInsPt.Y, blkLFNBJLInsPt.Z)
				End If
			Loop

			If LFN_SMP_Name <> String.Empty Then
				dctLFN_SLD_SMP_AttTAG_AttVal.Clear()
				dctLFN_SLD_SMP_AttTAG_AttVal = PopulateBlockAttributes(blkLFNSLDSMP)
				SetBlockAttributeValues(dctLFN_SLD_SMP_AttTAG_AttVal)
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SMP_NAME1") = LFN_SMP_Name
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SMP_NAME2") = LFN_SMP_Name
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SMP_NAME3") = LFN_SMP_Name
				dctLFN_SLD_SMP_AttTAG_AttVal("TYPE1") = LFN_SMP_TYPE
				dctLFN_SLD_SMP_AttTAG_AttVal("TYPE2") = LFN_SMP_TYPE
				dctLFN_SLD_SMP_AttTAG_AttVal("TYPE3") = LFN_SMP_TYPE
				dctLFN_SLD_SMP_AttTAG_AttVal("RATIO") = LFN_SMP_RATIO

				dctLFN_SLD_SMP_AttTAG_AttVal("ADDRESS1") = LFN_SMP_ADDRESS
				dctLFN_SLD_SMP_AttTAG_AttVal("ADDRESS2") = LFN_SMP_ADDRESS
				dctLFN_SLD_SMP_AttTAG_AttVal("ADDRESS3") = LFN_SMP_ADDRESS
				dctLFN_SLD_SMP_AttTAG_AttVal("SUBURB_STATE1") = LFN_SMP_SUBURB_STATE
				dctLFN_SLD_SMP_AttTAG_AttVal("SUBURB_STATE2") = LFN_SMP_SUBURB_STATE
				dctLFN_SLD_SMP_AttTAG_AttVal("SUBURB_STATE3") = LFN_SMP_SUBURB_STATE
				dctLFN_SLD_SMP_AttTAG_AttVal("PROJECT_ID") = DesignProjectID
				dctLFN_SLD_SMP_AttTAG_AttVal("PROJECT_NAME") = ApplicationName.ToString
				dctLFN_SLD_SMP_AttTAG_AttVal("NETWORK_TYPE") = "LFN"
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SSS_NAME1") = LFN_SSS_NAME
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SSS_NAME2") = LFN_SSS_NAME
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SSS_NAME3") = LFN_SSS_NAME
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SSS_FIBRE1") = LFN_SSS_Fibre
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SSS_FIBRE2") = LFN_SSS_Fibre
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SSS_FIBRE3") = LFN_SSS_Fibre
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SSS_LENGTH1") = LFN_SSS_LENGTH
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SSS_LENGTH2") = LFN_SSS_LENGTH
				dctLFN_SLD_SMP_AttTAG_AttVal("SLD_SSS_LENGTH3") = LFN_SSS_LENGTH
				AttTAG.Clear()
				AttVal.Clear()
				AttTAG = dctLFN_SLD_SMP_AttTAG_AttVal.Keys.ToList()
				AttVal = dctLFN_SLD_SMP_AttTAG_AttVal.Values.ToList()

				blkLFNSMPInsPt = New Point3d(blkLFNBJLInsPt.X + 71.25, blkLFNBJLInsPt.Y, blkLFNBJLInsPt.Z)
				Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
				InsertBlockWithAttributes(ProjectType, blkLFNSLDSMP, "0", blkType, blkLFNSMPInsPt, blkLFNSMPRotation, blkLFNSMPInsPt, AttTAG, AttVal)

				Dim SMP_SPLICE_TABLE_InsPT As Point3d = New Point3d(blkLFNSMPInsPt.X + 30.25, blkLFNSMPInsPt.Y, blkLFNSMPInsPt.Z)
				dicLFN_Splice = New List(Of String)
				dicLFN_Splice.Add(LFN_SSS_NAME + Chr(32) + LFN_SSS_Fibre_Sequence + Chr(32) + Chr(61) + Chr(32) + LFN_SMP_SPLITTER_IN)
				dicLFN_Splice.Add(LFN_SMP_SPLITTER_OUT + Chr(32) + Chr(61) + Chr(32) + LFN_SDS_NAME + Chr(32) + LFN_SDS_Fibre_Sequence)
				dicLFN_SMP_Splice = New Dictionary(Of String, List(Of String))
				dicLFN_SMP_Splice.Add(LFN_SMP_Name, dicLFN_Splice)
				Start_SLD_TABLE(dicLFN_SMP_Splice, LFN_SMP_Name, SMP_SPLICE_TABLE_InsPT)
				If LFN_PCD_Name <> String.Empty Then
					dctLFN_SLD_PCD_AttTAG_AttVal.Clear()
					dctLFN_SLD_PCD_AttTAG_AttVal = PopulateBlockAttributes(blkLFNSLDPCD)
					SetBlockAttributeValues(dctLFN_SLD_PCD_AttTAG_AttVal)

					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_PCD_NAME1") = LFN_PCD_Name
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_PCD_NAME2") = LFN_PCD_Name
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_PCD_NAME3") = LFN_PCD_Name

					dctLFN_SLD_PCD_AttTAG_AttVal("ADDRESS1") = LFN_PCD_ADDRESS
					dctLFN_SLD_PCD_AttTAG_AttVal("ADDRESS2") = LFN_PCD_ADDRESS
					dctLFN_SLD_PCD_AttTAG_AttVal("ADDRESS3") = LFN_PCD_ADDRESS
					dctLFN_SLD_PCD_AttTAG_AttVal("SUBURB_STATE1") = LFN_PCD_SUBURB_STATE
					dctLFN_SLD_PCD_AttTAG_AttVal("SUBURB_STATE2") = LFN_PCD_SUBURB_STATE
					dctLFN_SLD_PCD_AttTAG_AttVal("SUBURB_STATE3") = LFN_PCD_SUBURB_STATE
					dctLFN_SLD_PCD_AttTAG_AttVal("PROJECT_ID") = DesignProjectID
					dctLFN_SLD_PCD_AttTAG_AttVal("PROJECT_NAME") = ApplicationName.ToString
					dctLFN_SLD_PCD_AttTAG_AttVal("NETWORK_TYPE") = "LFN"
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_SDS_NAME1") = LFN_SDS_NAME
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_SDS_NAME2") = LFN_SDS_NAME
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_SDS_NAME3") = LFN_SDS_NAME
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_SDS_FIBRE1") = LFN_SDS_Fibre
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_SDS_FIBRE2") = LFN_SDS_Fibre
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_SDS_FIBRE3") = LFN_SDS_Fibre
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_SDS_LENGTH1") = LFN_SDS_LENGTH
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_SDS_LENGTH2") = LFN_SDS_LENGTH
					dctLFN_SLD_PCD_AttTAG_AttVal("SLD_SDS_LENGTH3") = LFN_SDS_LENGTH

					AttTAG.Clear()
					AttVal.Clear()
					AttTAG = dctLFN_SLD_PCD_AttTAG_AttVal.Keys.ToList()
					AttVal = dctLFN_SLD_PCD_AttTAG_AttVal.Values.ToList()

					blkLFNPCDInsPt = New Point3d(blkLFNSMPInsPt.X + 71.25, blkLFNSMPInsPt.Y, blkLFNSMPInsPt.Z)
					Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
					InsertBlockWithAttributes(ProjectType, blkLFNSLDPCD, "0", blkType, blkLFNPCDInsPt, blkLFNPCDRotation, blkLFNPCDInsPt, AttTAG, AttVal)

					Dim PCD_SPLICE_TABLE_InsPT As Point3d = New Point3d(blkLFNPCDInsPt.X + 30.25, blkLFNPCDInsPt.Y, blkLFNPCDInsPt.Z)
					dicLFN_Splice = New List(Of String)
					dicLFN_Splice.Add(LFN_SDS_NAME + Chr(32) + LFN_SDS_Fibre_Sequence + Chr(32) + Chr(61) + Chr(32) + LFN_PIC_NAME + Chr(32) + "F" + LFN_PIC_SEQUENCE)
					dicLFN_PCD_Splice = New Dictionary(Of String, List(Of String))
					dicLFN_PCD_Splice.Add(LFN_PCD_Name, dicLFN_Splice)
					Start_SLD_TABLE(dicLFN_PCD_Splice, LFN_PCD_Name, PCD_SPLICE_TABLE_InsPT)
					'*************************************************************************************************
					If LFN_NTD_Name <> String.Empty Then
						LFN_NTD_ADDRESS = LFN_PCD_ADDRESS
						LFN_NTD_SUBURB_STATE = LFN_PCD_SUBURB_STATE
						dctLFN_SLD_NTD_AttTAG_AttVal.Clear()
						dctLFN_SLD_NTD_AttTAG_AttVal = PopulateBlockAttributes(blkLFNSLDNTD)
						SetBlockAttributeValues(dctLFN_SLD_NTD_AttTAG_AttVal)

						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_NTD_NAME1") = LFN_NTD_Name
						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_NTD_NAME2") = LFN_NTD_Name
						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_NTD_NAME3") = LFN_NTD_Name

						dctLFN_SLD_NTD_AttTAG_AttVal("ADDRESS1") = LFN_NTD_ADDRESS
						dctLFN_SLD_NTD_AttTAG_AttVal("ADDRESS2") = LFN_NTD_ADDRESS
						dctLFN_SLD_NTD_AttTAG_AttVal("ADDRESS3") = LFN_NTD_ADDRESS

						dctLFN_SLD_NTD_AttTAG_AttVal("SUBURB_STATE1") = LFN_NTD_SUBURB_STATE
						dctLFN_SLD_NTD_AttTAG_AttVal("SUBURB_STATE2") = LFN_NTD_SUBURB_STATE
						dctLFN_SLD_NTD_AttTAG_AttVal("SUBURB_STATE3") = LFN_NTD_SUBURB_STATE

						dctLFN_SLD_NTD_AttTAG_AttVal("PROJECT_ID") = DesignProjectID
						dctLFN_SLD_NTD_AttTAG_AttVal("PROJECT_NAME") = ApplicationName.ToString
						dctLFN_SLD_NTD_AttTAG_AttVal("NETWORK_TYPE") = "LFN"

						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_NAME1") = LFN_PIC_NAME
						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_NAME2") = LFN_PIC_NAME
						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_NAME3") = LFN_PIC_NAME
						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_FIBRE1") = LFN_PIC_Fibre
						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_FIBRE2") = LFN_PIC_Fibre
						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_FIBRE3") = LFN_PIC_Fibre
						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_LENGTH1") = LFN_PIC_LENGTH
						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_LENGTH2") = LFN_PIC_LENGTH
						dctLFN_SLD_NTD_AttTAG_AttVal("SLD_PDC_LENGTH3") = LFN_PIC_LENGTH
						AttTAG.Clear()
						AttVal.Clear()
						AttTAG = dctLFN_SLD_NTD_AttTAG_AttVal.Keys.ToList()
						AttVal = dctLFN_SLD_NTD_AttTAG_AttVal.Values.ToList()

						blkLFNNTDInsPt = New Point3d(blkLFNPCDInsPt.X + 71.25, blkLFNPCDInsPt.Y, blkLFNPCDInsPt.Z)
						Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
						InsertBlockWithAttributes(ProjectType, blkLFNSLDNTD, "0", blkType, blkLFNNTDInsPt, blkLFNNTDRotation, blkLFNNTDInsPt, AttTAG, AttVal)
						'Redefine_Att_Tag(blkLFNSLDNTD, "PDC", "PIC")
					End If '*******************************************************************************
				End If
			End If
		End If
	End Sub
	Private Shared Sub CaptureLFNCableTable()
		LFN_CTL_AT_Cable_Side_CSV = String.Empty
		LFN_SMP_AT_Cable_Side_CSV = String.Empty
		LFN_SPLITTER_AT_Cable_Side_CSV = String.Empty
		LFN_CTL_AT_DROP_Side_CSV = String.Empty
		LFN_SMP_AT_DROP_Side_CSV = String.Empty
		dicLFN_Splice = New List(Of String)
		dicLFN_Splice.Add(LFN_FSL_NAME + Chr(32) + LFN_FSL_Fibre_Sequence + Chr(32) + Chr(61) + Chr(32) + LFN_SPLITTER_Name + Chr(32) + LFN_SPLITTER_Branch)
		Dim SplitterData As String = String.Empty
		Dim PICData As String = String.Empty
		Dim BJL_AT_CSV As String = String.Empty
		LFN_FSL_LENGTH = String.Empty
		Dim SDS_AT_CSV As String = String.Empty
		Dim blnFoundSpliceSummary As Boolean = False
		SLDdtaTable1 = New DataTable
		Dim TemADAprefix As String = LFN_NTD_Name.Substring(0, LFN_NTD_Name.IndexOf("-NTD"))
		Dim FileEntries As String() = Directory.GetFiles(FolderPath, "*.csv")
		Dim LFN_BJL_NAME_ARRAY As String() = LFN_BJL_Name.Keys.ToArray()
		dicLFN_BJL_Splice = New Dictionary(Of String, List(Of String))
		Dim BJL_Splice_List As New List(Of String)
		For Each FileName As String In FileEntries
			SLDdtaTable1 = ConvertCSVtoDataTable(FileName)
			For Each Row As DataRow In SLDdtaTable1.Rows
				If UCase(Row.ItemArray(1).ToString).Contains("CABLE SIDE") AndAlso UCase(Row.ItemArray(1).ToString).Contains("CTL") Then
					LFN_CTL_AT_Cable_Side_CSV = Row.ItemArray(1).ToString.Substring(0, Row.ItemArray(1).ToString.IndexOf(Chr(58)))
				ElseIf UCase(Row.ItemArray(2).ToString) <> "START SPLICE" AndAlso UCase(Row.ItemArray(2).ToString).Contains("SPL") AndAlso Not UCase(Row.ItemArray(2).ToString).Contains("SPLICE CASE") AndAlso Not UCase(Row.ItemArray(2).ToString).Contains("SPLITTER") Then
					LFN_SPLITTER_AT_Cable_Side_CSV = Row.ItemArray(2).ToString.Substring(0, Row.ItemArray(2).ToString.IndexOf(Chr(58)))
				ElseIf UCase(Row.ItemArray(1).ToString).Contains("DROP SIDE") AndAlso UCase(Row.ItemArray(1).ToString).Contains("CTL") Then
					LFN_CTL_AT_DROP_Side_CSV = Row.ItemArray(1).ToString.Substring(0, Row.ItemArray(1).ToString.IndexOf(Chr(58)))
					If LFN_CTL_AT_Cable_Side_CSV <> LFN_CTL_AT_DROP_Side_CSV Then
						Throw New System.Exception("MISMATCH In CSV")
					End If
				ElseIf UCase(Row.ItemArray(1).ToString).Contains("CABLE SIDE") AndAlso UCase(Row.ItemArray(1).ToString).Contains("SMP") Then
					LFN_SMP_AT_Cable_Side_CSV = Row.ItemArray(1).ToString.Substring(0, Row.ItemArray(1).ToString.IndexOf(Chr(58)))
				ElseIf UCase(Row.ItemArray(1).ToString).Contains("DROP SIDE") AndAlso UCase(Row.ItemArray(1).ToString).Contains("SMP") Then
					LFN_SMP_AT_DROP_Side_CSV = Row.ItemArray(1).ToString.Substring(0, Row.ItemArray(1).ToString.IndexOf(Chr(58)))
					If LFN_SMP_AT_Cable_Side_CSV <> LFN_SMP_AT_DROP_Side_CSV Then
						Throw New System.Exception("MISMATCH In CSV")
					End If
				ElseIf LFN_SSS_LENGTH = String.Empty AndAlso LFN_SSS_NAME <> String.Empty AndAlso UCase(Row.ItemArray(0).ToString).Contains(LFN_SSS_NAME) Then
					If Double.Parse(Row.ItemArray(6).ToString) > 0 Then
						LFN_SSS_LENGTH = Row.ItemArray(6).ToString
					ElseIf Double.Parse(Row.ItemArray(5).ToString) > 0 Then
						LFN_SSS_LENGTH = Row.ItemArray(5).ToString
					End If
					LFN_SSS_LENGTH = String.Format("{0:F1}", Double.Parse(LFN_SSS_LENGTH.ToString)) + "m"
				ElseIf LFN_CTL_AT_DROP_Side_CSV <> String.Empty AndAlso UCase(Row.ItemArray(0).ToString).StartsWith("SPLICING SUMMARY") Then
					blnFoundSpliceSummary = True
					'ElseIf LFN_CTL_AT_DROP_Side_CSV <> String.Empty AndAlso UCase(Row.ItemArray(2).ToString).StartsWith("UNCONNECTED") Then
					'	
				ElseIf LFN_CTL_AT_DROP_Side_CSV <> String.Empty AndAlso blnFoundSpliceSummary = True AndAlso UCase(Row.ItemArray(0).ToString) <> String.Empty AndAlso Not UCase(Row.ItemArray(0).ToString).StartsWith("UNCONNECTED FIBERS") Then
					SplitterData = String.Empty
					PICData = String.Empty
					SplitterData = Row.ItemArray(0).ToString
					SplitterData = SplitterData.Replace(":", String.Empty)
					SplitterData = SplitterData.Replace("Drop Side", "OUT")
					SplitterData = SplitterData.Replace("CTL", "SPL")
					SplitterData = TemADAprefix + "-" + SplitterData.Substring(SplitterData.IndexOf("SPL"), 3) + "-" + SplitterData.Substring(SplitterData.IndexOf("SPL") + 3)
					Dim PosOfColon As Integer = Row.ItemArray(2).ToString.IndexOf(Chr(58)) - (Row.ItemArray(2).ToString.IndexOf("PIC") + 3)
					PICData = TemADAprefix + "-PIC-" + Row.ItemArray(2).ToString.Substring(Row.ItemArray(2).ToString.IndexOf("PIC") + 3, PosOfColon) + Chr(32) + "F" + Row.ItemArray(2).ToString.Substring(Row.ItemArray(2).ToString.IndexOf(Chr(32)) + 1)
					dicLFN_Splice.Add(SplitterData + Chr(32) + Chr(61) + Chr(32) + PICData)
					blnFoundSpliceSummary = False
				ElseIf Row.ItemArray(1).ToString.Contains("-BJL-") AndAlso Not LFN_BJL_NAME_ARRAY.Contains(Row.ItemArray(1).ToString) Then
					Throw New System.Exception("ERROR IN BJL Reading AT: " + FileName)
				ElseIf BJL_AT_CSV = String.Empty AndAlso Row.ItemArray(1).ToString.Contains("-BJL-") Then
					BJL_AT_CSV = Row.ItemArray(1).ToString
				ElseIf SDS_AT_CSV = String.Empty AndAlso Row.ItemArray(0).ToString.Contains("-SDS-") Then
					SDS_AT_CSV = Row.ItemArray(0).ToString
					If SDS_AT_CSV.Contains(LFN_SDS_NAME) Then
						If Double.Parse(Row.ItemArray(6).ToString) > 0 Then
							LFN_SDS_LENGTH = Row.ItemArray(6).ToString
						ElseIf Double.Parse(Row.ItemArray(5).ToString) > 0 Then
							LFN_SDS_LENGTH = Row.ItemArray(5).ToString
						End If
						LFN_SDS_LENGTH = String.Format("{0:F1}", Double.Parse(LFN_SDS_LENGTH.ToString)) + "m"
					End If
				ElseIf LFN_FSL_NAME <> String.Empty AndAlso Row.ItemArray(0).ToString.Contains(LFN_FSL_NAME) Then
					If Double.Parse(Row.ItemArray(6).ToString) > 0 Then
						LFN_FSL_LENGTH = Row.ItemArray(6).ToString
					ElseIf Double.Parse(Row.ItemArray(5).ToString) > 0 Then
						LFN_FSL_LENGTH = Row.ItemArray(5).ToString
					End If
					LFN_FSL_LENGTH = String.Format("{0:F1}", Double.Parse(LFN_FSL_LENGTH.ToString)) + "m"
				ElseIf BJL_AT_CSV <> String.Empty AndAlso UCase(Row.ItemArray(0).ToString).StartsWith("SPLICING SUMMARY") Then
					blnFoundSpliceSummary = True
					'ElseIf BJL_AT_CSV <> String.Empty AndAlso UCase(Row.ItemArray(0).ToString).StartsWith("UNCONNECTED") Then
					'	blnFoundSpliceSummary = False
				ElseIf BJL_AT_CSV <> String.Empty AndAlso blnFoundSpliceSummary = True AndAlso UCase(Row.ItemArray(0).ToString) <> String.Empty AndAlso Not UCase(Row.ItemArray(0).ToString).StartsWith("UNCONNECTED FIBERS") Then
					Dim Read_Cable_IN As String = Row.ItemArray(0).ToString
					Dim Cable_IN_Sequence As String = Chr(70) + Read_Cable_IN.Substring(Read_Cable_IN.IndexOf(":") + 2)
					Read_Cable_IN = Read_Cable_IN.Substring(0, Read_Cable_IN.IndexOf(":"))
					Read_Cable_IN = Read_Cable_IN.Substring(0, 4) + Chr(45) + Read_Cable_IN.Substring(4, 2) + Chr(45) + Read_Cable_IN.Substring(6, 2) + Chr(45) + Read_Cable_IN.Substring(8, 3) + Chr(45) + Read_Cable_IN.Substring(11, 3) + Chr(32) + Cable_IN_Sequence
					Dim Read_Cable_OUT As String = Row.ItemArray(2).ToString
					Dim Cable_OUT_Sequence As String = Chr(70) + Read_Cable_OUT.Substring(Read_Cable_OUT.IndexOf(":") + 2)

					Read_Cable_OUT = Read_Cable_OUT.Substring(0, Read_Cable_OUT.IndexOf(":"))
					Read_Cable_OUT = Read_Cable_OUT.Substring(0, 4) + Chr(45) + Read_Cable_OUT.Substring(4, 2) + Chr(45) + Read_Cable_OUT.Substring(6, 2) + Chr(45) + Read_Cable_OUT.Substring(8, 3) + Chr(45) + Read_Cable_OUT.Substring(11, 3) + Chr(32) + Cable_OUT_Sequence
					BJL_Splice_List.Add(Read_Cable_IN + Chr(32) + Chr(61) + Chr(32) + Read_Cable_OUT)
					'dicLFN_Splice = New List(Of String)
				ElseIf BJL_AT_CSV <> String.Empty AndAlso UCase(Row.ItemArray(0).ToString).StartsWith("UNCONNECTED FIBERS") AndAlso BJL_Splice_List.Count > 0 Then
					dicLFN_BJL_Splice.Add(BJL_AT_CSV, BJL_Splice_List)

					'dicLFN_Splice.Add(LFN_FSL_NAME + Chr(32) + LFN_FSL_Fibre_Sequence + Chr(32) + Chr(61) + Chr(32) + LFN_SSS_NAME + Chr(32) + LFN_SSS_Fibre_Sequence)
					blnFoundSpliceSummary = False
					BJL_Splice_List = New List(Of String)
					BJL_AT_CSV = String.Empty


				End If
			Next
		Next
	End Sub
	Private Shared Sub ReadLFNTraceReport(ByVal DataTableToRead As DataTable)
		LFN_FJL_Name = String.Empty
		LFN_FJL_ADDRESS = String.Empty
		LFN_FJL_SUBURB_STATE = String.Empty
		LFN_BJL_Name = New Dictionary(Of String, List(Of String))
		LFN_BJL_ADDRESS = String.Empty
		LFN_BJL_SUBURB_STATE = String.Empty

		LFN_PCD_Name = String.Empty
		LFN_PCD_ADDRESS = String.Empty
		LFN_PCD_SUBURB_STATE = String.Empty

		LFN_CTL_NAME = String.Empty
		LFN_CTL_ADDRESS = String.Empty
		LFN_CTL_SUBURB_STATE = String.Empty

		LFN_FSL_NAME = String.Empty
		LFN_FSL_Fibre = String.Empty
		LFN_FSL_Fibre_Sequence = String.Empty
		LFN_FSL_LENGTH = String.Empty

		LFN_SMP_Name = String.Empty
		LFN_SMP_ADDRESS = String.Empty
		LFN_SMP_SUBURB_STATE = String.Empty
		LFN_SMP_TYPE = String.Empty
		LFN_SMP_RATIO = String.Empty
		LFN_SMP_SPLITTER_IN = String.Empty
		LFN_SMP_SPLITTER_OUT = String.Empty

		LFN_SSS_NAME = String.Empty
		LFN_SSS_Fibre = String.Empty
		LFN_SSS_Fibre_Sequence = String.Empty
		LFN_SSS_LENGTH = String.Empty

		LFN_SDS_NAME = String.Empty
		LFN_SDS_Fibre = String.Empty
		LFN_SDS_Fibre_Sequence = String.Empty
		LFN_SDS_LENGTH = String.Empty


		LFN_SPLITTER_Name = String.Empty
		LFN_SPLITTER_Branch = String.Empty
		LFN_PIC_NAME = String.Empty
		LFN_PIC_Fibre = String.Empty
		LFN_PIC_LENGTH = String.Empty
		LFN_PIC_SEQUENCE = String.Empty
		LFN_ICD_Name = String.Empty
		LFN_ICD_ADDRESS = String.Empty
		LFN_ICD_SUBURB_STATE = String.Empty
		LFN_PDC_NAME = String.Empty
		LFN_PDC_Fibre = String.Empty
		LFN_PDC_SEQUENCE = String.Empty
		LFN_PDC_LENGTH = String.Empty
		LFN_NTD_Name = String.Empty
		LFN_NTD_ADDRESS = String.Empty
		LFN_NTD_SUBURB_STATE = String.Empty
		Dim ADASearch As String = ADA_Prefix.Substring(0, 4)
		Dim SMPRowNum As Integer = 0
		For Each Row As DataRow In DataTableToRead.Rows
			Dim ValueAtZero As Object = Row.ItemArray(0)
			Dim ValueAtOne As Object = Row.ItemArray(1)
			Dim ValueAtTHREE As Object = Row.ItemArray(3)
			Dim ValueAtSix As Object = Row.ItemArray(6)
			Dim ValueAtSeven As Object = Row.ItemArray(7)

			Select Case True
				Case LFN_FJL_Name = String.Empty AndAlso ValueAtZero IsNot DBNull.Value AndAlso ValueAtZero.ToString.Contains("FJL")
					LFN_FJL_Name = ValueAtZero.ToString.Substring(ValueAtZero.ToString.IndexOf(ADASearch))
					LFN_FJL_ADDRESS = ValueAtOne.ToString
					LFN_FJL_SUBURB_STATE = DesignSuburb.ToString + Chr(44) + Chr(32) + DesignState.ToString + Chr(45) + DesignPostCode.ToString

				Case ValueAtZero IsNot DBNull.Value AndAlso ValueAtZero.ToString.Contains("BJL")
					LFN_BJL_ADDRESS = ValueAtOne.ToString
					LFN_BJL_SUBURB_STATE = DesignSuburb.ToString + Chr(44) + Chr(32) + DesignState.ToString + Chr(45) + DesignPostCode.ToString
					LFN_BJL_Name.Add(ValueAtZero.ToString.Substring(ValueAtZero.ToString.IndexOf(ADASearch)), New List(Of String)({LFN_BJL_ADDRESS, LFN_BJL_SUBURB_STATE, LFN_FSL_NAME, LFN_FSL_Fibre, LFN_FSL_Fibre_Sequence, LFN_FSL_LENGTH}))

					LFN_FSL_NAME = String.Empty
					LFN_FSL_Fibre = String.Empty
					LFN_FSL_Fibre_Sequence = String.Empty
					LFN_FSL_LENGTH = String.Empty


				Case LFN_PCD_Name = String.Empty AndAlso ValueAtZero IsNot DBNull.Value AndAlso ValueAtZero.ToString.Contains("PCD")
					LFN_PCD_Name = ValueAtZero.ToString.Substring(ValueAtZero.ToString.IndexOf(ADASearch))
					LFN_PCD_ADDRESS = DesignLocation.ToString
					LFN_PCD_SUBURB_STATE = DesignSuburb.ToString + Chr(44) + Chr(32) + DesignState.ToString + Chr(45) + DesignPostCode.ToString

				Case LFN_SMP_Name = String.Empty AndAlso ValueAtZero IsNot DBNull.Value AndAlso ValueAtZero.ToString.Contains("SMP")
					LFN_SMP_Name = ValueAtZero.ToString.Substring(ValueAtZero.ToString.IndexOf(ADASearch))
					LFN_SMP_ADDRESS = ValueAtOne.ToString
					LFN_SMP_SUBURB_STATE = DesignSuburb.ToString + Chr(44) + Chr(32) + DesignState.ToString + Chr(45) + DesignPostCode.ToString
					LFN_SMP_TYPE = ValueAtZero.ToString.Substring(0, ValueAtZero.ToString.IndexOf("Port") - 1)
					Dim RatioStartPt As Integer = (ValueAtZero.ToString.IndexOf(Chr(58)) - 1)
					Dim RatioEndPt As Integer = (ValueAtZero.ToString.IndexOf("Splitter):") - 1)
					LFN_SMP_RATIO = ValueAtZero.ToString.Substring(RatioStartPt, RatioEndPt - RatioStartPt)
					SMPRowNum = DataTableToRead.Rows.IndexOf(Row)
				Case LFN_SMP_Name <> String.Empty AndAlso DataTableToRead.Rows.IndexOf(Row) = SMPRowNum + 1
					LFN_SMP_SPLITTER_IN = ValueAtOne.ToString
					LFN_SMP_SPLITTER_IN = LFN_SMP_SPLITTER_IN.Substring(0, 4) + Chr(45) + LFN_SMP_SPLITTER_IN.Substring(4, 2) + Chr(45) + LFN_SMP_SPLITTER_IN.Substring(6, 2) + Chr(45) + LFN_SMP_SPLITTER_IN.Substring(8, 3) + Chr(45) + LFN_SMP_SPLITTER_IN.Substring(11, 3) + Chr(32) + ValueAtTHREE + "1"
				Case LFN_SMP_Name <> String.Empty AndAlso DataTableToRead.Rows.IndexOf(Row) = SMPRowNum + 2
					LFN_SMP_SPLITTER_OUT = ValueAtOne.ToString
					LFN_SMP_SPLITTER_OUT = LFN_SMP_SPLITTER_OUT.Substring(0, 4) + Chr(45) + LFN_SMP_SPLITTER_OUT.Substring(4, 2) + Chr(45) + LFN_SMP_SPLITTER_OUT.Substring(6, 2) + Chr(45) + LFN_SMP_SPLITTER_OUT.Substring(8, 3) + Chr(45) + LFN_SMP_SPLITTER_OUT.Substring(11, 3) + Chr(32) + ValueAtTHREE
				Case LFN_FSL_NAME = String.Empty AndAlso ValueAtOne IsNot DBNull.Value AndAlso ValueAtOne.ToString.Contains("FSL")
					LFN_FSL_NAME = ValueAtOne.ToString.Substring(ValueAtOne.ToString.IndexOf(ADASearch))
					LFN_FSL_Fibre = ValueAtOne.ToString.Substring(0, ValueAtOne.ToString.IndexOf(Chr(32)) - 1)
					LFN_FSL_Fibre = Regex.Replace(LFN_FSL_Fibre, "[A-Za-z_:]", "") + "F"
					LFN_FSL_Fibre_Sequence = "F" + Row.ItemArray(2).ToString
					LFN_FSL_LENGTH = String.Format("{0:F1}", Double.Parse(ValueAtSeven.ToString)) + "m"
				Case LFN_SSS_NAME = String.Empty AndAlso ValueAtOne IsNot DBNull.Value AndAlso ValueAtOne.ToString.Contains("SSS") AndAlso ValueAtOne.ToString.Contains(ADA_Prefix)
					LFN_SSS_NAME = ValueAtOne.ToString.Substring(ValueAtOne.ToString.IndexOf(ADASearch))
					LFN_SSS_Fibre = ValueAtOne.ToString.Substring(0, ValueAtOne.ToString.IndexOf(Chr(32)) - 1)
					LFN_SSS_Fibre = Regex.Replace(LFN_SSS_Fibre, "[A-Za-z_:]", "") + "F"
					LFN_SSS_Fibre_Sequence = "F" + Row.ItemArray(2).ToString
				Case LFN_SDS_NAME = String.Empty AndAlso ValueAtOne IsNot DBNull.Value AndAlso ValueAtOne.ToString.Contains("SDS") AndAlso ValueAtOne.ToString.Contains(ADA_Prefix)
					LFN_SDS_NAME = ValueAtOne.ToString.Substring(ValueAtOne.ToString.IndexOf(ADASearch))
					LFN_SDS_Fibre = ValueAtOne.ToString.Substring(0, ValueAtOne.ToString.IndexOf(Chr(32)) - 1)
					LFN_SDS_Fibre = Regex.Replace(LFN_SSS_Fibre, "[A-Za-z_:]", "") + "F"
					LFN_SDS_Fibre_Sequence = "F" + Row.ItemArray(2).ToString
				Case LFN_FSL_NAME <> String.Empty AndAlso LFN_SPLITTER_Name = String.Empty AndAlso ValueAtSix IsNot DBNull.Value AndAlso ValueAtSix.ToString.Contains("SPL")
					LFN_SPLITTER_Name = Row.ItemArray(6).ToString
					LFN_SPLITTER_Branch = Row.ItemArray(2).ToString

				Case LFN_CTL_NAME = String.Empty AndAlso ValueAtZero IsNot DBNull.Value AndAlso ValueAtZero.ToString.Contains("CTL")
					LFN_CTL_NAME = ValueAtZero.ToString.Substring(ValueAtZero.ToString.IndexOf(ADASearch))

				Case LFN_CTL_NAME <> String.Empty AndAlso ValueAtOne IsNot DBNull.Value AndAlso ValueAtOne.ToString.Contains("PIC")
					LFN_CTL_ADDRESS = ValueAtSix.ToString.Substring(ValueAtSix.ToString.IndexOf(Chr(58)) + 3)
					LFN_CTL_SUBURB_STATE = DesignSuburb.ToString + Chr(44) + Chr(32) + DesignState.ToString + Chr(45) + DesignPostCode.ToString
					LFN_PIC_NAME = ValueAtOne.ToString.Substring(ValueAtOne.ToString.IndexOf(ADASearch))
					LFN_PIC_Fibre = ValueAtOne.ToString.Substring(ValueAtOne.ToString.IndexOf(Chr(58)) - 2, 2)
					LFN_PIC_LENGTH = "0.0m"
					LFN_PIC_SEQUENCE = Row.ItemArray(2).ToString

				Case LFN_PCD_Name <> String.Empty AndAlso ValueAtOne IsNot DBNull.Value AndAlso ValueAtOne.ToString.Contains("PIC")
					LFN_PIC_NAME = ValueAtOne.ToString.Substring(ValueAtOne.ToString.IndexOf(ADASearch))
					LFN_PIC_Fibre = ValueAtOne.ToString.Substring(ValueAtOne.ToString.IndexOf(Chr(58)) - 2, 2)
					LFN_PIC_LENGTH = "0.0m"
					LFN_PIC_SEQUENCE = Row.ItemArray(2).ToString

				Case LFN_ICD_Name = String.Empty AndAlso ValueAtZero IsNot DBNull.Value AndAlso ValueAtZero.ToString.Contains("ICD")
					LFN_ICD_Name = ValueAtZero.ToString.Substring(ValueAtZero.ToString.IndexOf(ADASearch))

				Case LFN_ICD_Name <> String.Empty AndAlso ValueAtOne IsNot DBNull.Value AndAlso ValueAtOne.ToString.Contains("PDC")
					LFN_ICD_ADDRESS = ValueAtSix.ToString.Substring(ValueAtSix.ToString.IndexOf(Chr(58)) + 3)
					LFN_ICD_SUBURB_STATE = DesignSuburb.ToString + Chr(44) + Chr(32) + DesignState.ToString + Chr(45) + DesignPostCode.ToString
					LFN_PDC_NAME = ValueAtOne.ToString.Substring(ValueAtOne.ToString.IndexOf(ADASearch))
					LFN_PDC_Fibre = ValueAtOne.ToString.Substring(ValueAtOne.ToString.IndexOf(Chr(58)) - 2, 2)
					LFN_PDC_LENGTH = "0.0m"
					LFN_PDC_SEQUENCE = Row.ItemArray(2).ToString

				Case LFN_NTD_Name = String.Empty AndAlso ValueAtZero IsNot DBNull.Value AndAlso ValueAtZero.ToString.Contains("NTD")
					LFN_NTD_Name = ValueAtZero.ToString.Substring(ValueAtZero.ToString.IndexOf(ADASearch))
					LFN_NTD_ADDRESS = LFN_ICD_ADDRESS
					LFN_NTD_SUBURB_STATE = LFN_ICD_SUBURB_STATE
			End Select
		Next
	End Sub
	Public Function GetAllWorksheets(ByVal fileName As String) As Sheets
		Dim theSheets As Sheets
		Using document As SpreadsheetDocument = SpreadsheetDocument.Open(fileName, False)
			Dim wbPart As WorkbookPart = document.WorkbookPart
			theSheets = wbPart.Workbook.Sheets()
		End Using
		Return theSheets
	End Function
	Private Shared Sub CreateLFNDataTable()
		SLDdtaTable1 = New DataTable
		SLDdtaTable2 = New DataTable
		Dim CellAddress As String = String.Empty
		Dim SheetName As String = String.Empty
		Dim value As String = Nothing
		Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(Path.Combine(FolderPath, FileName), False)
			Dim wbPart As WorkbookPart = spreadsheetDocument.WorkbookPart
			Dim thesheetcollection As Sheets = wbPart.Workbook.GetFirstChild(Of Sheets)()
			For Each Readsheet As Sheet In thesheetcollection
				SheetName = Readsheet.Name
				Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().Where(Function(s) s.Name = SheetName).FirstOrDefault()
				Dim wsPart As WorksheetPart = CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)
				Dim theWorksheet As Worksheet = wsPart.Worksheet
				Dim thesheetdata As SheetData = theWorksheet.GetFirstChild(Of SheetData)()
				Dim thecurrentrow As IEnumerable(Of Row) = thesheetdata.Descendants(Of Row)()
				For Each cell As Spreadsheet.Cell In thecurrentrow.ElementAt(0)
					CellAddress = cell.CellReference
					If SheetName = "1" Then
						SLDdtaTable1.Columns.Add(GetCellValue(wbPart, wsPart, CellAddress))
					ElseIf SheetName = "2" Then
						SLDdtaTable2.Columns.Add(GetCellValue(wbPart, wsPart, CellAddress))
					End If
				Next
				If SheetName = "1" AndAlso thecurrentrow.ElementAt(0).Count = 8 Then
					SLDdtaTable1.Columns.Add("Branch")
				ElseIf SheetName = "2" AndAlso thecurrentrow.ElementAt(0).Count = 8 Then
					SLDdtaTable2.Columns.Add("Branch")
				End If
				For Each row As Row In thecurrentrow
					Dim dataRow As DataRow = Nothing
					If Integer.Parse(row.RowIndex) > 1 Then
						If SheetName = "1" Then
							dataRow = SLDdtaTable1.NewRow()
						ElseIf SheetName = "2" Then
							dataRow = SLDdtaTable2.NewRow()
						End If
						Dim i As Integer = 0
						While i < row.Descendants(Of Cell)().Count() - 1
							CellAddress = row.Descendants(Of Cell)().ElementAt(i).CellReference
							dataRow(i) = GetCellValue(wbPart, wsPart, CellAddress)
							i += 1
						End While
						If row.ChildElements.Count = 1 Then
							CellAddress = row.Descendants(Of Cell)().ElementAt(i).CellReference
							dataRow(i) = GetCellValue(wbPart, wsPart, CellAddress)
						End If
						Dim TempADA_Prefix As String = ADA_Prefix.Substring(0, 7)
						If dataRow.ItemArray(0) IsNot DBNull.Value AndAlso dataRow.ItemArray(1) IsNot DBNull.Value Then
							If Integer.Parse(row.RowIndex) = 2 AndAlso Not dataRow.ItemArray(1).ToString.Contains(TempADA_Prefix) Then
								Throw New System.Exception("TRACE REPORT NOT ASSOCIATED TO PROJECT: " + TempADA_Prefix)
							End If
						End If
						'End If
						If SheetName = "1" Then
							SLDdtaTable1.Rows.Add(dataRow)
						ElseIf SheetName = "2" Then
							SLDdtaTable2.Rows.Add(dataRow)
						End If
					End If
				Next
			Next
			spreadsheetDocument.Close()
		End Using
	End Sub
	Private Shared Sub CreateDataTable()
		Dim CellAddress As String = String.Empty
		SLDdtaTable1 = New DataTable
		Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(Path.Combine(FolderPath, FileName), False)
			Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
			For Each wsPart As WorksheetPart In workbookPart.WorksheetParts
				Dim xlWSheet As Worksheet = wsPart.Worksheet
				Dim sheetData As SheetData = xlWSheet.GetFirstChild(Of SheetData)()
				Dim rows As IEnumerable(Of Spreadsheet.Row) = sheetData.Descendants(Of Spreadsheet.Row)()
				For Each cell As Spreadsheet.Cell In rows.ElementAt(0)
					CellAddress = cell.CellReference
					SLDdtaTable1.Columns.Add(GetCellValue(workbookPart, wsPart, CellAddress))
				Next
				If rows.ElementAt(0).Count = 8 Then
					SLDdtaTable1.Columns.Add("Branch")
				End If
				For Each row As Row In rows
					If Integer.Parse(row.RowIndex) > 1 Then
						Dim dataRow As DataRow = SLDdtaTable1.NewRow()
						Dim i As Integer = 0
						While i < row.Descendants(Of Cell)().Count() - 1
							CellAddress = row.Descendants(Of Cell)().ElementAt(i).CellReference
							dataRow(i) = GetCellValue(workbookPart, wsPart, CellAddress)
							i += 1
						End While
						If row.ChildElements.Count = 1 Then
							CellAddress = row.Descendants(Of Cell)().ElementAt(i).CellReference
							dataRow(i) = GetCellValue(workbookPart, wsPart, CellAddress)
						End If
						Dim TempADA_Prefix As String = ADA_Prefix.Substring(0, 7)
						If dataRow.ItemArray(0) IsNot DBNull.Value AndAlso dataRow.ItemArray(1) IsNot DBNull.Value Then
							If Integer.Parse(row.RowIndex) = 2 AndAlso Not dataRow.ItemArray(1).ToString.Contains(TempADA_Prefix) Then
								Throw New System.Exception("TRACE REPORT NOT ASSOCIATED TO PROJECT: " + TempADA_Prefix)
							End If
						End If
						SLDdtaTable1.Rows.Add(dataRow)
					End If
				Next
			Next
			spreadsheetDocument.Close()
		End Using
	End Sub
	Private Shared Sub OpenTraceReport()
		FileName = String.Empty
		Do Until UCase(FileName).Contains("TRACE REPORT")
			Using OFDialog As New System.Windows.Forms.OpenFileDialog()
				OFDialog.Title = "SELECT Trace Report File"
				OFDialog.Multiselect = True
				OFDialog.Filter = "CSV Files(*.csv)|*.csv|xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
				OFDialog.FilterIndex = 3
				OFDialog.RestoreDirectory = True
				If OFDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
					FolderPath = Path.GetDirectoryName(OFDialog.FileName)
					FileName = Path.GetFileName(OFDialog.FileName)
				ElseIf System.Windows.Forms.DialogResult.Cancel Then
					Throw New System.Exception("SLD GENERATION CLOSED ABRUPTLY")
					Exit Sub
				End If
				If Not FileName.Contains("Trace Report") Then
					MsgBox("WRONG TRACE REPORT", vbOKOnly + vbInformation, "WRONG TRACE REPORT")
				End If
			End Using
		Loop
	End Sub
	Public Shared Sub StartBDOD_DFN_SLD()
		OpenTraceReport()
		PrimaryTraceFile = FileName
		CreateDataTable()
		OpenAndReadTraceReeport()
		PopulateCable_IN()
		Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
		StartSLD_IN_CAD()
		If blnLTCFound = False Then
			UpdatePONPatchTemplate()
		End If
	End Sub
	Public Shared Sub StartBDOD_LFN_SLD()
		OpenTraceReport()
		PrimaryTraceFile = FileName
		CreateLFNDataTable()
		RepeatCount = 1
		If SLDdtaTable2.Rows.Count > 0 Then
			RepeatCount = 2
		End If
		Dim DataTableToRead As DataTable = SLDdtaTable1
        Dim blnIgnorePrimarySLD As Boolean = False
        '******************************************
        'Dim dr As System.Data.DataRow() = DataTableToRead.Select("User like '%-DJL-%'")
        'Dim DJLName As String = dr(0).ItemArray.ElementAt(0)
        'Dim Filter_Row As DataRow() = DataTableToRead.Select("User = '" & DJLName & "'")
        'Dim Rowindex As Integer = DataTableToRead.Rows.IndexOf(Filter_Row(0))
        'Dim DSS_Name As String = DataTableToRead.Rows(Rowindex - 1).Item("Start entity")
        ''Dim SelColumns As DataColumnCollection = DataTableToRead.Columns
        ''Dim Colindex As Integer
        ''If SelColumns.Contains(DJLName) Then
        ''    Colindex = SelColumns.IndexOf(DJLName)
        ''End If
        ''Dim Colindex As Integer = SelColumns.IndexOf(DJLName)
        'MsgBox("Row Num: " + Rowindex.ToString + vbLf + "Column Bum: ")
        '****************************************
        Do While RepeatCount > 0
			ReadLFNTraceReport(DataTableToRead)
			If blnIgnorePrimarySLD = False Then
				CaptureLFNCableTable()
				Start_LFN_SLD_IN_CAD()

				If RepeatCount = 2 Then
					blkLFCTLInsPt = New Point3d(blkLFCTLInsPt.X, blkLFCTLInsPt.Y + 3, blkLFCTLInsPt.Z)
				End If
				blnIgnorePrimarySLD = True
			End If
			Start_LFN_SLD_AFTER_CTL_IN_CAD()
			DataTableToRead.Clear()
			DataTableToRead = SLDdtaTable2
			RepeatCount -= 1
			blkLFCTLInsPt = New Point3d(blkLFCTLInsPt.X, blkLFCTLInsPt.Y - 6, blkLFCTLInsPt.Z)
		Loop
	End Sub
End Class
