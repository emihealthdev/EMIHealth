strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService. _
    ExecQuery("Select * from CIM_DataFile where Path = '\\Users\\Tom\\Documents\\Inotek\\EMI\\Working\\'and Drive ='C:'")

For Each objFile in colFiles
	strCurrentFile = objFile.Name
'	Wscript.Echo strCurrentFile
' Read a Text File Character-by-Character
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strCurrentFile, 1)
	Set objFile2 = objFSO.GetFile(strCurrentFile)
	strFileName = objFSO.GetFileName(objFile2)
	
	' Initialize fields that aren't always populated
	strBrokerName = ""
	strBrokerTIN = ""
	strBrokerLicense = ""
	strExchangeAssignedPolID = ""
	strCarrierSubscriberID = ""
	strCarrierDependentID = ""
	strReferenceNo = "NONE"
	strSSN = "SSSSSSSSS"
	strPlanID = "NONE"
	strCarrierPlanID = ""
	strEndDate = ""
	strResAddrLine2 = ""
	strResCounty = ""
	strMailAddrLine1 = "" 
	strMailAddrLine2 = "" 
	strMailCity = "" 
	strMailState = "" 
	strMailZip = ""
	strPrimaryPhone = "" 
	strPrimaryEmail = "" 
	strAltPhone = ""
	strRPLName = ""
	strRPFName = ""
	strRPSSN = ""
	strRPRelationship = ""
	strRPAddrLine1 = "" 
	strRPAddrLine2 = "" 
	strRPCity = "" 
	strRPState = "" 
	strRPZip = ""
	
	Do Until objFile.AtEndOfStream
	    strCharacters = objFile.ReadLine()
'	    strBroker = ""
	    If Left(strCharacters,4) = "ISA*" Then
	    	strHeader = Trim(Piece(strCharacters,"*",9,9)) & "*" _
	    	& piece(strCharacters,"*",10,10) & "*" _
	    	& piece(strCharacters,"*",11,11) & "*" _ 
	    	& piece(strCharacters,"*",14,14) & "*" _
	    	& piece(strCharacters,"*",16,16)
	    End If 
	    If Left(strCharacters,3) = "ST*" Then
	    		strEnvelope = piece(strCharacters,"*",3,3) 
	    End If
	    If Left(strCharacters,6) = "QTY*TO" Then
			strEnvelope = strEnvelope & "*" _
			& piece(strCharacters,"*",3,3) 
		End If
		If Left(strCharacters,6) = "REF*38" Then
			strExchangeGroupID = piece(strCharacters,"*",3,3)
		End If
		If Left(strCharacters,6) = "REF*1L" Then
			strExchangeAssignedPolID = piece(strCharacters,"*",3,3)
		End If
		If Left(strCharacters,5) = "N1*P5" Then
			strSponsor = piece(strCharacters,"*",3,3) & "*"_
			& piece(strCharacters,"*",5,5)
		End If
		If Left(strCharacters,5) = "N1*IN" Then
			strPayer = piece(strCharacters,"*",3,3) & "*"_
			& piece(strCharacters,"*",5,5)
		End If
		If Left(strCharacters,5) = "N1*BO" Then
			strBrokerName = piece(strCharacters,"*",3,3)
			strBrokerTIN = piece(strCharacters,"*",5,5)
		End If
		If Left(strCharacters,4) = "ACT*" Then
			strBrokerLicense = Piece(strCharacters,"*",2,2)
		End If
		If Left(strCharacters,4) = "INS*" Then
			strIsSubscriber = Piece(strCharacters,"*",2,2)
		End If
		If Left(strCharacters,4) = "INS*" Then
			strRelation = Piece(strCharacters,"*",3,3)
		End If
		If Left(strCharacters,4) = "INS*" Then
			strMaintType = Piece(strCharacters,"*",4,4)
		End If
		If Left(strCharacters,4) = "INS*" Then
			str2750 = "NONE"
			strDependentID = ""
		End If
		If Left(strCharacters,6) = "REF*0F" Then
			strSubscriberID = piece(strCharacters,"*",3,3)
		End If
		If Left(strCharacters,6) = "REF*ZZ" Then
			strCarrierSubscriberID = piece(strCharacters,"*",3,3)
		End If
		If Left(strCharacters,6) = "REF*17" Then
			If IsNumeric(Replace(Piece(strCharacters,"*",3,3),"~","")) _
				And str2750 = "NONE" Then
				strDependentID = Piece(strCharacters,"*",3,3)
			End If
		End If
		If Left(strCharacters,6) = "REF*23" Then
			strCarrierDependentID = Piece(strCharacters,"*",3,3)
		End If
		If Left(strCharacters,4) = "HLH*" Then
			strTobaccoUse = piece(strCharacters,"*",2,2)
		End If
		If Left(strCharacters,6) = "PER*IP" Then
		'PER03
			If Piece(strCharacters,"*",4,4) = "TE" Then
				strPrimaryPhone = Piece(strCharacters,"*",5,5)
			End If
			If Piece(strCharacters,"*",4,4) = "EM" Then
				strPrimaryEmail = Piece(strCharacters,"*",5,5)
			End If
			If Piece(strCharacters,"*",4,4) = "AP" Then
				strAltPhone = Piece(strCharacters,"*",5,5)
			End If
		'PER05
			If Piece(strCharacters,"*",6,6) = "TE" Then
				strPrimaryPhone = Piece(strCharacters,"*",7,7)
			End If
			If Piece(strCharacters,"*",6,6) = "EM" Then
				strPrimaryEmail = Piece(strCharacters,"*",7,7)
			End If
			If Piece(strCharacters,"*",6,6) = "AP" Then
				strAltPhone = Piece(strCharacters,"*",7,7)
			End If	
		'PER07
			If Piece(strCharacters,"*",8,8) = "TE" Then
				strPrimaryPhone = Piece(strCharacters,"*",9,9)
			End If
			If Piece(strCharacters,"*",8,8) = "EM" Then
				strPrimaryEmail = Piece(strCharacters,"*",9,9)
			End If
			If Piece(strCharacters,"*",8,8) = "AP" Then
				strAltPhone = Piece(strCharacters,"*",9,9)
			End If				
		End If
	'Logic to check for subscriber
		If strIsSubscriber = "Y" Then
			If Left(strCharacters,6) = "REF*6O" Then
				strReferenceNo = piece(strCharacters,"*",3,3)
			End If
			If str2750 =  "APTCAmt" Then
				If Left(strCharacters,6) = "REF*9V" Then
					strAPTCAmt = piece(strCharacters,"*",3,3)
				End If
			End If
			If str2750 =  "CSRAmt" Then
				If Left(strCharacters,6) = "REF*9V" Then
					strCSRAmt = piece(strCharacters,"*",3,3)
				End If
			End If
			
			If str2750 =  "RateArea" Then
				If Left(strCharacters,6) = "REF*9X" Then
					strRateArea = piece(strCharacters,"*",3,3)
				End If
			End If
			If str2750 =  "TotResAmt" Then
				If Left(strCharacters,6) = "REF*9V" Then
					strTotResAmt = piece(strCharacters,"*",3,3)
				End If
			End If
			If str2750 =  "PreAmtTot" Then
				If Left(strCharacters,6) = "REF*9X" Then
					strPreAmtTot = piece(strCharacters,"*",3,3)
				End If
			End If
			If str2750 =  "TotEmpResAmt" Then
				If Left(strCharacters,6) = "REF*9V" Then
					strPreAmtTot = piece(strCharacters,"*",3,3)
				End If
			End If
		End If
		If Left(strCharacters,6) = "NM1*IL" Then
			strNM1 = piece(strCharacters,"*",4,4) & "*" _
			& piece(strCharacters,"*",5,5) 
			strSSN = piece(strCharacters,"*",10,10)
		End If
		If Left(strCharacters,4) = "DMG*" Then
			strDOB = piece(strCharacters,"*",3,3) 
			strGender = piece(strCharacters,"*",4,4)
                        strRace = piece(strCharacters,"*",6,6)
		End If
		If Left(strCharacters,6) = "REF*CE" Then
			strPlanID = piece(strCharacters,"*",3,3)
		End If
		If Left(strCharacters,6) = "REF*X9" Then
			strCarrierPlanID = piece(strCharacters,"*",3,3)
		End If
		If Left(strCharacters,7) = "DTP*348" Then
			strEffDate = piece(strCharacters,"*",4,4)
		End If
		If Left(strCharacters,7) = "DTP*356" Then
			strEnrGroupStartDate = piece(strCharacters,"*",4,4)
		End If
		If Left(strCharacters,7) = "DTP*357" Then
			strEnrGroupEndDate = piece(strCharacters,"*",4,4)
		End If
		If Left(strCharacters,7) = "DTP*349" Then
			strEndDate = piece(strCharacters,"*",4,4)
		End If
		'Determine which 2750 Loop we are in
		If strCharacters =  "N1*75*APTC AMT~" Then
			str2750 = "APTCAmt"
		End If
		If strCharacters =  "N1*75*CSR AMT~" Then
			str2750 = "CSRAmt"
		End If
		If strCharacters =  "N1*75*PRE AMT 1~" Then
			str2750 = "PreAmt"
		End If
		If strCharacters =  "N1*75*RATING AREA~" Then
			str2750 = "RateArea"
		End If
		If strCharacters =  "N1*75*TOT RES AMT~" Then
			str2750 = "TotResAmt"
		End If
		If strCharacters =  "N1*75*PRE AMT TOT~" Then
			str2750 = "PreAmtTot"
		End If
		If strCharacters =  "N1*75*TOT EMP RES AMT~" Then
			str2750 = "TotEmpResAmt"
		End If
		If strCharacters =  "N1*75*REQUEST SUBMIT TIMESTAMP~" Then
			str2750 = "ReqSubTimestamp"
		End If
		If strCharacters =  "N1*75*SOURCE EXCHANGE ID~" Then
			str2750 = "SourceExchangeID"
		End If
		If strCharacters =  "N1*75*ADDL MAINT REASON~" Then
			str2750 = "AddlMaintReason"
		End If
		'Determine which address loop we are in		
		If Left(strCharacters,6) =  "NM1*IL" Then
			strAddrLoop = "Residence"
		End If
		If Left(strCharacters,6) =  "NM1*31" Then
			strAddrLoop = "Mailing"
		End If
		If Left(strCharacters,6) =  "NM1*QD" Then
			strAddrLoop = "ResponsibleParty"
		End If
		If Left(strCharacters,6) =  "NM1*S1" Then
			strAddrLoop = "ResponsibleParty"
		End If
		'Get 2750 values 
		If str2750 =  "PreAmt" Then
				If Left(strCharacters,6) = "REF*9X" Then
					strPreAmt = piece(strCharacters,"*",3,3)
				End If
		End If
		'Get Address values 
		If strAddrLoop =  "Residence" Then
				If Left(strCharacters,3) = "N3*" Then
					strResAddrLine1 = piece(strCharacters,"*",2,2)
					strResAddrLine2 = piece(strCharacters,"*",3,3)
				End If
				If Left(strCharacters,3) = "N4*" Then
					strResCity = piece(strCharacters,"*",2,2)
					strResState = piece(strCharacters,"*",3,3)
					strResZip = piece(strCharacters,"*",4,4)
					strResCounty = piece(strCharacters,"*",7,7)
				End If				
		End If
		If strAddrLoop =  "Mailing" Then
				If Left(strCharacters,3) = "N3*" Then
					strMailAddrLine1 = piece(strCharacters,"*",2,2)
					strMailAddrLine2 = piece(strCharacters,"*",3,3)
				End If
				If Left(strCharacters,3) = "N4*" Then
					strMailCity = piece(strCharacters,"*",2,2)
					strMailState = piece(strCharacters,"*",3,3)
					strMailZip = piece(strCharacters,"*",4,4)
				End If				
		End If
		If strAddrLoop =  "ResponsibleParty" Then
				If Left(strCharacters,3) = "N3*" Then
					strRPAddrLine1 = piece(strCharacters,"*",2,2)
					strRPAddrLine2 = piece(strCharacters,"*",3,3)
				End If
				If Left(strCharacters,3) = "N4*" Then
					strRPCity = piece(strCharacters,"*",2,2)
					strRPState = piece(strCharacters,"*",3,3)
					strRPMailZip = piece(strCharacters,"*",4,4)
				End If				
		End If
		'Get Responsible Party Name
		If Left(strCharacters,6) = "NM1*S1" Then
			strRPRelationship = piece(strCharacters,"*",2,2)
			strRPLName = piece(strCharacters,"*",4,4)
			strRPFName = piece(strCharacters,"*",5,5)
			strRPSSN = piece(strCharacters,"*",10,10)
		End If
		If Left(strCharacters,6) = "NM1*QD" Then
			strRPRelationship = piece(strCharacters,"*",2,2)
			strRPLName = piece(strCharacters,"*",4,4)
			strRPFName = piece(strCharacters,"*",5,5)
			strRPSSN = piece(strCharacters,"*",10,10)
		End If
			
'		WScript.Echo str2750
'		WScript.Echo strDependentID
		
		If Left(strCharacters,7) = "LE*2700" Then
				
'			Const ForAppending = 8

			Set objTextFile = objFSO.OpenTextFile _
    			("C:\Users\Tom\Documents\Inotek\EMI\Temp\FFM834.txt", 8, True)
			strLineEnd = vbcrLF
'			strNewLine = strNewLine & Replace(strFileName & "*" & strHeader & "*" & strEnvelope & "*" _
			strNewLine = Replace(strFileName & "*" & strHeader & "*" & strEnvelope & "*" & strExchangeGroupID & "*" _
			& strExchangeAssignedPolID & "*" & strSponsor & "*" & strPayer & "*" & strBrokerName & "*" _
			& strBrokerTIN & "*" & strBrokerLicense & "*" & strIsSubscriber & "*" & strRelation & "*" & strTobaccoUse & "*" _
			& strMaintType & "*"& strSubscriberID & "*" & strReferenceNo & "*" & strCarrierSubscriberID & "*" & strCarrierDependentID & "*" _
			& strEnrGroupStartDate & "*" & strEnrGroupEndDate & "*" _
			& strDependentID & "*" & strNM1 & "*" & strSSN & "*" & strDOB & "*" & strGender & "*" & strRace & "*" & strPlanID & "*" & strCarrierPlanID & "*" _
			& strEffDate & "*" & strEndDate & "*" & strAPTCAmt & "*" & strCSRAmt & "*" & strPreAmt & "*" & strRateArea & "*" & strTotResAmt & "*" _
			& strPreAmtTot & "*" & strResAddrLine1 & "*" & strResAddrLine2 & "*" & strResCity & "*" & strResState & "*" & strResZip & "*" & strResCounty & "*" _
			& strMailAddrLine1 & "*" & strMailAddrLine2 & "*" & strMailCity & "*" & strMailState & "*" & strMailZip & "*" _
			& strRPRelationship & "*" & strRPLName & "*" & strRPFName & "*" & strRPSSN & "*" _
			& strRPAddrLine1 & "*" & strRPAddrLine2 & "*" & strRPCity & "*" & strRPState & "*" & strRPZip & "*" _
			& strPrimaryPhone & "*" & strPrimaryEmail & "*" & strAltPhone,"~","")
'			& strLineEnd ,"~","")
			objTextFile.WriteLine(strNewLine)
			objTextFile.Close
			'Clear values from fields that are not always populated
			If Left(strExchangeGroupID,9) = "COH-INDV1" Then
				strBrokerName = ""
				strBrokerTIN = ""
				strBrokerLicense = ""
			End If
			strExchangeAssignedPolID = ""
			strCarrierSubscriberID = ""
			strCarrierDependentID = ""
			strSSN = "SSSSSSSSS"
'			strReferenceNo = "NONE"
			strPlanID = "NONE"
			strCarrierPlanID = ""
			strEnrGroupStartDate = ""
			strEnrGroupEndDate = ""
			strEndDate = ""
			strResAddrLine2 = ""
			strResCounty = ""
			strPrimaryPhone = "" 
			strPrimaryEmail = "" 
			strAltPhone = ""
			strMailAddrLine1 = "" 
			strMailAddrLine2 = "" 
			strMailCity = "" 
			strMailState = "" 
			strMailZip = ""
			strRPLName = ""
			strRPFName = ""
			strRPSSN = ""
			strRPRelationship = ""
			strRPAddrLine1 = "" 
			strRPAddrLine2 = "" 
			strRPCity = "" 
			strRPState = "" 
			strRPZip = ""
		End If
	
	Loop

Next


'    WScript.echo strNewLine
'    WScript.Echo strNM1
	WScript.Quit(returnValue)


Function Piece(Searchstring, Separator, Index1, Index2)
Dim t, IndexCount
Piece = ""
t = Split(Searchstring, Separator)
If UBound(t) + 1 < Index1 Then Exit Function
If UBound(t) + 1 < Index2 Then Index2 = UBound(t) + 1
If Index2 = 0 Or Index2 <= Index1 Then
    If UBound(t) > 0 Then Piece = t(Index1 - 1)
  
Else
    For IndexCount = Index1 To Index2
        Piece = Piece & t(IndexCount - 1)
        If IndexCount <> Index2 Then Piece = Piece & Separator
        
    Next 'IndexCount
End If
End Function