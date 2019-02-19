const My_SQL_Server = ""
const My_SQL_DB = ""
const DES_DEI = ""
const DES_DEI_STC_UR_DIR = ""
const DES_DEI_COLs = "VENDOR, PLATFORM, DISPUTE_CATEGORY, BILL_DATE, DISPLAY_STATUS, STC_CLAIM_NUMBER, ACCOUNT_NUMBER, DATE_SUBMITTED, ILEC_CONFIRMATION, ILEC_COMMENTS, DISPUTE_AMOUNT, CREDIT_APPROVED, DENIED, UNRESOLVED_AMOUNT, CREDIT_RECEIVED_INVOICE_DATE, CREDIT_RECEIVED_AMOUNT, EXCESS_CREDIT, ESCALATE, ESCALATE_DATE, ESCALATE_AMOUNT, CLOSE_ESCALATE_REASON, WTN, DISPUTE_REASON, RESOLUTION_DATE, DATE_UPDATED, [INDEX]"
const DES_STC_COLs = "[Vendor], [Platform], [Dispute Category ], [Bill Date], [Display Status], [STC Claim Number], [Account Number (BAN)], [Date Submitted], [ILEC Confirmation], [ILEC Comments], [Dispute Amount], [Credit Approved], [Denied], [Unresolved Amount], [Credit Received - Invoice Date], [Credit Received - Amount], [Excess Credit], [Escalate], [Escalate Date], [Escalate Amount], [Close/Escalate Reason], [WTN], [Dispute reason], [Resolution Date], [Date Updated], [Index]"

Public WshShell, oFSO, Log_Path, networkInfo, DES_DEI_STC_UR_DIR_PATH, Batch, My_List, My_SQL, SQL_Filepath
dim Temp, myquery, oFolder, oFileCollection, oFile, myitems(1), Q, mydata, SQL_List(4)

Set WshShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set networkInfo = CreateObject("WScript.NetWork")

select case Weekday(Now())
	case 1
		Temp = DateAdd("d", -2, Now())
	case 2
		Temp = DateAdd("d", -3, Now())
	case 3
		Temp = DateAdd("d", -4, Now())
	case 4
		Temp = DateAdd("d", -5, Now())
	case 5
		Temp = DateAdd("d", -6, Now())
	case 6
		Temp = DateAdd("d", 0, Now())
	case 7
		Temp = DateAdd("d", -1, Now())
end select

if len(cstr(month(Temp)))=2 and len(cstr(day(Temp)))=2 then
	batch = cstr(year(Temp)) & cstr(month(Temp)) & cstr(day(Temp))
end if
if len(cstr(month(Temp)))=2 and not len(cstr(day(Temp)))=2 then
	batch = cstr(year(Temp)) & cstr(month(Temp)) & "0" & cstr(day(Temp))
end if
if not len(cstr(month(Temp)))=2 and len(cstr(day(Temp)))=2 then
	batch = cstr(year(Temp)) & "0" & cstr(month(Temp)) & cstr(day(Temp))
end if
if not len(cstr(month(Temp)))=2 and not len(cstr(day(Temp)))=2 then
	batch = cstr(year(Temp)) & "0" & cstr(month(Temp)) & "0" & cstr(day(Temp))
end if

Set oFolder = oFSO.GetFolder(DES_DEI_STC_UR_DIR)
Set oFileCollection = oFolder.Files

For each oFile in oFileCollection
	if (oFSO.getextensionname(Cstr(oFile.Name)) = "mdb" or oFSO.getextensionname(Cstr(oFile.Name)) = "accdb") and instr(1, oFile.Name, batch) > 0 and instr(1, oFile.Name, "Granite Dispute Tracking") > 0 then
		if len(myitems(0)) > 0 then
			if myitems(0) < oFile.DateLastModified then
				myitems(0) = oFile.DateLastModified
				myitems(1) = oFile.Name
			end if
		else
			myitems(0) = oFile.DateLastModified
			myitems(1) = oFile.Name
		end if
	end if
Next

set oFolder = nothing
Set oFileCollection = Nothing
Set oFile = Nothing

Log_Path = replace(WScript.ScriptFullName,".vbs","") & "_Log.txt"

if len(myitems(1)) > 0 then
	DES_DEI_STC_UR_DIR_PATH = DES_DEI_STC_UR_DIR & myitems(1)
	SQL_Filepath = oFSO.GetParentFolderName(WScript.ScriptFullName) & "\SQL_Scripts\DEI_SQL.sql"

	Append_ODS "truncate table " & DES_DEI

	myquery = "select distinct * from (select " & DES_STC_COLs & " from [Granite Dispute Tracking Dbase Updated Records])"

	Query_AccDB myquery, My_List

	if IsArray(My_List) then
		Q = Ceil((ubound(My_List, 2) + 1) / 1000) - 1
		redim My_SQL(Q)

		for Q = lbound(My_SQL, 1) to ubound(My_SQL, 1)
			if ((Q + 1) * 1000) - 1 > ubound(My_List, 2) then
				Append_Records (Q * 1000), ubound(My_List, 2), mydata
			else
				Append_Records (Q * 1000), ((Q + 1) * 1000) - 1, mydata
			end if
			My_SQL(Q) = mydata
		Next
		set My_List = nothing

		for Q = lbound(My_SQL, 1) to ubound(My_SQL, 1)
			SQL_List(Q mod 5) = My_SQL(Q)
			if Q > 3 and Q mod 5 = 4 then
				Append_ODS join(SQL_List, chr(13))
				SQL_List(0) = ""
				SQL_List(1) = ""
				SQL_List(2) = ""
				SQL_List(3) = ""
				SQL_List(4) = ""
			end if
		next

		if not Q mod 5 = 0 then
			Append_ODS join(SQL_List, chr(13))
		end if

		set My_SQL = nothing

		Append_ODS "update " & DES_DEI & " set EDIT_DT=getdate(), SOURCE_FILE='Updated Records " & batch & "' where EDIT_DT is null"

		Execute_SQL

		msgbox("Upload is now Finished!")
	else
		write_log Now() & " * Error * " & networkInfo.UserName & " * No Data found for workbook (" & myitems(1) & ")"
	end if
else
	write_log Now() & " * Error * " & networkInfo.UserName & " * Extract Updated Records hasn't been unzipped or doesn't have 'Granite Dispute Tracking' with batch date and .mdb/.accdb format. Plz unzip/fix"
end if
	
sub Query_AccDB(myquery, ReturnArray)
	On Error Resume Next
	Dim constr, conn, rs, myresults
	set ReturnArray = nothing

	constr = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & DES_DEI_STC_UR_DIR_PATH & ";Exclusive=1"

	Set conn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")

	conn.Open constr

	If Err.Number <> 0 Then
		write_log Now() & " * Error * " & networkInfo.UserName & " * Open ACC SQL Conn (" & Err.Description & ")"
		Set conn = Nothing
		exit sub
	end if

	conn.CommandTimeout = 0

	rs.Open myquery, conn

	If Err.Number <> 0 Then
		write_log Now() & " * Error * " & networkInfo.UserName & " * ACC SQL Query (" & Err.Description & ")"
		Set conn = Nothing
		exit sub
	end if

	If Not rs.EOF Then
        	ReturnArray = rs.getrows()
    	End If
    
	rs.Close

	conn.Close

    	Set rs = Nothing
	Set conn = Nothing
end sub

sub Append_Records(mystart, myend, mydata)
	on error resume next
	dim My_Records, My_Temp, S, T, myline, Temp_List, myitem

	S = cint(ubound(My_List, 1))

	redim My_Temp(S)
	S = myend - mystart
	redim My_Records(S)
	

	for S = mystart to myend
		myline = S - mystart
		for T = lbound(My_List, 1) to ubound(My_List, 1)
			myitem = My_List(T, S)

			if isnull(myitem) or isempty(myitem) then
				My_Temp(T) = "NULL"
			else
				My_Temp(T) = replace(trim(myitem),"'","''")
			end if
		next
		if isarray(My_Temp) then
			My_Records(myline) = "(" & replace("'" & join(My_Temp, "', '") & "'", "'NULL'", "NULL") & ")"
		end if
	next
	mydata = "insert into " & DES_DEI & " (" & DES_DEI_COLs & ") values " & join(My_Records, ",") & ";"
end sub

Sub Append_ODS(myquery)
	On Error Resume Next
	Dim constr, conn

	constr = "Provider=SQLOLEDB;Data Source=" & My_SQL_Server & ";Initial Catalog=" & My_SQL_DB & ";Integrated Security=SSPI;"
	Set conn = CreateObject("ADODB.Connection")

	conn.Open constr

	If Err.Number <> 0 Then
		write_log Now() & " * Error * " & networkInfo.UserName & " * Open SQL Conn (" & Err.Description & ")"
		Set conn = Nothing
		exit sub
	end if

	conn.CommandTimeout = 0

	conn.Execute myquery

	If Err.Number <> 0 Then
		write_log Now() & " * Error * " & networkInfo.UserName & " * SQL Execute Query (" & Err.Description & ")"
		Set conn = Nothing
		exit sub
	end if
    
	conn.Close

	If Err.Number <> 0 Then
		write_log Now() & " * Error * " & networkInfo.UserName & " * SQL Close Con (" & Err.Description & ")"
	end if
    
	Set conn = Nothing
End Sub

Sub Execute_SQL()
	Dim objFile, strLine

	if oFSO.fileexists(SQL_Filepath) then
		Set objFile = oFSO.OpenTextFile(SQL_Filepath)
		Do Until objFile.AtEndOfStream
			if len(strLine) > 0 then
				strLine= strLine & vbcrlf & objFile.ReadLine
			else
    				strLine= objFile.ReadLine
			end if
		Loop
		objfile.close

		Append_ODS strLine
	else
		write_log Now() & " * Error * " & networkInfo.UserName & " * SQL Script (" & SQL_Filepath & ") does not exist"
	end if

	set objFile = Nothing
end sub

Sub Write_Log(ByVal text)
	Dim objFile, strLine

	if oFSO.fileexists(Log_Path) then
		Set objFile = oFSO.OpenTextFile(Log_Path)
		Do Until objFile.AtEndOfStream
			if len(strLine) > 0 then
				strLine= strLine & vbcrlf & objFile.ReadLine
			else
    				strLine= objFile.ReadLine
			end if
		Loop
		objfile.close
		Set objfile = oFSO.CreateTextFile(Log_Path,True)
		objfile.write strLine & vbcrlf & text
		objfile.close
	else

		Set objfile = oFSO.CreateTextFile(Log_Path,True)
		objfile.write text
		objfile.close
	end if

	msgbox(text)

	set objFile = Nothing
End Sub

Function Ceil(x)
    If Round(x) = x Then
        Ceil = x
    Else
        Ceil = Round(x + 0.5)
    End If
End Function

Function IsArray(anArray)
    Dim I
    On Error Resume Next
    I = UBound(anArray, 1)
    If Err.Number = 0 Then
        IsArray = True
    Else
        IsArray = False
    End If
End Function
