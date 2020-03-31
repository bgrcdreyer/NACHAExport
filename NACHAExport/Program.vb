Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Net.Mail
Imports System.Net
Imports System.Data
Module Program
    Dim SectionError As String = "SectionError Initialized"
    Public Sub Main(args As String())
        Try
            BuildNacha()
        Catch ex As Exception
            Console.WriteLine(ex)
            Console.WriteLine(SectionError)
        End Try
        Console.ReadKey()
    End Sub

    Private Sub BuildNacha()

        Dim strBuffer As String = ""
        Dim strBGRTaxID As String = "310841900"        'BGR Inc. Tax ID
        Dim strBankName As String = "STOCK YARDS BANK"
        Dim strCompName As String = "BGR INC"
        Dim strBGRRouting As String = "083000564"
        Dim strFedRouting As String = "081000045"
        Dim intBatchCnt As Integer = 1
        Dim intBatchRows As Integer = 0
        Dim dblDebitTotal As Double = 0
        Dim dblCreditTotal As Double = 0
        Dim dblAmount As Double = 0
        Dim lngEntryHash As Long = 0
        Dim intLength As Integer = 0
        Dim i As Long = 0
        Dim intDetailLineCount As Integer = 0
        Dim intTotalLineCount As Integer = 0
        Dim strTotal As String = ""
        Dim intBlockCount As Integer = 0
        Dim intDummyLines As Integer = 0
        Dim strPayLots As String = ""

        SectionError = "Creating FileStream"
        Dim path As String = "C:\NACHA_" + Now().ToString("yyMMddHHmm") + ".txt"
        Dim NACHA As System.IO.FileStream = File.Create(path)

        SectionError = "Building Data Table"
        Dim SQLStatement As String = <Sql><![CDATA[SELECT p.NUM_0 As Number
	                                                    ,z.BPSNAM_0 As BPSNAM_0
	                                                    ,z.RTGNUM_0 As RTGNUM_0
	                                                    ,z.ACCNUM_0 As ACCNUM_0
	                                                    ,d.AMTLIN_0 As AMTCUR_0
	                                                    ,p.PAYLOT_0 As PAYLOT_0
                                                    FROM [x3v6].[PILOT].[PAYMENTH] p
                                                    INNER JOIN PILOT.ZACH z ON p.BPR_0 = z.BPSNUM_0
                                                    INNER JOIN PILOT.PAYMENTD d ON p.NUM_0 = d.NUM_0
                                                    WHERE p.PAM_0 = 'WTR' and p.BAN_0 = 'SYBGR' and p.PAYTYP_0 = 'PAYWT' and p.STA_0 = 1 and d.DENCOD_0 = 'PAYIS'
                                                    ORDER BY BPSNAM_0
                                                    ]]></Sql>.Value

        Dim NACHAData As DataTable = OpenDataSet(SQLStatement)

        SectionError = "Testing for records"
        If NACHAData.Rows.Count > 0 Then

            SectionError = "Writing File Header Record"
            intTotalLineCount += 1
            strBuffer += "1"                                                            'Record Type Code; 1 = File Header
            strBuffer += "01"                                                           'Priority Code; Always = 01
            strBuffer += " " + strFedRouting                                            'Immediate Destination
            strBuffer += " " + strBGRRouting                                            'Immediate Origin; SYB Routing
            strBuffer += Now().ToString("yyMMddHHmm")                                   'File Creation Date
            strBuffer += "A"                                                            'File ID Modifier
            strBuffer += "094"                                                          'Record Size; number of bytes per record; Always = 094
            strBuffer += "10"                                                           'Blocking Factor; Block at 10
            strBuffer += "1"                                                            'Format Code; Enter 1
            strBuffer += UCase(strBankName) + StrDup(23 - Len(strBankName), " ")        'Immediate Destination Name; must be 23 characters
            strBuffer += UCase(strCompName) + StrDup(23 - Len(strCompName), " ")        'Immediate Origin Name; must be 23 characters
            strBuffer += Now().ToString("yyMMddHH")                                     'Reference Code; Optional, using date + hour stamp
            strBuffer += vbCrLf

            SectionError = "Writing Batch Header Record"
            intTotalLineCount += 1
            strBuffer += "5"                                                            'Record Type Code; 5 = Batch Header
            strBuffer += "225"                                                          'Service Class Code; 200 ACH mixed Dr/Cr; 220 ACH Cr Only; 225 ACH Dr Only
            strBuffer += UCase(strCompName) + StrDup(16 - Len(strCompName), " ")        'Company Name
            strBuffer += StrDup(20, " ")                                                'Discretionary Data; optional
            strBuffer += "1" + strBGRTaxID                                              'Company Identification = BGR Tax ID
            strBuffer += "PPD"                                                          'Standard Entry Class; PPD(Prearranged Payments and Deposit entries)
            strBuffer += StrDup(10, "0")                                                'Company Entry Description
            strBuffer += StrDup(6, " ")                                                 'Company Descriptive Date
            strBuffer += Now().ToString("yyMMdd")                                       'Effective Entry Date
            strBuffer += StrDup(3, " ")                                                 'Reserved; leave this field blank
            strBuffer += "1"                                                            'Originator Status Code; enter 1 to ID BGR as a depository institution bound by rules of ACH
            strBuffer += Left(strBGRRouting, 8)                                         'Originating Financial Institution - ie. BGR Routing number
            strBuffer += StrDup(7 - Len(CType(intBatchCnt, String)), "0") & intBatchCnt 'Batch Number
            strBuffer += vbCrLf

            For i = 0 To NACHAData.Rows.Count - 1

                If InStr(1, NACHAData.Rows(i).Item("PAYLOT_0"), strPayLots) > 0 Then
                Else
                    strPayLots += NACHAData.Rows(i).Item("PAYLOT_0") + "; "
                End If

                SectionError = "Writing PPD Entry Detail Record #" + i.ToString()
                intTotalLineCount += 1
                intDetailLineCount += 1
                strBuffer += "6"                                                                                            'Record Type Code; 6 = PPD Entry Detail
                strBuffer += "22"                                                                                           'Transaction Code; 22 = Deposit destined for a Checking Account
                strBuffer += Left(NACHAData.Rows(i).Item("RTGNUM_0"), 8)                                                    'Receiving DFI Identification
                lngEntryHash += EntryHash(Left(NACHAData.Rows(i).Item("RTGNUM_0"), 8))
                strBuffer += Right(NACHAData.Rows(i).Item("RTGNUM_0"), 1)                                                   'Check Digit = 9th digit of Receiving institute's routing number
                strBuffer += NACHAData.Rows(i).Item("ACCNUM_0")                                                             'DFI Account Number, left justified and filled to 17 characters
                strBuffer += StrDup(17 - Len(NACHAData.Rows(i).Item("ACCNUM_0")), " ")

                SectionError = "Writing Amount to PPD Detail Record #" + i.ToString()
                dblDebitTotal += NACHAData.Rows(i).Item("AMTCUR_0")
                strBuffer += FormatCurrency(NACHAData.Rows(i).Item("AMTCUR_0"), 8, 2, False)                                'Amount

                SectionError = "Writing Remainder of PPD Detail Record #" + i.ToString()
                strBuffer += StrDup(6, "0")                                                                                 'Recipient Identification Number; optional
                strBuffer += UCase(Left(Trim(NACHAData.Rows(i).Item("BPSNAM_0")), 31))
                If Len(Trim(NACHAData.Rows(i).Item("BPSNAM_0"))) < 31 Then
                    strBuffer += StrDup(31 - Len(Trim(NACHAData.Rows(i).Item("BPSNAM_0"))), "0")                            'Recipient Name
                End If
                strBuffer += StrDup(2, " ")                                                                                 'Discretionary Data; Optional
                strBuffer += "0"                                                                                            'Addenda Record Indicator; 0 = No addenda provided
                strBuffer += Left(strBGRRouting, 8) + StrDup(7 - Len((i + 1).ToString()), "0") + (i + 1).ToString()         'Bank ABA Number; Trace Number for record (7 digit row index)
                strBuffer += vbCrLf
            Next

            SectionError = "Writing Batch Control Record"
            intTotalLineCount += 1
            strBuffer += "8"                                                                                                'Record Type Code; 8 = Batch Control Record
            strBuffer += "225"                                                                                              'Service Class Code; 200 ACH mixed Dr/Cr; 220 ACH Cr Only; 225 ACH Dr Only
            strBuffer += StrDup(6 - Len(intDetailLineCount.ToString()), "0") + intDetailLineCount.ToString()
            If Len(lngEntryHash.ToString()) >= 10 Then
                strBuffer += Right(lngEntryHash.ToString(), 10)                                                             'Entry Hash = total of all positions 4-11 on each 6 record, use final 10 positions
            Else
                strBuffer += StrDup(10 - Len(lngEntryHash.ToString()), "0") + lngEntryHash.ToString()
            End If

            SectionError = "Writing Debit Total to Batch Control"
            strBuffer += FormatCurrency(dblDebitTotal, 10, 2, False)

            SectionError = "Writing Credit Total to Batch Control"
            strBuffer += FormatCurrency(dblCreditTotal, 10, 2, False)

            SectionError = "Writing remainder of Batch Control Record"
            strBuffer += "1" + strBGRTaxID                                                                                  'Company Identification
            strBuffer += Space(19)                                                                                          'Message Authentication Code
            strBuffer += Space(6)                                                                                           'Reserved for federal use
            strBuffer += Left(strBGRRouting, 8)                                                                             'Originating Financial Institution
            strBuffer += StrDup(7 - Len(intBatchCnt.ToString()), "0") + intBatchCnt.ToString()                                    'Batch Number
            strBuffer += vbCrLf

            SectionError = "Writing File Control Record"
            intTotalLineCount += 1
            strBuffer += "9"                                                                                                'Record Type Code; 9 = File Control Record
            strBuffer += StrDup(6 - Len(intBatchCnt.ToString()), "0") + intBatchCnt.ToString()                              'Batch Count
            If intTotalLineCount Mod 10 > 0 Then
                intBlockCount = (intTotalLineCount - (intTotalLineCount Mod 10) + 10) / 10
            Else
                intBlockCount = intTotalLineCount / 10
            End If
            strBuffer += StrDup(6 - Len(intBlockCount.ToString()), "0") + intBlockCount.ToString()                          'Block Count
            strBuffer += StrDup(8 - Len(intDetailLineCount.ToString()), "0") + intDetailLineCount.ToString()                'Entry/Addenda Count
            If Len(lngEntryHash.ToString()) >= 10 Then
                strBuffer += Right(lngEntryHash.ToString(), 10)                                                             'Entry Hash = total of all positions 4-11 on each 6 record, use final 10 positions
            Else
                strBuffer += StrDup(10 - Len(lngEntryHash.ToString()), "0") + lngEntryHash.ToString()
            End If

            SectionError = "Writing Debit Total to File Control"
            strBuffer += FormatCurrency(dblDebitTotal, 10, 2, False)
            SectionError = "Writing Credit Total to File Control"
            strBuffer += FormatCurrency(dblCreditTotal, 10, 2, False)
            SectionError = "Writing remainder of File Control Record"
            strBuffer += StrDup(39, " ") + vbCrLf                                                                           'Reserved; leave blank

            SectionError = "Inserting Dummy Rows to make total lines divisible by 10"
            intDummyLines = 10 - (intTotalLineCount Mod 10)
            If intDummyLines = 10 Then : intDummyLines = 0 : End If
            For i = 1 To intDummyLines
                strBuffer += StrDup(94, "9") + vbCrLf
            Next

            AddText(NACHA, strBuffer)
            NACHA.Close()

            SendHTMLEmail(ToAddress:="cdreyer@packbgr.com", FromAddress:="", Subject:="Your NACHA Export", Body:="Please see attached for your NACHA export file.", Attachments:=path)

            Console.WriteLine(strBuffer)
            Console.WriteLine(intTotalLineCount.ToString() + "|" + intDummyLines.ToString())
            Console.WriteLine("Batches exported: " + strPayLots)

            System.IO.File.Delete(path)

        Else
            Console.WriteLine("No records found that meet criteria: PAM_0 = 'WTR', BAN_0 = 'SYBGR', PAYTYP_0 = 'PAYWT', STA_0 = 'Entered' and DENCOD_0 = 'PAYIS' ")
        End If
    End Sub

    ' This function is used to make a datatable containing the SQL SELECT statement passed to it
    Public Function OpenDataSet(ByRef strSQL As String) As DataTable
        Dim dc As SqlConnection = New SqlConnection("Server=BGRSAGE\X3V6;Database=x3v6;UID=sa;PWD=tiger")
        Dim ds As New DataSet
        Dim cmd As New SqlCommand(strSQL, dc)
        cmd.CommandTimeout = 0
        Dim da As New SqlDataAdapter(cmd)
        da.Fill(ds, "1")
        OpenDataSet = ds.Tables("1")
        dc.Close()
    End Function

    ' This is used to run a SQL INSERT, UPDATE, or DELETE statement
    Public Sub ExecuteSQLQuery(ByVal strSQL As String)
        Dim dc As SqlConnection = New SqlConnection("Server=BGRSAGE\X3V6;Database=x3v6;UID=sa;PWD=tiger")
        Dim SQLcmd As New SqlCommand
        SQLcmd.Connection = dc
        SQLcmd.CommandText = strSQL
        SQLcmd.CommandTimeout = 0
        If Not dc.State = ConnectionState.Open Then dc.Open()
        SQLcmd.ExecuteNonQuery()
        dc.Close()
    End Sub

    Public Function EntryHash(ByVal strInput As String) As Long
        Dim lngBuffer As Long = 0
        For i = 1 To 8
            lngBuffer += CType(Mid(strInput, i, 1), Integer)
        Next
        EntryHash = lngBuffer
    End Function

    Private Sub AddText(ByVal fs As FileStream, ByVal value As String)
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(value)
        fs.Write(info, 0, info.Length)
    End Sub

    Private Function FormatCurrency(ByVal dblInput As Double, ByVal intDollarDigs As Integer, ByVal intCentDigs As Integer, ByVal blnDecimalOutput As Boolean) As String
        Dim strDollars As String = ""
        Dim strCents As String = ""
        Dim strTotal As String = Math.Round(dblInput, intCentDigs).ToString()
        Dim arrTotal() As String = Split(strTotal, ".")

        If arrTotal(0).Length > intDollarDigs Then
            Throw New System.Exception("Overflow FormatCurrency - Dollar Digits")
        Else
            strDollars = StrDup(intDollarDigs - arrTotal(0).Length, "0") + arrTotal(0)
        End If

        If UBound(arrTotal) = 1 Then
            If arrTotal(1).Length > intCentDigs Then
                Throw New System.Exception("Overflow FormatCurrency - Cent Digits")
            Else
                strCents = arrTotal(1) + StrDup(intCentDigs - arrTotal(1).Length, "0")
            End If
        ElseIf Ubound(arrTotal) = 0 Then
            strCents = StrDup(intCentDigs, "0")
        End If

        If blnDecimalOutput Then
            FormatCurrency = strDollars + "." + strCents
        Else
            FormatCurrency = strDollars + strCents
        End If

    End Function


    ' This sends an HTML email
    Public Sub SendHTMLEmail(ByVal ToAddress As String, ByVal FromAddress As String, ByVal Subject As String, ByVal Body As String, Optional ByVal CCAddress As String = "", Optional ByVal BCCAddress As String = "", Optional ByVal Attachments As String = "")
        Dim EmailMessage As New MailMessage
        Dim SimpleSMTP As New SmtpClient("10.200.11.30") '172.16.13.11 192.168.1.7
        SimpleSMTP.UseDefaultCredentials = False
        SimpleSMTP.DeliveryMethod = SmtpDeliveryMethod.Network
        Dim Attch As Mail.Attachment
        EmailMessage.From = New MailAddress(FromAddress)
        ' Pulling out ToAddresses
        Do While InStr(ToAddress, ";") <> 0
            EmailMessage.To.Add(ToAddress.Substring(0, InStr(ToAddress, ";") - 1))
            ToAddress = ToAddress.Substring(InStr(ToAddress, ";"), ToAddress.Length - InStr(ToAddress, ";"))
        Loop
        EmailMessage.To.Add(ToAddress)
        ' Pulling out CC Addresses
        If CCAddress <> "" Then
            Do While InStr(CCAddress, ";") <> 0
                EmailMessage.CC.Add(CCAddress.Substring(0, InStr(CCAddress, ";") - 1))
                CCAddress = CCAddress.Substring(InStr(CCAddress, ";"), CCAddress.Length - InStr(CCAddress, ";"))
            Loop
            EmailMessage.CC.Add(CCAddress)
        End If
        'Pulling out BCC Addresses
        If BCCAddress <> "" Then
            Do While InStr(BCCAddress, ";") <> 0
                EmailMessage.Bcc.Add(BCCAddress.Substring(0, InStr(BCCAddress, ";") - 1))
                BCCAddress = BCCAddress.Substring(InStr(BCCAddress, ";"), BCCAddress.Length - InStr(BCCAddress, ";"))
            Loop
            EmailMessage.Bcc.Add(BCCAddress)
        End If
        'Add Attachments
        If Attachments <> "" Then
            Do While InStr(Attachments, ",") <> 0
                Attch = New Mail.Attachment(Attachments.Substring(0, InStr(Attachments, ",") - 1))
                EmailMessage.Attachments.Add(Attch)
                Attachments = Attachments.Substring(InStr(Attachments, ","), Attachments.Length - InStr(Attachments, ","))
            Loop
            Attch = New Mail.Attachment(Attachments)
            EmailMessage.Attachments.Add(Attch)
        End If
        EmailMessage.Subject = (Subject)
        EmailMessage.Body = (Body)
        EmailMessage.IsBodyHtml = True
        SimpleSMTP.Port = 25
        'SimpleSMTP.EnableSsl = True

        SimpleSMTP.Credentials = New NetworkCredential("scan@packbgr.com", "7100Gano") '"bgr\ituser", "890iu890" bgr\tbailey
        SimpleSMTP.Send(EmailMessage)
    End Sub

End Module