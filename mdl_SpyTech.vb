Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Windows.Forms
'Imports CrystalDecisions.CrystalReports.Engine
'Imports CrystalDecisions.Shared
Imports System.IO
Module mdl_SpyTech
    Public SubscriberVerifiedConnection As SqlConnection
    Public SubscriberVerifiedCommand As SqlCommand
    Public SubscriberVerifiedReader As SqlDataReader
    Public SubscriberCNICCommand As SqlCommand
    Public SubscriberCNICReader As SqlDataReader

    Public OthersConnection As SqlConnection
    Public OthersCommand As SqlCommand
    Public OthersReader As SqlDataReader
    Public ServerName As String = System.Windows.Forms.SystemInformation.ComputerName
    Public IsAddRecord As Boolean = False

    Public DS_CommonNumbers As New DS_CNIC
    Public DT_CommonNumbers As DataTable = DS_CommonNumbers.Tables.Add("tblCommonNumbers")
    Function CreateDTCommonNumbers(ByVal TotalNumbersOfCDRs As Integer, ByVal CDRNames() As String)
        
        DT_CommonNumbers.Columns.Add("CommonNumbers", System.Type.GetType("System.String"))

        For j As Integer = 0 To TotalNumbersOfCDRs - 1
            DT_CommonNumbers.Columns.Add(CDRNames(j), System.Type.GetType("System.String"))
        Next
    End Function
    Function PopulateDVG(ByVal Details As String, ByVal PhoneNumber As String, ByVal Photograph As Byte())
        Dim Row_Index As Integer
        frm_Spy_Tech.DVG_SpyTech.ColumnHeadersVisible = True


        Row_Index = frm_Spy_Tech.DVG_SpyTech.Rows.Count
        'inserting new row 
        frm_Spy_Tech.DVG_SpyTech.Rows.Insert(Row_Index)
        'displaying in phone number column of datagridview
        frm_Spy_Tech.DVG_SpyTech.Rows(Row_Index).Cells("Phone_Number").Value = PhoneNumber.ToString

        frm_Spy_Tech.DVG_SpyTech.Rows(Row_Index).Cells("Detail").Value = Details.ToString
        If Not (IsDBNull(Photograph)) Then
            frm_Spy_Tech.DVG_SpyTech.Rows(Row_Index).Cells("Photograph").Value = Photograph
        End If
        
        frm_Spy_Tech.DVG_SpyTech.Width = frm_Spy_Tech.DVG_SpyTech.Columns(0).Width + frm_Spy_Tech.DVG_SpyTech.Columns(1).Width + frm_Spy_Tech.DVG_SpyTech.Columns(2).Width + 20

        If frm_Spy_Tech.Panel3.Width > frm_Spy_Tech.DVG_SpyTech.Width Then
            frm_Spy_Tech.DVG_SpyTech.Left = frm_Spy_Tech.Panel3.Width / 2 - frm_Spy_Tech.DVG_SpyTech.Width / 2
        Else
            frm_Spy_Tech.DVG_SpyTech.Width = frm_Spy_Tech.Panel3.Width - 10
            frm_Spy_Tech.DVG_SpyTech.Left = 0
        End If
        'frm_Spy_Tech.DVG_SpyTech.Refresh()
        'frm_Spy_Tech.DVG_SpyTech.Columns(0).Width = frm_Spy_Tech.Panel3.Width / 9
        'frm_Spy_Tech.DVG_SpyTech.Columns(1).Width = (frm_Spy_Tech.Panel3.Width / 9) * 3
        'frm_Spy_Tech.DVG_SpyTech.Columns(2).Width = (frm_Spy_Tech.Panel3.Width / 9) * 5
        'frm_Spy_Tech.DVG_SpyTech.Left = 0
        'frm_Spy_Tech.DVG_SpyTech.Top = 0
        'frm_Spy_Tech.DVG_SpyTech.Width = frm_Spy_Tech.DVG_SpyTech.Columns(0).Width + frm_Spy_Tech.DVG_SpyTech.Columns(1).Width + frm_Spy_Tech.DVG_SpyTech.Columns(2).Width

    End Function
    Public Function getTableName(ByVal MobileNo As String) As String
        Dim Mobilecode As String = MobileNo.Substring(MobileNo.Length - 10, 2)
        Dim TableName As String
        If Mobilecode = "31" Then
            'Zong
            TableName = "Masterdb2021"
        ElseIf Mobilecode = "30" Then
            'Mobilink
            TableName = "Masterdb2022"
            'ElseIf Mobilecode = "32" Then
            '    'Warid --Mobilink
            '    TableName = "Masterdb2022"
            'ElseIf Mobilecode = "33" Then
            '    'Ufone
            '    TableName = "Masterdb2022"
            'ElseIf Mobilecode = "34" Then
            '    'Telenor
            '    TableName = "TeleNor2022"
            'ElseIf Mobilecode = "35" Then
            '    'SCOM (AJK & Northern Areas)
            '    TableName = "Masterdb2022"
        End If
        Return TableName
    End Function
    Function FindFromExise(Optional ByVal RegNo As String = "", Optional ByVal CHasisNo As String = "", Optional ByVal EngineNo As String = "")
        Try
            OthersConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;") 
            OthersConnection.Open()
            Dim NameOfTable As String
            If RegNo <> "" Then
                NameOfTable = "KarachiExcise"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarIslamabad"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarLahore"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarMultan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarNankanasahib"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarOkara"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarRahimyarkhan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarRawalpindi"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarSarghoda"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCycleRawalpindi"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorRikshaMultan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE RegistrationNumber like '" & RegNo & "'"
                Call RetrieveOthers("", NameOfTable)


            ElseIf EngineNo <> "" Then
                NameOfTable = "KarachiExcise"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "KarachiExcise"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber2 like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarIslamabad"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarLahore"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarMultan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarNankanasahib"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarOkara"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarRahimyarkhan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarRawalpindi"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarSarghoda"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCycleRawalpindi"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorRikshaMultan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE EngineNumber like '" & EngineNo & "'"
                Call RetrieveOthers("", NameOfTable)
            ElseIf CHasisNo <> "" Then
                If IsNumeric(CHasisNo) Then

                    NameOfTable = "KarachiExcise"
                    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                    Call RetrieveOthers("", NameOfTable)

                    NameOfTable = "KarachiExcise"
                    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber2 like '" & CHasisNo & "'"
                    Call RetrieveOthers("", NameOfTable)
                End If

                NameOfTable = "MotorCarIslamabad"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarLahore"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarMultan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarNankanasahib"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarOkara"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarRahimyarkhan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarRawalpindi"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCarSarghoda"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorCycleRawalpindi"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "MotorRikshaMultan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE ChasisNumber like '" & CHasisNo & "'"
                Call RetrieveOthers("", NameOfTable)
            End If

            OthersConnection.Close()
        Catch ex As Exception
            MsgBox("Error in searching from exise", MsgBoxStyle.OkOnly)
        End Try
    End Function
    Function FindFromEmployee(Optional ByVal BeltNo As String = "", Optional ByVal Name As String = "", Optional ByVal CNIC As String = "")
        Try
            OthersConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;") 'Trusted_Connection=True;")
            OthersConnection.Open()
            Dim NameOfTable As String
            If BeltNo <> "" Then
                NameOfTable = "Employee0"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE BeltNumber like '" & BeltNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "Employee1"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Belt_No like '" & BeltNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "Employee2"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE BeltNo like '" & BeltNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "HCAndCs"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE BeltNumber like '" & BeltNo & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "ISIs"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE BeltNumber like '" & BeltNo & "'"
                Call RetrieveOthers("", NameOfTable)
                NameOfTable = "ASIs"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE BeltNumber like '" & BeltNo & "'"
                Call RetrieveOthers("", NameOfTable)

            ElseIf Name <> "" Then
                NameOfTable = "Employee0"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Name like '" & Name & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "Employee1"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Name like '" & Name & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "Employee2"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Name like '" & Name & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "HCAndCs"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Name like '" & Name & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "Inspectors"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Name like '" & Name & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "ISIs"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Name like '" & Name & "'"
                Call RetrieveOthers("", NameOfTable)

                NameOfTable = "TASI"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Name like '" & Name & "'"
                Call RetrieveOthers("", NameOfTable)
                NameOfTable = "ASIs"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Name like '" & Name & "'"
                Call RetrieveOthers("", NameOfTable)
                
            ElseIf CNIC <> "" Then
                

                NameOfTable = "Employee1"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC like '" & CNIC & "'"
                Call RetrieveOthers("", NameOfTable)

            End If

            OthersConnection.Close()
        Catch ex As Exception
            MsgBox("Error in searching from others" & vbCrLf & Err.Description, MsgBoxStyle.OkOnly)
        End Try
    End Function
    Function FindFromOthers(Optional ByVal CNIC As String = "", Optional ByVal PhoneNumber As String = "", Optional ByVal IsCNIC As Boolean = False, Optional ByVal IsPhoneNumber As Boolean = False)
        Try
            OthersConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;") 'USER ID=sajjad;Password=rajpoot;") 'Trusted_Connection=True;")
            OthersConnection.Open()
            Dim NameOfTable As String
            If (CNIC <> "" And PhoneNumber = "") Or (CNIC <> "" And PhoneNumber <> "" And IsCNIC = True) Then
                NameOfTable = "licence"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "Employee1"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "MotorCarIslamabad"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "MotorCarLahore"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "MotorCarMultan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "MotorCarNankanasahib"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "MotorCarOkara"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "MotorCarRahimyarkhan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "MotorCarRawalpindi"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "MotorCarSarghoda"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "MotorCycleRawalpindi"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "MotorRikshaMultan"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                NameOfTable = "Vehicles"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                OthersConnection.Close()
                'open connection for Rawalpindi_Licence
                OthersConnection = New SqlConnection("Server=" + ServerName + ";Database=Rawalpindi_Licence;Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;") 'Trusted_Connection=True;")
                OthersConnection.Open()

                NameOfTable = "AA"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)
                If CNIC.Length = 13 Then
                    NameOfTable = "FLP"
                    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & Left(CNIC, 5) & "-" & CNIC.Substring(5, 7) & "-" & Right(CNIC, 1) & "'"
                    Call RetrieveOthers(PhoneNumber, NameOfTable)
                End If
                OthersConnection.Close()

                'open connection for Licence
                OthersConnection = New SqlConnection("Server=" + ServerName + ";Database=Licence;Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;") 'Trusted_Connection=True;")
                OthersConnection.Open()

                NameOfTable = "LC_Detail"
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)

            ElseIf (PhoneNumber <> "" And CNIC = "") Or (CNIC <> "" And PhoneNumber <> "" And IsPhoneNumber = True) Then
                NameOfTable = "Employee0"
                'PhoneNumber = "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7)
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7) & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)

                NameOfTable = "Employee1"

                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Right(PhoneNumber, 10) & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)

                NameOfTable = "ASIs"
                'PhoneNumber = "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7)
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7) & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)

                NameOfTable = "Employee2"

                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE MobileNo = '" & Right(PhoneNumber, 10) & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)

                NameOfTable = "ASIs"
                'PhoneNumber = "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7)
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7) & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)

                NameOfTable = "HCAndCs"
                'PhoneNumber = "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7)
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7) & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)

                NameOfTable = "Inspectors"
                'PhoneNumber = "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7)
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7) & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)

                NameOfTable = "ISIs"
                'PhoneNumber = "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7)
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7) & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)

                NameOfTable = "TASI"
                'PhoneNumber = "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7)
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Left(Right(PhoneNumber, 10), 3) & "-" & Right(PhoneNumber, 7) & "'"
                Call RetrieveOthers(PhoneNumber, NameOfTable)

            End If
            Dim recordNo As Integer = DS_PhoneNumber.Tables(1).Rows.Count()
            'Dim T1 As CrystalDecisions.CrystalReports.Engine.TextObject

            'T1 = objRpt.ReportDefinition.Sections(1).ReportObjects("Text1")
            'T1.Text = "Analysis of:  " & PhoneNumber
            'objRpt.SetDataSource(DS_PhoneNumber.Tables(1))
            'frm_CrystalReport.CrystalReportViewer1.ReportSource = objRpt
            ''frm_CrystalReport.CrystalReportViewer1.Refresh()
            'frm_CrystalReport.Show()
            'frm_CrystalReport.CrystalReportViewer1.Refresh()

            ' Call exportAsWordDoc()
        Catch ex As Exception
            MsgBox("Error in FindFromOthers" & vbCrLf & Err.Description, MsgBoxStyle.OkOnly)
        End Try
    End Function
    Function RetrieveOthers(ByVal PhoneNumber As String, ByVal NameOfTable As String)
        Try
            OthersCommand = New SqlCommand(QueryStringPhoneNumber, OthersConnection)
            OthersReader = OthersCommand.ExecuteReader(CommandBehavior.KeyInfo)
            'SchemaTable = SubscriberVerifiedReader.GetSchemaTable()
            Dim RecordCount As Integer = OthersReader.FieldCount()
            Dim PhotoGraph As Byte() = Nothing
            Dim NewImage As Image
            Dim ms As System.IO.MemoryStream
            While OthersReader.Read()
                'Dim tblName As String = SchemaTable.TableName()
                PhotoGraph = Nothing
                Dim NumberOfFields As Integer = OthersReader.FieldCount()
                Dim Details As String = "Table Name:  " & NameOfTable & vbCrLf
                Dim CNIC As String = "CNIC"
                For i As Integer = 0 To NumberOfFields - 1
                    'DT_PhoneNumber.Rows.Add(PhoneNumber, "Mobilink2", PTCL_Reader.GetName(i), PTCL_Reader(i))
                    If OthersReader.IsDBNull(i) Then
                    Else
                        If (OthersReader.GetName(i) = "imgObject") Or (OthersReader.GetName(i) = "imgObj") Or (OthersReader.GetName(i) = "Picture") Then
                            PhotoGraph = OthersReader(i)
                            NewImage = ResizedImage(PhotoGraph, 120, 160)
                            ms = New System.IO.MemoryStream
                            NewImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                            PhotoGraph = ms.ToArray()
                            NewImage.Dispose()
                            ms.Dispose()
                            'ElseIf OthersReader.GetName(i) = "PhoneNumber" Then
                            'PhoneNumber = OthersReader(i)
                        Else
                            Details = Details & Trim(OthersReader.GetName(i)) & ":    " & Trim(OthersReader(i)) & vbCrLf
                            'CNIC = "CNIC:    " & PTCL_Reader(i)
                        End If
                    End If
                Next

                'Row_DT_PhoneNumber(3) = Details
                ' DT_PhoneNumber.Rows.Add(PhoneNumber, NameOfTable, Details, PhotoGraph)
                Call PopulateDVG(Details, PhoneNumber, PhotoGraph)
                Call AddRecordInTable(PhoneNumber, Details, PhotoGraph)


            End While
            OthersReader.Close()
        Catch ex As Exception
            ' MsgBox("Error in RetrieveFromOthers", MsgBoxStyle.OkOnly)
        End Try
    End Function
    Function FindCNICDB2020(ByRef CNIC As String, Optional ByVal PhoneNumber As String = "")
        Dim QueryString As String
        Dim Row_Index As Integer
        Dim Details As String
        Dim phnNo As String
        If PhoneNumber <> "" Then
            phnNo = StandardNumber(PhoneNumber)
        End If

        Try
            SubscriberVerifiedConnection = New SqlConnection("Server=" + ServerName + ";Database=MasterDB2023;Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;") 'Trusted_Connection=True;")
            SubscriberVerifiedConnection.Open()
            Dim NameOfTable1 As String = "Masterdb2023"
            'Dim NameOfTable2 As String = "Masterdb2021"
            'QueryString = "SELECT * FROM PTCLALL WHERE CNIC = '" & CNIC & "'"
            'SubscriberVerifiedCommand = New SqlCommand(QueryString, SubscriberVerifiedConnection)
            'SubscriberVerifiedReader = SubscriberVerifiedCommand.ExecuteReader
            ' 'Searching in Telenor 10 tables from 0 to 9
            'For NumberOfTables As Integer = 1 To 10
            '    NameOfTable = "Mobilink" & NumberOfTables
            'QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable1 + " WHERE CNIC = '" & CNIC & "'"
            QueryStringPhoneNumber = "SELECT  MSISDN,CONCAT('Name: ',NAME,CHAR(10),'CNIC: ',CNIC,CHAR(10),'Address: ',ADDRESS,CHAR(10),'Company: ',COMPANY) AS Detail FROM " + NameOfTable1 + " WHERE CNIC = '" & CNIC & "' "
            '"Union All SELECT  MSISDN,CONCAT('Name: ',NAME,CHAR(10),'CNIC: ',CNIC,CHAR(10),'Address: ',ADDRESS) AS Detail FROM " + NameOfTable2 + " WHERE CNIC = '" & CNIC & "'"
            Call RetrieveSubscriberCNIC(PhoneNumber, NameOfTable1)
            frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            'Next
        Catch ex As Exception
            MsgBox("Error in CNIC Search", MsgBoxStyle.OkOnly)
        End Try
    End Function

    Function FindCNIC(ByRef CNIC As String, Optional ByVal PhoneNumber As String = "")
        Dim QueryString As String
        Dim Row_Index As Integer
        Dim Details As String
        Try
            Dim NameOfTable As String
            SubscriberVerifiedConnection = New SqlConnection("Server=" + ServerName + ";Database=VerifiedSub2018;Trusted_Connection=True;") ' User Id=sajjad;Password=rajpoot;") 'Trusted_Connection=True;")
            SubscriberVerifiedConnection.Open()
            'QueryString = "SELECT * FROM PTCLALL WHERE CNIC = '" & CNIC & "'"
            'SubscriberVerifiedCommand = New SqlCommand(QueryString, SubscriberVerifiedConnection)
            'SubscriberVerifiedReader = SubscriberVerifiedCommand.ExecuteReader
            ' 'Searching in Telenor 10 tables from 0 to 9
            For NumberOfTables As Integer = 1 To 10
                NameOfTable = "Mobilink" & NumberOfTables
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveSubscriberCNIC(PhoneNumber, NameOfTable)
                frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Telenor 6 tables from 0 to 5
            For NumberOfTables As Integer = 1 To 10
                NameOfTable = "Telenor" & NumberOfTables
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveSubscriberCNIC(PhoneNumber, NameOfTable)
                frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Ufone 6 tables from 0 to 5
            For NumberOfTables As Integer = 1 To 4
                NameOfTable = "Ufone" & NumberOfTables
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveSubscriberCNIC(PhoneNumber, NameOfTable)
                frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Warid 6 tables from 0 to 5
            For NumberOfTables As Integer = 1 To 6
                NameOfTable = "Warid1" '& NumberOfTables
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveSubscriberCNIC(PhoneNumber, NameOfTable)
                frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next
            NameOfTable = "WaridPost1" '& NumberOfTables
            QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
            Call RetrieveSubscriberCNIC(PhoneNumber, NameOfTable)
            frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            ' 'Searching in Zong 7 tables from 0 to 6
            For NumberOfTables As Integer = 1 To 6
                NameOfTable = "Zong" & NumberOfTables
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
                Call RetrieveSubscriberCNIC(PhoneNumber, NameOfTable)
                frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Vfone 2 tables from 0 to 1
            ' '' '' '' '' ''For NumberOfTables As Integer = 1 To 2
            ' '' '' '' '' ''    NameOfTable = "Vfone" & NumberOfTables
            ' '' '' '' '' ''    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
            ' '' '' '' '' ''    Call RetrieveSubscriberCNIC(PhoneNumber, NameOfTable)
            ' '' '' '' '' ''    frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            ' '' '' '' '' ''Next

            '' '' '' '' '' '' 'Searching in PTCL 2 tables from 1 to 2
            '' '' '' '' '' ''For NumberOfTables As Integer = 1 To 2
            ' '' '' '' '' ''NameOfTable = "PTCL" '& NumberOfTables
            ' '' '' '' '' ''QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
            ' '' '' '' '' ''Call RetrieveSubscriberCNIC(PhoneNumber, NameOfTable)
            ' '' '' '' '' ''frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            'Next
            'There is no CNIC in following talbes
            '' 'Searching in NTC 2 tables from 1 to 2
            'For NumberOfTables As Integer = 1 To 2
            '    Dim NameOfTable As String = "NTC" & NumberOfTables

            '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE CNIC = '" & CNIC & "'"
            '    Call RetrieveSubscriber("", NameOfTable)
            '    frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            'Next
            SubscriberVerifiedConnection.Close()

        Catch ex As Exception
            MsgBox("Error in CNIC Search", MsgBoxStyle.OkOnly)
        End Try
    End Function

    
    Public LockPhoneNumberToCNIC As String = Nothing
    Public LockCNICToCNIC As String = Nothing
    Public FirstCNICRecord As Boolean = True
    Function SearchCNIC(ByVal PhoneNumber As String) As String
        Dim PhNo As String = Right(RTrim(PhoneNumber), 10)
        Dim CNIC As String = Nothing
        Dim DbName() As String = {"VerifiedSub2018", "OldSubScribersVerified", "Others"}
        Dim TbName() As String = {"Mobilink", "Telenor", "Ufone", "Warid", "Zong"}
        Dim QueryCNIC As String = Nothing
        Dim FieldName As String = Nothing
        Dim isCNICFound As Boolean = False

        For DbNumber As Integer = 0 To 2
            SubscriberVerifiedConnection = New SqlConnection("Server=" + ServerName + ";Database=" + DbName(DbNumber) + ";Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;") 'Trusted_Connection=True;")
            SubscriberVerifiedConnection.Open()
            If DbName(DbNumber) = "VerifiedSub2018" Then
                FieldName = "MobileNo"
            ElseIf DbName(DbNumber) = "OldSubScribersVerified" Or DbName(DbNumber) = "Others" Then
                FieldName = "PhoneNumber"
            End If
            For TbNumber As Integer = 0 To 12
                If DbName(DbNumber) = "VerifiedSub2018" Then
                    QueryCNIC = "Select CNIC from " & TbName(TbNumber + 1) & TbNumber & " Where " & FieldName & "= '" & PhNo & "'"
                ElseIf DbName(DbNumber) = "OldSubScribersVerified" Or DbName(DbNumber) = "Others" Then
                    QueryCNIC = "Select CNIC from " & TbName(TbNumber) & " Where " & FieldName & "= '" & PhNo & "'"
                End If
                SubscriberCNICCommand = New SqlCommand(QueryCNIC, SubscriberVerifiedConnection)
                Try
                    SubscriberCNICReader = SubscriberCNICCommand.ExecuteReader
                Catch ex As Exception
                    Exit For
                End Try

                While SubscriberCNICReader.Read
                    CNIC = SubscriberCNICReader(0).ToString
                    isCNICFound = True
                    Exit For
                End While
            Next
            SubscriberCNICReader.Close()
            SubscriberVerifiedConnection.Close()
            SubscriberCNICCommand.Dispose()
            If isCNICFound = True Then
                Return CNIC
            End If
        Next

        Return CNIC
    End Function
    Function RetrieveSubscriberCNIC(ByVal PhoneNumber As String, ByVal NameOfTable As String)
        Try
            SubscriberCNICCommand = New SqlCommand(QueryStringPhoneNumber, SubscriberVerifiedConnection)
            SubscriberCNICReader = SubscriberCNICCommand.ExecuteReader(CommandBehavior.KeyInfo)
            'SchemaTable = SubscriberVerifiedReader.GetSchemaTable()
            Dim RecordCount As Integer = SubscriberCNICReader.FieldCount()
            Dim PhotoGraph As Byte() = Nothing
            Dim NewImage As Image
            Dim ms As System.IO.MemoryStream

            Dim ListItem As String = Nothing

            While SubscriberCNICReader.Read()
                'Dim tblName As String = SchemaTable.TableName()
                PhotoGraph = Nothing
                Dim NumberOfFields As Integer = SubscriberCNICReader.FieldCount()
                ' Dim Details As String = Nothing
                Dim Details As String = ""   '= "Table Name:  " & NameOfTable & vbCrLf
                Dim CNIC As String = Nothing
                For i As Integer = 0 To NumberOfFields - 1
                    'DT_PhoneNumber.Rows.Add(PhoneNumber, "Mobilink2", PTCL_Reader.GetName(i), PTCL_Reader(i))
                    If SubscriberCNICReader.IsDBNull(i) Then
                    Else
                        If (SubscriberCNICReader.GetName(i) = "imgObject") Or (SubscriberCNICReader.GetName(i) = "imgObj") Or (SubscriberCNICReader.GetName(i) = "Picture") Then
                            PhotoGraph = SubscriberCNICReader(i)
                            NewImage = ResizedImage(PhotoGraph, 120, 160)
                            ms = New System.IO.MemoryStream
                            NewImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                            PhotoGraph = ms.ToArray()
                            NewImage.Dispose()
                            ms.Dispose()
                        ElseIf SubscriberCNICReader.GetName(i) = "Detail" Then
                            'CNIC = SubscriberCNICReader(i)
                            Details = Details & SubscriberCNICReader(i)
                        ElseIf SubscriberCNICReader.GetName(i) = "CNIC" Then
                            CNIC = SubscriberCNICReader(i)
                            Details = Details & SubscriberCNICReader.GetName(i) & ":    " & SubscriberCNICReader(i) & vbCrLf
                        ElseIf SubscriberCNICReader.GetName(i) = "MobileNo" Or SubscriberCNICReader.GetName(i) = "MSISDN" Or SubscriberCNICReader.GetName(i) = "Phone" Or SubscriberCNICReader.GetName(i) = "PhoneNo" Then
                            PhoneNumber = SubscriberCNICReader(i)
                            ListItem = SubscriberCNICReader(i)
                            If ListItem.Substring(0, 2) = "92" Then
                                ListItem = ListItem.Substring(2, ListItem.Length - 2)
                            ElseIf ListItem.Substring(0, 1) = "0" Then
                                ListItem = ListItem.Substring(1, ListItem.Length - 1)
                            End If
                            If frm_Spy_Tech.Lstbx_CNIC_PhoneNumbers.Items.Contains(ListItem) Then
                                Continue While
                            End If
                            If Not frm_Spy_Tech.Lstbx_CNIC_PhoneNumbers.Items.Contains(ListItem) Then
                                frm_Spy_Tech.Lstbx_CNIC_PhoneNumbers.Items.Add(ListItem)
                            End If

                        Else
                            Details = Details & Trim(SubscriberCNICReader.GetName(i)) & ":    " & Trim(SubscriberCNICReader(i)) & vbCrLf

                        End If
                    End If
                Next

                'Row_DT_PhoneNumber(3) = Details
                If (Right(LockPhoneNumber, 10) <> Right(PhoneNumber, 10)) Or (LockCNIC <> CNIC) Or (IsFirstRecord = True) Then
                    'Row_DT_PhoneNumber(3) = Details
                    ' DT_PhoneNumber.Rows.Add(PhoneNumber, NameOfTable, Details, PhotoGraph)
                    Call PopulateDVG(Details, PhoneNumber, PhotoGraph)
                    frm_Spy_Tech.DVG_SpyTech.Refresh()
                    Call AddRecordInTable(PhoneNumber, Details, PhotoGraph)
                End If
                If LockCNIC <> CNIC Then
                    Call FindFromOthers(CNIC)
                End If
                LockCNIC = CNIC
                LockPhoneNumber = PhoneNumber
                IsFirstRecord = False
                'DT_PhoneNumber.Rows.Add(PhoneNumber, NameOfTable, Details, PhotoGraph)
                'Call PopulateDVG(Details, PhoneNumber, PhotoGraph)
                'Call FindFromOthers(CNIC)
            End While
            SubscriberCNICReader.Close()
        Catch ex As Exception
            'MsgBox("Error in Retrieve Subscriber CNIC", MsgBoxStyle.OkOnly)
        End Try
    End Function
    Public DS_PhoneNumber As New DS_CNIC
    'Public DT_Column As New DataColumn
    Public DT_PhoneNumber As DataTable = DS_PhoneNumber.Tables.Add("tblPhoneNumber")
    Public QueryStringPhoneNumber As String
    Function CreateSearchStore()

        DT_PhoneNumber.Columns.Add("PhoneNumber", System.Type.GetType("System.String"))
        DT_PhoneNumber.Columns.Add("TableName", System.Type.GetType("System.String"))
        DT_PhoneNumber.Columns.Add("Detail", System.Type.GetType("System.String"))
        Dim Photo As New DataColumn()
        Photo = New DataColumn("Photograph", GetType(System.Byte()))
        DT_PhoneNumber.Columns.Add(Photo)
        'DT_PhoneNumber.Columns.Add("Photo", System.Type.GetType("System.Drawing.Bitmap"))

    End Function
    Public LockCNIC As String = Nothing
    Public LockPhoneNumber As String = Nothing
    Dim IsFirstRecord As Boolean = True
    Public IsCDR As Boolean = False
    Public NumbersOfCalls As String = Nothing
    Public NumbersNotFound As String
    Public IsNumberFoundCDR As Boolean = False
    Public GprCNIC As String = Nothing
    Public IsCNIC_Need As Boolean = False
    Public Network As String
    Public SubName As String
    Public DateOfActivation As String
    Public isFoundVerified As Boolean = False

    Function RetrieveMasterDB2020(ByVal PhoneNumber As String, ByVal NameOfTable As String)
        Try

            SubscriberVerifiedCommand = New SqlCommand(QueryStringPhoneNumber, SubscriberVerifiedConnection)
            SubscriberVerifiedReader = SubscriberVerifiedCommand.ExecuteReader(CommandBehavior.KeyInfo)
            'SchemaTable = SubscriberVerifiedReader.GetSchemaTable()
            Dim RecordCount As Integer = SubscriberVerifiedReader.FieldCount()
            Dim PhotoGraph As Byte() = Nothing
            Dim NewImage As Image
            Dim ms As System.IO.MemoryStream

            'IsNumberFoundCDR = False
            While SubscriberVerifiedReader.Read()
                IsNumberFoundCDR = True
                'Dim tblName As String = SchemaTable.TableName()
                PhotoGraph = Nothing
                Dim NumberOfFields As Integer = SubscriberVerifiedReader.FieldCount()
                Dim Details As String ''= "Table Name:  " & NameOfTable & vbCrLf
                Dim CNIC As String = Nothing
                For i As Integer = 0 To NumberOfFields - 1
                    'DT_PhoneNumber.Rows.Add(PhoneNumber, "Mobilink2", PTCL_Reader.GetName(i), PTCL_Reader(i))
                    If SubscriberVerifiedReader.IsDBNull(i) Then
                    Else

                        If (SubscriberVerifiedReader.GetName(i) = "imgObject") Or (SubscriberVerifiedReader.GetName(i) = "imgObj") Or (SubscriberVerifiedReader.GetName(i) = "Picture") Then
                            PhotoGraph = SubscriberVerifiedReader(i)
                            NewImage = ResizedImage(PhotoGraph, 120, 160)
                            ms = New System.IO.MemoryStream
                            NewImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                            PhotoGraph = ms.ToArray()
                            NewImage.Dispose()
                            ms.Dispose()
                        ElseIf UCase(SubscriberVerifiedReader.GetName(i)) = "STATUS" Or UCase(SubscriberVerifiedReader.GetName(i)) = "BVS" Then
                            If UCase(Trim(SubscriberVerifiedReader(i).ToString).Substring(0, 1)) = "U" Then
                                GoTo UnverifiedNum
                            Else
                                Details = Details & Trim(SubscriberVerifiedReader.GetName(i)) & ":    " & Trim(SubscriberVerifiedReader(i)) & vbCrLf
                            End If
                        ElseIf SubscriberVerifiedReader.GetName(i) = "CNIC" Then
                            CNIC = SubscriberVerifiedReader(i)
                            GprCNIC = CNIC
                            Network = NameOfTable
                            Details = Details & Trim(SubscriberVerifiedReader.GetName(i)) & ":    " & Trim(SubscriberVerifiedReader(i)) & vbCrLf
                        ElseIf SubscriberVerifiedReader.GetName(i) = "NAME" Then
                            SubName = SubscriberVerifiedReader(i)
                            Details = Details & Trim(SubscriberVerifiedReader.GetName(i)) & ":    " & Trim(SubscriberVerifiedReader(i)) & vbCrLf
                        ElseIf SubscriberVerifiedReader.GetName(i) = "MSISDN" Then
                            'PhoneNumber = SubscriberVerifiedReader(i
                        ElseIf UCase(SubscriberVerifiedReader.GetName(i)) = "ADDRESS" Then
                            DateOfActivation = SubscriberVerifiedReader(i)
                            Details = Details & Trim(SubscriberVerifiedReader.GetName(i)) & ":    " & Trim(SubscriberVerifiedReader(i)) & vbCrLf
                        Else

                            Details = Details & Trim(SubscriberVerifiedReader.GetName(i)) & ":    " & Trim(SubscriberVerifiedReader(i)) & vbCrLf

                        End If
                    End If

                Next
                If IsCNIC_Need = False Then
                    If (Right(LockPhoneNumber, 10) <> Right(PhoneNumber, 10)) Or (LockCNIC <> CNIC) Or (IsFirstRecord = True) Then
                        'Row_DT_PhoneNumber(3) = Details
                        If IsCDR = True Then
                            Dim CDR_PhoneNumberAndCount As String = PhoneNumber & vbCrLf & "(" & NumbersOfCalls & ")"
                            'DT_PhoneNumber.Rows.Add(CDR_PhoneNumberAndCount, NameOfTable, Details, PhotoGraph)
                            Call PopulateDVG(Details, CDR_PhoneNumberAndCount, PhotoGraph)
                            Call AddRecordInTable(CDR_PhoneNumberAndCount, Details, PhotoGraph)
                        Else
                            ' DT_PhoneNumber.Rows.Add(PhoneNumber, NameOfTable, Details, PhotoGraph)
                            Call PopulateDVG(Details, PhoneNumber, PhotoGraph)
                            Call AddRecordInTable(PhoneNumber, Details, PhotoGraph)
                        End If
                    End If
                    If LockCNIC <> CNIC Then
                        'following function is disabled for rasheed
                        Call FindFromOthers(CNIC)
                    End If
                End If
                LockCNIC = CNIC
                LockPhoneNumber = PhoneNumber
                IsFirstRecord = False
                isFoundVerified = True
UnverifiedNum:

            End While

            SubscriberVerifiedReader.Close()
        Catch ex As Exception
            ' MsgBox("Error in Retrieve From Subscribers", MsgBoxStyle.OkOnly)
        End Try
    End Function

    Function RetrieveSubscriber(ByVal PhoneNumber As String, ByVal NameOfTable As String)
        Try

            SubscriberVerifiedCommand = New SqlCommand(QueryStringPhoneNumber, SubscriberVerifiedConnection)
            SubscriberVerifiedReader = SubscriberVerifiedCommand.ExecuteReader(CommandBehavior.KeyInfo)
            'SchemaTable = SubscriberVerifiedReader.GetSchemaTable()
            Dim RecordCount As Integer = SubscriberVerifiedReader.FieldCount()
            Dim PhotoGraph As Byte() = Nothing
            Dim NewImage As Image
            Dim ms As System.IO.MemoryStream

            'IsNumberFoundCDR = False
            While SubscriberVerifiedReader.Read()
                IsNumberFoundCDR = True
                'Dim tblName As String = SchemaTable.TableName()
                PhotoGraph = Nothing
                Dim NumberOfFields As Integer = SubscriberVerifiedReader.FieldCount()
                Dim Details As String = "Table Name:  " & NameOfTable & vbCrLf
                Dim CNIC As String = Nothing
                For i As Integer = 0 To NumberOfFields - 1
                    'DT_PhoneNumber.Rows.Add(PhoneNumber, "Mobilink2", PTCL_Reader.GetName(i), PTCL_Reader(i))
                    If SubscriberVerifiedReader.IsDBNull(i) Then
                    Else

                        If (SubscriberVerifiedReader.GetName(i) = "imgObject") Or (SubscriberVerifiedReader.GetName(i) = "imgObj") Or (SubscriberVerifiedReader.GetName(i) = "Picture") Then
                            PhotoGraph = SubscriberVerifiedReader(i)
                            NewImage = ResizedImage(PhotoGraph, 120, 160)
                            ms = New System.IO.MemoryStream
                            NewImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                            PhotoGraph = ms.ToArray()
                            NewImage.Dispose()
                            ms.Dispose()
                        ElseIf UCase(SubscriberVerifiedReader.GetName(i)) = "STATUS" Or UCase(SubscriberVerifiedReader.GetName(i)) = "BVS" Then
                            If UCase(Trim(SubscriberVerifiedReader(i).ToString).Substring(0, 1)) = "U" Then
                                GoTo UnverifiedNum
                            Else
                                Details = Details & Trim(SubscriberVerifiedReader.GetName(i)) & ":    " & Trim(SubscriberVerifiedReader(i)) & vbCrLf
                            End If
                        ElseIf SubscriberVerifiedReader.GetName(i) = "CNIC" Then
                            CNIC = SubscriberVerifiedReader(i)
                            GprCNIC = CNIC
                            Network = NameOfTable
                            Details = Details & Trim(SubscriberVerifiedReader.GetName(i)) & ":    " & Trim(SubscriberVerifiedReader(i)) & vbCrLf
                        ElseIf SubscriberVerifiedReader.GetName(i) = "Name" Then
                            SubName = SubscriberVerifiedReader(i)
                            Details = Details & Trim(SubscriberVerifiedReader.GetName(i)) & ":    " & Trim(SubscriberVerifiedReader(i)) & vbCrLf
                        ElseIf SubscriberVerifiedReader.GetName(i) = "PhoneNumber" Then
                            'PhoneNumber = SubscriberVerifiedReader(i
                        ElseIf UCase(SubscriberVerifiedReader.GetName(i)) = "ADDRESS" Then
                            DateOfActivation = SubscriberVerifiedReader(i)
                            Details = Details & Trim(SubscriberVerifiedReader.GetName(i)) & ":    " & Trim(SubscriberVerifiedReader(i)) & vbCrLf
                        Else

                            Details = Details & Trim(SubscriberVerifiedReader.GetName(i)) & ":    " & Trim(SubscriberVerifiedReader(i)) & vbCrLf

                        End If
                    End If

                Next
                If IsCNIC_Need = False Then
                    If (Right(LockPhoneNumber, 10) <> Right(PhoneNumber, 10)) Or (LockCNIC <> CNIC) Or (IsFirstRecord = True) Then
                        'Row_DT_PhoneNumber(3) = Details
                        If IsCDR = True Then
                            Dim CDR_PhoneNumberAndCount As String = PhoneNumber & vbCrLf & "(" & NumbersOfCalls & ")"
                            'DT_PhoneNumber.Rows.Add(CDR_PhoneNumberAndCount, NameOfTable, Details, PhotoGraph)
                            Call PopulateDVG(Details, CDR_PhoneNumberAndCount, PhotoGraph)
                            Call AddRecordInTable(CDR_PhoneNumberAndCount, Details, PhotoGraph)
                        Else
                            ' DT_PhoneNumber.Rows.Add(PhoneNumber, NameOfTable, Details, PhotoGraph)
                            Call PopulateDVG(Details, PhoneNumber, PhotoGraph)
                            Call AddRecordInTable(PhoneNumber, Details, PhotoGraph)
                        End If
                    End If
                    If LockCNIC <> CNIC Then
                        'following function is disabled for rasheed
                        Call FindFromOthers(CNIC)
                    End If
                End If
                LockCNIC = CNIC
                LockPhoneNumber = PhoneNumber
                IsFirstRecord = False
                isFoundVerified = True
UnverifiedNum:

            End While

            SubscriberVerifiedReader.Close()
        Catch ex As Exception
            ' MsgBox("Error in Retrieve From Subscribers", MsgBoxStyle.OkOnly)
        End Try
    End Function
    Sub FindPhoneNumberOldDB(ByVal PhoneNumber As String)
        Dim QueryString As String
        Dim SchemaTable As DataTable
        Dim PhNo As String = Right(PhoneNumber, 10)
        Try
            SubscriberVerifiedConnection = New SqlConnection("Server=" + ServerName + ";Database=OldSubScribersVerified;Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;")   'Trusted_Connection=True;")
            SubscriberVerifiedConnection.Open()

            'Searching in Mobilink 10 tables from 0 to 9
            ' ''For NumberOfTables As Integer = 0 To 9
            ' ''    Dim NameOfTable As String = "Mobilink" & NumberOfTables
            ' ''    If NumberOfTables = 0 Then
            ' ''        QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "92" & Right(PhoneNumber, 10) & "'"
            ' ''        Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            ' ''    ElseIf NumberOfTables = 1 Then
            ' ''        QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
            ' ''        Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            ' ''    Else
            ' ''        QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
            ' ''        Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            ' ''    End If
            ' ''    'QueryStringPhoneNumber = "SELECT * FROM Mobilink2 WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
            ' ''    'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            ' ''Next

            ' 'Searching in Telenor 6 tables from 0 to 5
            For NumberOfTables As Integer = 0 To 5
                Dim NameOfTable As String = "Telenor" & NumberOfTables
                'If NumberOfTables = 0 Then
                '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & PhNo & "'"
                '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'ElseIf NumberOfTables = 1 Then
                '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Right(PhoneNumber, 10) & "'"
                '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'Else
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & PhNo & "'"
                Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                ' End If
                'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Ufone 6 tables from 0 to 5
            For NumberOfTables As Integer = 0 To 5
                Dim NameOfTable As String = "Ufone" & NumberOfTables
                'If NumberOfTables = 0 Then
                '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "92" & Right(PhoneNumber, 10) & "'"
                '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)

                'Else
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & PhNo & "'"
                Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'End If
                ' frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Warid 6 tables from 0 to 5
            For NumberOfTables As Integer = 0 To 5
                Dim NameOfTable As String = "Warid" & NumberOfTables
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & PhNo & "'"
                Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                ' frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Zong 7 tables from 0 to 6
            For NumberOfTables As Integer = 0 To 5
                Dim NameOfTable As String = "Zong" & NumberOfTables
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & PhNo & "'"
                Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Vfone 2 tables from 0 to 1
            For NumberOfTables As Integer = 0 To 1
                Dim NameOfTable As String = "Vfone" & NumberOfTables
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & PhNo & "'"
                Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in PTCL 2 tables from 1 to 2
            For NumberOfTables As Integer = 1 To 2
                Dim NameOfTable As String = "PTCL" & NumberOfTables
                If NumberOfTables = 1 Then
                    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Right(PhoneNumber, 9) & "'"
                    Call RetrieveSubscriber(PhoneNumber, NameOfTable)

                Else
                    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & Right(PhoneNumber, 9) & "'"
                    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                End If
                ' frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in NTC 2 tables from 1 to 2
            For NumberOfTables As Integer = 1 To 2
                Dim NameOfTable As String = "NTC" & NumberOfTables

                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & Right(PhoneNumber, 9) & "'"
                Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            'Dim recordNo As Integer = DS_PhoneNumber.Tables(1).Rows.Count()
            'Dim T1 As CrystalDecisions.CrystalReports.Engine.TextObject

            'T1 = objRpt.ReportDefinition.Sections(1).ReportObjects("Text1")
            'T1.Text = "Analysis of:  " & PhoneNumber
            'objRpt.SetDataSource(DS_PhoneNumber.Tables(1))
            'frm_CrystalReport.CrystalReportViewer1.ReportSource = objRpt
            ''frm_CrystalReport.CrystalReportViewer1.Refresh()
            ''frm_CrystalReport.Show()
            'frm_CrystalReport.CrystalReportViewer1.Refresh()

            'Call exportAsWordDoc()
            SubscriberVerifiedConnection.Close()
        Catch ex As Exception
            MsgBox("Error in Find Phone Number", MsgBoxStyle.OkOnly)
        End Try
    End Sub
    Function StandardNumber(ByVal PhoneNumber As String) As String
        If PhoneNumber.StartsWith("92") Then
            Return PhoneNumber
        ElseIf PhoneNumber.Length = 10 Then
            PhoneNumber = "92" + PhoneNumber
        ElseIf (PhoneNumber.Length = 11) And (PhoneNumber.StartsWith("0") = False) Then
            PhoneNumber = "92" + PhoneNumber.Substring(1)
        End If
        Return PhoneNumber
    End Function

    Function FindNumberDB2021(ByVal PhoneNumber As String)
        Dim QueryString As String
        Dim SchemaTable As DataTable
        isFoundVerified = False
        Dim PhNo As String = StandardNumber(PhoneNumber)
        Try
            SubscriberVerifiedConnection = New SqlConnection("Server=" + ServerName + ";Database=MasterDB2023;Trusted_Connection=True;") '("Server=" + ServerName + ";Database=MasterDB2022;User Id=sajjad;Password=rajpoot;")

            SubscriberVerifiedConnection.Open()
            Dim NameOfTable As String = "Masterdb2023"
            ''Searching in Mobilink 10 tables from 0 to 9
            'For NumberOfTables As Integer = 1 To 10
            '    NameOfTable = "Mobilink" & NumberOfTables
            '    'If NumberOfTables = 0 Then
            '    '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "92" & Right(PhoneNumber, 10) & "'"
            '    '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            '    'ElseIf NumberOfTables = 1 Then
            '    '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
            '    '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            '    'Else
            QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE MSISDN = '" & PhNo & "'"
            Call RetrieveMasterDB2020(PhNo, NameOfTable)
            ' End If
            'QueryStringPhoneNumber = "SELECT * FROM Mobilink2 WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
            'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            'Next
            SubscriberVerifiedConnection.Close()
        Catch ex As Exception
            MsgBox("Error in Find Phone Number", MsgBoxStyle.OkOnly)
        End Try

    End Function
    Function FindNumberDB2022(ByVal PhoneNumber As String)
        Dim QueryString As String
        Dim SchemaTable As DataTable
        isFoundVerified = False
        Dim PhNo As String = StandardNumber(PhoneNumber)
        Try
            SubscriberVerifiedConnection = New SqlConnection("Server=" + ServerName + ";Database=MasterDB2023;Trusted_Connection=True;") '("Server=" + ServerName + ";Database=MasterDB2022;User Id=sajjad;Password=rajpoot;")

            SubscriberVerifiedConnection.Open()
            'Dim NameOfTable As String = getTableName(PhNo)
            Dim NameOfTable As String = "Masterdb2023"
            ''Searching in Mobilink 10 tables from 0 to 9
            'For NumberOfTables As Integer = 1 To 10
            '    NameOfTable = "Mobilink" & NumberOfTables
            '    'If NumberOfTables = 0 Then
            '    '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "92" & Right(PhoneNumber, 10) & "'"
            '    '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            '    'ElseIf NumberOfTables = 1 Then
            '    '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
            '    '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            '    'Else
            QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE MSISDN = '" & PhNo & "'"
            Call RetrieveMasterDB2020(PhNo, NameOfTable)
            ' End If
            'QueryStringPhoneNumber = "SELECT * FROM Mobilink2 WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
            'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            'Next
            SubscriberVerifiedConnection.Close()
        Catch ex As Exception
            MsgBox("Error in Find Phone Number", MsgBoxStyle.OkOnly)
        End Try

    End Function
    Function FindNumberDB2221(ByVal PhoneNumber As String)
        Dim QueryString As String
        Dim SchemaTable As DataTable
        isFoundVerified = False
        Dim PhNo As String = StandardNumber(PhoneNumber)
        Try
            SubscriberVerifiedConnection = New SqlConnection("Server=" + ServerName + ";Database=MasterDB2023;Trusted_Connection=True;") '("Server=" + ServerName + ";Database=MasterDB2022;User Id=sajjad;Password=rajpoot;")

            SubscriberVerifiedConnection.Open()
            'Dim NameOfTable As String = getTableName(PhNo)
            Dim NameOfTable As String = "Masterdb2023"
            ''Searching in Mobilink 10 tables from 0 to 9
            'For NumberOfTables As Integer = 1 To 10
            '    NameOfTable = "Mobilink" & NumberOfTables
            '    'If NumberOfTables = 0 Then
            '    '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "92" & Right(PhoneNumber, 10) & "'"
            '    '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            '    'ElseIf NumberOfTables = 1 Then
            '    '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
            '    '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            '    'Else
            QueryStringPhoneNumber = "SELECT Masterdb2022.MSISDN,Masterdb2022.[CNIC],Masterdb2021.[NAME],TRIM(Masterdb2021.[ADDRESS]) AS ADDRESS,Masterdb2022.[COMPANY]  FROM " + NameOfTable + " left join Masterdb2021 on Masterdb2022.CNIC= Masterdb2021.CNIC WHERE Masterdb2022.MSISDN = '" & PhNo & "'"
            Call RetrieveMasterDB2020(PhNo, NameOfTable)
            ' End If
            'QueryStringPhoneNumber = "SELECT * FROM Mobilink2 WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
            'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            'Next
            SubscriberVerifiedConnection.Close()
        Catch ex As Exception
            MsgBox("Error in Find Phone Number", MsgBoxStyle.OkOnly)
        End Try

    End Function
    Function FindPhoneNumber(ByVal PhoneNumber As String)
        Dim QueryString As String
        Dim SchemaTable As DataTable
        isFoundVerified = False
        Dim PhNo As String = Right(PhoneNumber, 10)
        Try
            SubscriberVerifiedConnection = New SqlConnection("Server=" + ServerName + ";Database=VerifiedSub2018;Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;")   'Trusted_Connection=True;")
            SubscriberVerifiedConnection.Open()
            Dim NameOfTable As String
            'Searching in Mobilink 10 tables from 0 to 9
            For NumberOfTables As Integer = 1 To 10
                NameOfTable = "Mobilink" & NumberOfTables
                'If NumberOfTables = 0 Then
                '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "92" & Right(PhoneNumber, 10) & "'"
                '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'ElseIf NumberOfTables = 1 Then
                '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
                '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'Else
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE MobileNo = '" & PhNo & "'"
                Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                ' End If
                'QueryStringPhoneNumber = "SELECT * FROM Mobilink2 WHERE PhoneNumber = '" & Right(PhoneNumber, 10) & "'"
                'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Telenor 6 tables from 0 to 5
            For NumberOfTables As Integer = 1 To 10
                NameOfTable = "Telenor" & NumberOfTables
                'If NumberOfTables = 0 Then
                '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Right(PhoneNumber, 10) & "'"
                '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'ElseIf NumberOfTables = 1 Then
                '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "0" & Right(PhoneNumber, 10) & "'"
                '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'Else
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE MobileNo = '" & PhNo & "'"
                Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'End If
                'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Ufone 6 tables from 0 to 5
            For NumberOfTables As Integer = 1 To 4
                NameOfTable = "Ufone" & NumberOfTables
                'If NumberOfTables = 0 Then
                '    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNumber = '" & "92" & Right(PhoneNumber, 10) & "'"
                '    Call RetrieveSubscriber(PhoneNumber, NameOfTable)

                'Else
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE MobileNo = '" & PhNo & "'"
                Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'End If
                ' frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Warid 6 tables from 0 to 5
            For NumberOfTables As Integer = 1 To 6
                NameOfTable = "Warid" & NumberOfTables
                If NumberOfTables = 1 Then
                    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE MobileNo = '" & PhNo & "'"
                    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                Else
                    NameOfTable = "WaridPost" & NumberOfTables
                    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE MobileNo = '" & PhNo & "'"
                    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                End If
                ' frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Zong 7 tables from 0 to 6
            For NumberOfTables As Integer = 1 To 6
                NameOfTable = "Zong" & NumberOfTables
                QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE MobileNo = '" & PhNo & "'"
                Call RetrieveSubscriber(PhoneNumber, NameOfTable)
                'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            Next

            ' 'Searching in Vfone 2 tables from 0 to 1
            ' '' '' ''For NumberOfTables As Integer = 1 To 2
            ' '' '' ''    NameOfTable = "Vfone" & NumberOfTables
            ' '' '' ''    QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE PhoneNo = '" & PhNo & "'"
            ' '' '' ''    Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            ' '' '' ''    'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            ' '' '' ''Next

            ' 'Searching in PTCL 2 tables from 1 to 2
            'For NumberOfTables As Integer = 1 To 2
            '''''''''''''''  NameOfTable = "PTCL" '& NumberOfTables
            ''If NumberOfTables = 1 Then
            'QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Phone = '" & "0" & Right(PhoneNumber, 9) & "'"
            'Call RetrieveSubscriber(PhoneNumber, NameOfTable)

            'Else
            ' '' '' '' '' ''QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE Phone = '" & Right(PhoneNumber, 9) & "'"
            ' '' '' '' '' ''Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            'End If
            ' frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            ' Next

            ' 'Searching in NTC 2 tables from 1 to 2
            'For NumberOfTables As Integer = 1 To 1
            'Dim NameOfTable As String = "NTC" & NumberOfTables
            ' '' '' '' ''NameOfTable = "NTC"
            ' '' '' '' ''QueryStringPhoneNumber = "SELECT * FROM " + NameOfTable + " WHERE MSISDN = '" & Right(PhoneNumber, 9) & "'"
            ' '' '' '' ''Call RetrieveSubscriber(PhoneNumber, NameOfTable)
            'frm_Spy_Tech.prgbarSearch.Value = frm_Spy_Tech.prgbarSearch.Value + 1
            ' Next

            'Dim recordNo As Integer = DS_PhoneNumber.Tables(1).Rows.Count()
            'Dim T1 As CrystalDecisions.CrystalReports.Engine.TextObject

            'T1 = objRpt.ReportDefinition.Sections(1).ReportObjects("Text1")
            'T1.Text = "Analysis of:  " & PhoneNumber
            'objRpt.SetDataSource(DS_PhoneNumber.Tables(1))
            'frm_CrystalReport.CrystalReportViewer1.ReportSource = objRpt
            ''frm_CrystalReport.CrystalReportViewer1.Refresh()
            ''frm_CrystalReport.Show()
            'frm_CrystalReport.CrystalReportViewer1.Refresh()

            'Call exportAsWordDoc()
            SubscriberVerifiedConnection.Close()
        Catch ex As Exception
            MsgBox("Error in Find Phone Number", MsgBoxStyle.OkOnly)
        End Try
        If isFoundVerified = False Then
            Call FindPhoneNumberOldDB(PhoneNumber)
        End If

    End Function
    Dim InsertQuery As String
    Dim TargetConnection As SqlConnection
    Sub InsertCNIC_CropBTS(ByVal PhnNumber As String, ByVal CNIC As String, ByVal TableName As String)
        TargetConnection = New SqlConnection("Server=" + ServerName + ";Database=Others;Trusted_Connection=True;") 'User Id=sajjad;Password=rajpoot;")   'Trusted_Connection=True;")
        If TargetConnection.State = ConnectionState.Closed Then
            TargetConnection.Open()
        End If
        Dim InsertIMEICommand As SqlCommand
        InsertQuery = "INSERT INTO [" & TableName & "] VALUES ('" & PhnNumber & "','" & CNIC & "')"
        '" & CNIC & "'"
        InsertIMEICommand = New SqlCommand(InsertQuery, TargetConnection)
        Try
            InsertIMEICommand.ExecuteNonQuery()
        Catch ex As Exception

        End Try

        TargetConnection.Close()
        TargetConnection.Dispose()
        InsertIMEICommand.Dispose()
    End Sub
    Function exportAsWordDoc(ByVal DocumentTitle As String, ByVal PatheAndFileName As String, Optional ByVal IsCDR As Boolean = False)
        'Dim CrExportOptions As ExportOptions
        'Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
        'Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()
        'Try
        '    Dim T1 As CrystalDecisions.CrystalReports.Engine.TextObject

        '    'T1 = objRpt.ReportDefinition.Sections(1).ReportObjects("Text1")
        '    If IsCDR = True Then
        '        T1.Text = "Analysis of:  " & DocumentTitle
        '        CrDiskFileDestinationOptions.DiskFileName = PatheAndFileName & ".doc"
        '    ElseIf IsCDR = False Then
        '        T1.Text = "Analysis of:  " & DocumentTitle
        '        CrDiskFileDestinationOptions.DiskFileName = PatheAndFileName '"E:\Example\" & frm_Spy_Tech.txt_Phone_Number.Text & ".doc"


        '    End If

        '    'objRpt.SetDataSource(DS_PhoneNumber.Tables(1))
        '    'frm_CrystalReport.CrystalReportViewer1.ReportSource = objRpt
        '    'frm_CrystalReport.CrystalReportViewer1.Refresh()
        '    'frm_CrystalReport.Show()
        '    frm_CrystalReport.CrystalReportViewer1.Refresh()
        '    If frm_Spy_Tech.txt_Phone_Number.Text = "" Then
        '        'CrDiskFileDestinationOptions.DiskFileName = "E:\Example\" & "My" & ".doc"
        '    Else
        '        'CrDiskFileDestinationOptions.DiskFileName = "E:\Example\" & frm_Spy_Tech.txt_Phone_Number.Text & ".doc"
        '    End If
        '    ' CrExportOptions = objRpt.ExportOptions
        '    With CrExportOptions
        '        .ExportDestinationType = ExportDestinationType.DiskFile
        '        .ExportFormatType = ExportFormatType.WordForWindows
        '        .DestinationOptions = CrDiskFileDestinationOptions
        '        .FormatOptions = CrFormatTypeOptions
        '    End With
        '    ' objRpt.Export()
        '    MsgBox("Document has been Created:" & vbCrLf & PatheAndFileName, MsgBoxStyle.OkOnly)
        'Catch ex As Exception
        '    MsgBox("Error occurs during creating word Doc", MsgBoxStyle.OkOnly)
        'End Try
    End Function
    Function PTCL_Search(ByRef PTCL_Number As String)
        Dim QueryString As String
        Dim Row_Index As Integer
        Dim PTCL_Details As String

        'Try
        '    PTCL_Connection = New SqlConnection("Server=" + ServerName + ";Database=subscriberverified;Trusted_Connection=True;")

        '    PTCL_Connection.Open()
        '    QueryString = "SELECT * FROM PTCL01 WHERE PhoneNumber = '" & PTCL_Number & "'"
        '    PTCL_Command = New SqlCommand(QueryString, PTCL_Connection)
        '    PTCL_Reader = PTCL_Command.ExecuteReader

        '    If frm_Spy_Tech.chk_Add_Result.Checked = False Then
        '        frm_Spy_Tech.DVG_SpyTech.Rows.Clear()
        '    End If
        '    While PTCL_Reader.Read()
        '        frm_Spy_Tech.DVG_SpyTech.ColumnHeadersVisible = True


        '        Row_Index = frm_Spy_Tech.DVG_SpyTech.Rows.Count
        '        'inserting new row 
        '        frm_Spy_Tech.DVG_SpyTech.Rows.Insert(Row_Index)
        '        'displaying in phone number column of datagridview
        '        frm_Spy_Tech.DVG_SpyTech.Rows(Row_Index).Cells("Phone_Number").Value = PTCL_Number.ToString

        '        'storing data line by line
        '        PTCL_Details = "Name: " & PTCL_Reader(2)
        '        PTCL_Details = PTCL_Details & vbCrLf & "Address: " & PTCL_Reader(3)
        '        PTCL_Details = PTCL_Details & vbCrLf & "CNIC: " & PTCL_Reader(4)
        '        PTCL_Details = PTCL_Details & vbCrLf & "H Phone: " & PTCL_Reader(5)
        '        PTCL_Details = PTCL_Details & vbCrLf & "Reference Name: " & PTCL_Reader(6)
        '        PTCL_Details = PTCL_Details & vbCrLf & "Reference Address: " & PTCL_Reader(7)
        '        PTCL_Details = PTCL_Details & vbCrLf & "Reference CNIC: " & PTCL_Reader(8)
        '        PTCL_Details = PTCL_Details & vbCrLf & "C Type: " & PTCL_Reader(9)
        '        'displaying in detail column of datagridview
        '        frm_Spy_Tech.DVG_SpyTech.Rows(Row_Index).Cells("Detail").Value = PTCL_Details.ToString

        '    End While
        '    PTCL_Connection.Close()
        'Catch ex As Exception

        'End Try
    End Function
    Sub ConditionalDisplayBTS()
        Form1.Panel1.Visible = False
        Form1.GroupBox2.Visible = False
        Form1.GroupBox4.Visible = False
        Form1.GroupBox3.Visible = False
        Form1.chkCNIC.Visible = False
        Form1.cbFind_b_in_a.Visible = False
        Form1.Height = 250
        Form1.Button1.Left = 135
        Form1.Button1.Top = 175
    End Sub
End Module
