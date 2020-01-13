Public Class OBNK

    Dim SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Dim SBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Dim oInvoice As SAPbobsCOM.Documents
    Dim Duplicadas, Registradas, NExist As Integer

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
    End Sub

    Public Function findOBNK(ByVal FormUID As String, ByVal csDirectory As String)

        Dim coForm As SAPbouiCOM.Form
        Dim oGDataTable As SAPbouiCOM.DBDataSource
        Dim stQueryH, stQueryH2, stQueryH3, stQueryH4 As String
        Dim oRecSetH, oRecSetH2, oRecSetH3, oRecSetH4 As SAPbobsCOM.Recordset
        Dim Ref, CardCode, CardCodelog, CardName, CardNamelog, Account, User, stTabla, FrDate As String
        Dim Code, CurrentDate As String
        Dim DebAmount, CredAmnt As Double

        Try

            CurrentDate = Now.Year.ToString + "-" + Now.Month.ToString + "-" + Now.Day.ToString

            User = SBOCompany.UserName
            stTabla = "OACT"
            coForm = SBOApplication.Forms.Item(FormUID)

            oGDataTable = coForm.DataSources.DBDataSources.Item(stTabla)

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH3 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecSetH4 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Account = oGDataTable.GetValue("AcctCode", 0)

            stTabla = "OFLT"
            oGDataTable = coForm.DataSources.DBDataSources.Item(stTabla)
            FrDate = oGDataTable.GetValue("FrmDueDate", 0)

            stQueryH = "Select ""AcctCode"",""Ref"",""DueDate"",""DebAmount"",""CredAmnt"",""CardCode"",""CardName"" from ""OBNK"" where ""AcctCode""=" & Account & " and ""DueDate""='" & FrDate & "'"
            oRecSetH.DoQuery(stQueryH)

            oRecSetH.MoveFirst()

            If oRecSetH.RecordCount > 0 Then

                For i = 0 To oRecSetH.RecordCount - 1

                    Ref = oRecSetH.Fields.Item("Ref").Value
                    DebAmount = oRecSetH.Fields.Item("DebAmount").Value
                    CredAmnt = oRecSetH.Fields.Item("CredAmnt").Value
                    CardCode = oRecSetH.Fields.Item("CardCode").Value
                    CardName = oRecSetH.Fields.Item("CardName").Value

                    stQueryH2 = "Select ""U_CodeSN"",""U_NameSN"" from ""@LOG_OBNK"" where ""U_DueDate""='" & FrDate & "' and ""U_Debit""=" & DebAmount & " and ""U_Credit""=" & CredAmnt & " and ""U_AcctCode""=" & Account
                    oRecSetH2.DoQuery(stQueryH2)

                    If oRecSetH2.RecordCount > 0 Then

                        oRecSetH2.MoveLast()

                        CardCodelog = oRecSetH2.Fields.Item("U_CodeSN").Value
                        CardNamelog = oRecSetH2.Fields.Item("U_NameSN").Value

                        If CardCodelog <> CardCode And CardNamelog <> CardName Then

                            stQueryH4 = "Select case when length(count(""U_AcctCode"")+1)=1 then concat('00',TO_NVARCHAR(count(""U_AcctCode"")+1)) when length(count(""U_AcctCode"")+1)=2 then concat('0',TO_NVARCHAR(count(""U_AcctCode"")+1)) when length(count(""U_AcctCode"")+1)=3 then TO_NVARCHAR(count(""U_AcctCode"")+1) end as ""Codigo"" from ""@LOG_OBNK"" where ""U_AcctCode""=" & Account & " and ""U_DueDate""='" & FrDate & "'"
                            oRecSetH4.DoQuery(stQueryH4)

                            Code = FrDate + Account.Substring(7, 2).ToString + oRecSetH4.Fields.Item("Codigo").Value

                            stQueryH3 = "INSERT INTO ""@LOG_OBNK"" VALUES (" & Code & "," & Code & ",'" & FrDate & "','" & CurrentDate & "','" & Ref & "',null,'" & CardCode & "','" & CardName & "','" & Account & "','" & User & "'," & DebAmount & "," & CredAmnt & ")"
                            oRecSetH3.DoQuery(stQueryH3)

                        End If

                    Else

                        stQueryH4 = "Select case when length(count(""U_AcctCode"")+1)=1 then concat('00',TO_NVARCHAR(count(""U_AcctCode"")+1)) when length(count(""U_AcctCode"")+1)=2 then concat('0',TO_NVARCHAR(count(""U_AcctCode"")+1)) when length(count(""U_AcctCode"")+1)=3 then TO_NVARCHAR(count(""U_AcctCode"")+1) end as ""Codigo"" from ""@LOG_OBNK"" where ""U_AcctCode""=" & Account & " and ""U_DueDate""='" & FrDate & "'"
                        oRecSetH4.DoQuery(stQueryH4)

                        Code = FrDate + Account.Substring(7, 2).ToString + oRecSetH4.Fields.Item("Codigo").Value

                        stQueryH3 = "INSERT INTO ""@LOG_OBNK"" VALUES (" & Code & "," & Code & ",'" & FrDate & "','" & CurrentDate & "','" & Ref & "',null,'" & CardCode & "','" & CardName & "','" & Account & "','" & User & "'," & DebAmount & "," & CredAmnt & ")"
                        oRecSetH3.DoQuery(stQueryH3)

                    End If

                    oRecSetH.MoveNext()

                Next

            End If



        Catch ex As Exception

            SBOApplication.MessageBox("OBNK, fallo el barrido del log: " & ex.Message)

        End Try

    End Function

End Class
