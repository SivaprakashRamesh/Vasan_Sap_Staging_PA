Imports System.IO
Imports System.Data.SqlClient
Imports Microsoft.Win32
Module modStartup
    Dim dtheader As DataTable
    Dim dtline As DataTable
    Dim form_Name As String
    Dim ItemCode As String
    Dim tran As SqlTransaction
    Dim objMVasan As New MIPL_Vasan.Common_Module
    Dim dataset_log As DataSet
    Private fw As StreamWriter
    Dim LogFileName, FolderName As String
    Dim time As String = ""

    Public Sub Main()
        If FunCheckLicense() = False Then
            End
        End If
        Try
            ' '  MsgBox(Registry.CurrentUser.OpenSubKey("SOFTWARE", True))
            'System.Windows.Forms.Application.Run()
            'System.Threading.Thread.Sleep(10000)
            'MsgBox(Date.Now)
            'MsgBox(Date.Today)
            'MsgBox(Microsoft.VisualBasic.Format(Date.Now, "ddMMYYYY"))
            'MsgBox(Microsoft.VisualBasic.Format(Date.Now, "yyyyMMdd"))

            write_log("Exe Start")
            FolderName = "PASync_log" 'Trim(Date.Now.ToLongDateString)
            FolderName = FolderName.Replace(",", "_")
            FolderName = FolderName.Replace(" ", "_")
            If Not Directory.Exists("C:\QL\" & FolderName) Then
                Directory.CreateDirectory("C:\QL\" & FolderName)
            End If
            dataset_log = New DataSet
            'MDBName = "VSN_TECH"
            connection_StagingDB()
            connection_SAPDB()
            write_log("Company Connected")
            getting_SAP_information()

            Add_location_Master()
            System.Threading.Thread.Sleep(1000)

            Vendor_Master()
            System.Threading.Thread.Sleep(1000)

            TaxCode_determiantion()
            System.Threading.Thread.Sleep(1000)

            'PriceList_PA()
            'System.Threading.Thread.Sleep(1000)

            ITEMMASTER_ADD_NEW1()
            System.Threading.Thread.Sleep(1000)

            item_Group()
            System.Threading.Thread.Sleep(1000)

            Add_InventoryTransfer()
            System.Threading.Thread.Sleep(1000)

            
            'frmMcommon.ShowDialog()

            'Dim startinfo As New ProcessStartInfo
            'startinfo.FileName = "S:\MIPL_APPS\Outbound_Exe\SAP_Staging_Proactive\MIPL_Outbound_ProActive.exe"
            ''startinfo.FileName = "S:\MIPL_APPS\Outbound_Exe\SAP_Staging_Proactive\MIPL_Outbound_ProActive.exe"
            'form_Name = "ITEMMASTER"
            'ItemCode = ""
            'Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
            'If CommandLineArgs.Count > 0 Then
            '    form_Name = CommandLineArgs(0)
            '    ItemCode = CommandLineArgs(1)
            'End If
            'form_Name = "itemmaster"
            'ItemCode = "Z-0003"
            'If form_Name.ToString.ToUpper = "ITEMMASTER" Then
            '    startinfo.Arguments = "ITEM_MASTER"
            'ITEMMASTER_ADD_NEW1()
            'System.Threading.Thread.Sleep(1000)
            '    'ItemMaster_Add(ItemCode)
            'ElseIf form_Name.ToString.ToUpper = "ITEMGROUP" Then
            '    startinfo.Arguments = "ITEM_GROUP"
            'item_Group()
            'ElseIf form_Name.ToString.ToUpper = "LOCATIONMASTER" Then
            '    startinfo.Arguments = "LOCATION_MASTER"
            'Add_location_Master()
            'System.Threading.Thread.Sleep(1000)

            '    ' Add_location_Master1()
            'ElseIf form_Name.ToString.ToUpper = "VENDORMASTER" Then
            '    startinfo.Arguments = "VENDOR_MASTER"
            'Vendor_Master()
            'System.Threading.Thread.Sleep(1000)
            '    'ElseIf form_Name.ToString.ToUpper = "SALESEMPMASTER" Then
            '    '    startinfo.Arguments = "ITEM"
            '    '    Add_salesemployee_master()
            'ElseIf form_Name.ToString.ToUpper = "TAX" Then
            '    startinfo.Arguments = "TAX_MASTER"
            '            TaxCode_determiantion()
            'ElseIf form_Name.ToString.ToUpper = "INVENTRYTRANSFER" Then
            '    startinfo.Arguments = "STOCK_OUT"

            'System.Threading.Thread.Sleep(1000)
            'Add_InventoryTransfer()
            'ElseIf form_Name.ToString.ToUpper = "PRICE_LIST" Then
            '    startinfo.Arguments = "PRICE_LIST"
            'System.Threading.Thread.Sleep(1000)
            'PriceList_PA()
            'ElseIf form_Name.ToString.ToUpper = "NORMAL" Then
            '    startinfo.Arguments = "ITEM_GROUP"
            '    item_Group()
            '    Process.Start(startinfo)

            '    startinfo.Arguments = "ITEM_MASTER"
            '    ITEMMASTER_ADD_NEW1()
            '    Process.Start(startinfo)

            '    startinfo.Arguments = "LOCATION_MASTER"
            '    Add_location_Master()
            '    Process.Start(startinfo)

            '    startinfo.Arguments = "VENDOR_MASTER"
            '    Vendor_Master()
            '    Process.Start(startinfo)

            '    startinfo.Arguments = "TAX_MASTER"
            '    TaxCode_determiantion()
            '    Process.Start(startinfo)

            '    startinfo.Arguments = "STOCK_OUT"
            '    Add_InventoryTransfer()
            '    Process.Start(startinfo)

            '    startinfo.Arguments = "PRICE_LIST"
            '    PriceList_PA()
            'End If
            'strsql = "set @location='c:\windows\system32\cscript.exe //nologo  E:\PRAVEEN-SOURCE\Vasan\Outbound\Vasan_Staging_Idea\Vasan_Staging_Idea\bin\Debug\caller.vbs /param1:'+@param1 +' /param2:'+@param2 "
            'strsql += vbCrLf + " exec master..xp_cmdshell @location"
            con_SAP.Close()
            con_staging.Close()
            con_SAP.Dispose()
            con_staging.Dispose()
            'Application.Exit()


            ''Dim con As New SqlConnection("data source=bssteam8-pc;initial catalog=temp_db;user id=sa;password=Mipl1234")

            ''Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)


            ''If con.State = System.Data.ConnectionState.Closed Then
            ''con.Open()
            '' End If

            '' Dim Query As String = "INSERT INTO TEMP1 VALUES('" & C:\Users\user1\Desktop\MIPL_Outbound_ProActive.exe.Now.ToString() & "','" & DateTime.Now.ToString() & "')"

            ''Dim com As New SqlCommand(Query, con)

            ''com.ExecuteNonQuery()

            ''con.Close()
            ''con.Dispose()
            'Process.Start("C:\Users\user1\Desktop\MIPL_Outbound_ProActive.exe")
            'Process.Start(startinfo)

            'Application.Exit()
            End

        Catch ex As Exception
            write_log("Main Issue: " + ex.Message)
            objMVasan.Register_Log(1005, "GENERAL", dataset_log, ex.ToString, objMVasan.Get_Next_Key_Value(1005), 1, Date.Now)
            con_SAP.Close()
            con_staging.Close()
            con_SAP.Dispose()
            con_staging.Dispose()
            'Process.Start("C:\Users\user1\Desktop\MIPL_Outbound_ProActive.exe")
            'Dim startinfo As New ProcessStartInfo
            'Process.Start(startinfo)

            Application.Exit()
            'MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Vendor_Master()
        Try
            'Val(System.Text.RegularExpressions.Regex.Replace(objform.Items.Item("txttorc").Specific.string, "[a-zA-Z\b\s!`~@#$%^&*()__+=\|}{:;><,/?]", "")) > 0 And objform.Items.Item("5").Specific.string <> "" Then
            write_log(vbNewLine + "-----  BP Master Check --------")
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from s_ocrd"
            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            write_log("BP Master Header Insertion : " + strsql)

            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    'strsql = "insert into S_ocrd (VENDOR_CODE,VENDOR_NAME,VENDOR_ADDRESS,CURRENCY_CODE,LUPDATE,UPDATE_FLAG)"
                    'strsql += vbCrLf + "select T1.cardcode,T1.CardName,T1.Address + '_' + T1.Zipcode'Vendor Address',T1.currency,T2.U_LUPDATE,1  from " & MDBName & "..OCRD T1 join " & MDBName & "..flag_table T2 on T1.CardCode=T2.U_formcode "
                    'strsql += vbCrLf + " where T2.U_formname='VENDORMASTER'"
                    strsql = "insert into S_ocrd (VENDOR_CODE,BPtype,VENDOR_NAME,VENDOR_ADDRESS,CURRENCY_CODE,LUPDATE,UPDATE_FLAG,Add_Update,BPGroup)"
                    strsql += vbCrLf + "select  T1.cardcode ,T1.cardType,Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(T1.CardName,'""',' '),'&',' '),'''',' '),':',' '),'’',' '),'_',' '),'-',' '),'+',' '),'/',' '),'@',' '),'#',' '),'*',' '),'(',' '),')',' '),Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(T1.Address + ' ' + T1.Zipcode,'""',' '),'&',' '),'''',' '),':',' '),'’',' '),'_',' '),'-',' '),'+',' '),'/',' '),'@',' '),'#',' '),'*',' '),'(',' '),')',' ')'Vendor Address',T1.currency,getdate(),1,'A',(select Top 1 T3.GroupName  from " & MDBName & "..ocrg T3 where T3.GroupCode=T1.Groupcode ) 'BPGroup'  from " & MDBName & "..OCRD T1 where Isnull(T1.CardName,'')<>'' and  T1.Groupcode in (select U_BPGROUP from " & MDBName & "..[@miplBPGROUP] where U_PA='Y')"
                Else
                    'System.Text.RegularExpressions.Regex.Replace(T1.CardName, "[a-zA-Z\b\s!`~@#$%^&*()__+=\|}{:;><,/?]", "")
                    'strsql = "insert into S_ocrd (VENDOR_CODE,BPtype,VENDOR_NAME,VENDOR_ADDRESS,CURRENCY_CODE,LUPDATE,UPDATE_FLAG,Add_update,BPGroup)"
                    ''strsql += vbCrLf + "select T1.cardcode,T1.cardtype,Replace(Replace(Replace(Replace(Replace(T1.CardName,'""',' '),'&',' '),'''',' '),':',' '),'’',' '),Replace(Replace(Replace(Replace(Replace(T1.Address + '_' + T1.Zipcode,'""',' '),'&',' '),'''',' '),':',' '),'’',' ')'Vendor Address',T1.currency,T2.U_LUPDATE,1,T2.add_update,(select Top 1 T3.GroupName  from " & MDBName & "..ocrg T3 where T3.GroupCode=T1.Groupcode ) 'BPGroup'  from " & MDBName & "..OCRD T1 join " & MDBName & "..flag_table T2 on T1.CardCode=T2.U_formcode "
                    'strsql += vbCrLf + "select  distinct T1.cardcode ,T1.cardtype,Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(T1.CardName,'""',' '),'&',' '),'''',' '),':',' '),'’',' '),'_',' '),'-',' '),'+',' '),'/',' '),'@',' '),'#',' '),'*',' '),'(',' '),')',' '),Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(T1.Address + ' ' + T1.Zipcode,'""',' '),'&',' '),'''',' '),':',' '),'’',' '),'_',' '),'-',' '),'+',' '),'/',' '),'@',' '),'#',' '),'*',' '),'(',' '),')',' ')'Vendor Address',T1.currency,T2.U_LUPDATE,1,T2.add_update,(select Top 1 T3.GroupName  from " & MDBName & "..ocrg T3 where T3.GroupCode=T1.Groupcode ) 'BPGroup'  from " & MDBName & "..OCRD T1 join " & MDBName & "..flag_table T2 on T1.CardCode=T2.U_formcode "
                    'strsql += vbCrLf + " where T2.U_formname='VENDORMASTER'  AND Isnull(T1.CardName,'')<>'' and  T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "' and T1.Groupcode in (select U_BPGROUP from " & MDBName & "..[@miplBPGROUP] where U_PA='Y')"
                    strsql = "insert into S_ocrd (VENDOR_CODE,BPtype,VENDOR_NAME,VENDOR_ADDRESS,CURRENCY_CODE,LUPDATE,UPDATE_FLAG,Add_update,BPGroup)"
                    strsql += vbCrLf + "select  distinct T1.cardcode ,T1.cardtype,Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(T1.CardName,'""',' '),'&',' '),'''',' '),':',' '),'’',' '),'_',' '),'-',' '),'+',' '),'/',' '),'@',' '),'#',' '),'*',' '),'(',' '),')',' '),Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(T1.Address + ' ' + T1.Zipcode,'""',' '),'&',' '),'''',' '),':',' '),'’',' '),'_',' '),'-',' '),'+',' '),'/',' '),'@',' '),'#',' '),'*',' '),'(',' '),')',' ')'Vendor Address',T1.currency,T2.U_LUPDATE,1,T2.add_update,(select Top 1 T3.GroupName  from " & MDBName & "..ocrg T3 where T3.GroupCode=T1.Groupcode ) 'BPGroup'  from VSN19.dbo.OCRD T1 join " & MDBName & "..flag_table T2 on T1.CardCode=T2.U_formcode "
                    strsql += vbCrLf + "where T2.U_formname='VENDORMASTER'  AND Isnull(T1.CardName,'')<>'' and  T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "' and T1.Groupcode in (select U_BPGROUP from  " & MDBName & "..[@miplBPGROUP] where U_PA='Y')"
                End If
            End If
            write_log("BP Master Added : " + strsql)
            command_staging.CommandText = strsql
            command_staging.ExecuteNonQuery()
            'objMVasan.Register_Log(1005, "VENDOR_MASTER", dataset_log, "Success", objMVasan.Get_Next_Key_Value(1005), 0, Date.Now)


        Catch ex As Exception
            write_log("BP Master Issue : " + ex.Message)
            'objMVasan.Register_Log(1005, "VENDOR_MASTER", dataset_log, ex.ToString, objMVasan.Get_Next_Key_Value(1005), 1, Date.Now)
        End Try
    End Sub
    Private Sub ITEMMASTER_ADD_NEW1()
        Try
            write_log(vbNewLine + "-----  Item Master --------")
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from s_oitm with (nolock)"

            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            write_log("Item Master Header Insertion : " + strsql)

            If dt.Rows.Count > 0 Then
                'MsgBox(CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff"))
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    'strsql = "insert into S_oitm (PRODUCT_CODE,PRODUCT_NAME,PRODUCT_NATURE,PRODUCT_TYPE,UOM,PRODUCT_CATEGORY,PRODUCT_ATTRIBUTES,LUPDATE,UPDATE_FLAG,ADD_UPDATE) "
                    'strsql += vbCrLf + "select T1.ItemCode'Product Code',T1.ItemName'Product Name',CASE when T1.ManSerNum='N' and T1.ManBtchNum='N' then 2 when T1.ManSerNum='Y' then 1 else 0 end as 'Prduct Natrue'"
                    'strsql += vbCrLf + ",case when (select top 1 T2.itmstypcod from " & MDBName & "..View_ItemProperty T2 where T2.itemcode=T1.ITEMCODE)='1' then 0 when (select top 1 T2.itmstypcod from " & MDBName & "..View_ItemProperty T2 where T2.itemcode=T1.ITEMCODE)='16'  then 1 end as 'Product Type',T1.InvntryUom 'UOM',T1.ItmsGrpCod 'Product Category',"
                    'strsql += vbCrLf + "case when left(T1.ItemName,(select len(U_DSCP)  from " & MDBName & "..OITB  where ItmsGrpCod=T1.ItmsGrpCod))="
                    'strsql += vbCrLf + " (select (U_DSCP)  from " & MDBName & "..OITB  where ItmsGrpCod=T1.ItmsGrpCod) then replace((right(T1.ItemName,len(T1.ItemName)-1-(select len(U_DSCP)  from " & MDBName & "..OITB  where ItmsGrpCod=T1.ItmsGrpCod))),'_','|') else replace(ItemName,'_','|')  end as "
                    ''strsql += vbCrLf + ",1 'Product Type',T1.InvntryUom 'UOM',(select top 1 T12.ItmsGrpNam  from " & MDBName & "..OITB T12 where T12.ItmsGrpCod=T1.ItmsGrpCod) 'Product Category',"
                    'strsql += vbCrLf + "'PRODUCT_ATTRIBUTES',getdate()'LUPDate',1 'Update_Flag','A' 'add_update'"
                    'strsql += vbCrLf + "from " & MDBName & "..OITM T1 where T1.ItmsGrpCod in (select U_itemgroupcode from " & MDBName & "..[@itemgroup_udt] where U_receiver='PA') and T1.U_legcod<>'wrong'"
                    strsql = "insert into S_oitm (PRODUCT_CODE,PRODUCT_NAME,PRODUCT_NATURE,PRODUCT_TYPE,UOM,PRODUCT_CATEGORY,PRODUCT_ATTRIBUTES,LUPDATE,UPDATE_FLAG,ADD_UPDATE) "
                    strsql += vbCrLf + "select distinct T1.ItemCode,T1.ItemName 'Product Name',CASE when T1.ManSerNum='N' and T1.ManBtchNum='N' then 2 when T1.ManSerNum='Y' then 1 else 0 end as 'Prduct Natrue'"
                    strsql += vbCrLf + ",case when (select top 1 T2.itmstypcod from " & MDBName & "..View_ItemProperty T2 where T2.itemcode=T1.ITEMCODE)='1' then 0 when (select top 1 T2.itmstypcod from " & MDBName & "..View_ItemProperty T2 where T2.itemcode=T1.ITEMCODE)='16'  then 1 end as 'Product Type',T1.InvntryUom 'UOM',(select CASE when ISNULL(T12.U_excode,'')='' then Convert(varchar(15),t12.ItmsGrpCod)  else T12.U_excode end  from " & MDBName & "..OITB  as T12 where T12.ItmsGrpCod=T1.ItmsGrpCod) 'Product Category',"
                    strsql += vbCrLf + "case when left(T1.ItemName,(select len(U_DSCP)  from " & MDBName & "..OITB  where ItmsGrpCod=T1.ItmsGrpCod))="
                    strsql += vbCrLf + " (select (U_DSCP)  from " & MDBName & "..OITB  where ItmsGrpCod=T1.ItmsGrpCod) then replace((right(T1.ItemName,len(T1.ItemName)-1-(select len(U_DSCP)  from " & MDBName & "..OITB  where ItmsGrpCod=T1.ItmsGrpCod))),'_','|') else replace(ItemName,'_','|')  end as "
                    'strsql += vbCrLf + ",1 'Product Type',T1.InvntryUom 'UOM',(select top 1 T12.ItmsGrpNam  from " & MDBName & "..OITB T12 where T12.ItmsGrpCod=T1.ItmsGrpCod) 'Product Category',"
                    strsql += vbCrLf + "'PRODUCT_ATTRIBUTES',getdate()'LUPDate',1 'Update_Flag','A' 'add_update'"
                    strsql += vbCrLf + "from " & MDBName & "..OITM T1 where T1.ItmsGrpCod in (select U_itemgroupcode from " & MDBName & "..[@itemgroup_udt] where U_receiver='PA') and Isnull(T1.U_legcod,'')<>'wrong'"
                Else
                    'strsql = "insert into S_oitm (PRODUCT_CODE,PRODUCT_NAME,PRODUCT_NATURE,PRODUCT_TYPE,UOM,PRODUCT_CATEGORY,PRODUCT_ATTRIBUTES,LUPDATE,UPDATE_FLAG,ADD_UPDATE)"
                    ''strsql += vbCrLf + "select T1.ItemCode'Product Code',T1.ItemName'Product Name',CASE when T1.ManSerNum='N' and T1.ManBtchNum='N' then 2 when T1.ManSerNum='Y' then 1 else 0 end as 'Prduct Natrue'"
                    'strsql += vbCrLf + "Select distinct aa.ItemCode,replace(aa.[Product Name],'&','AND'),aa.[Prduct Natrue] ,aa.[Product Type] ,isnull(aa.UOM,'Nos'),"
                    'strsql += vbCrLf + "aa.[Product Category] ,replace(aa.PRODUCT_ATTRIBUTES,'&','AND'),aa.LUPDate ,aa.Update_Flag ,T10.add_update from (select T1.ItemCode,T1.ItemName 'Product Name',CASE when T1.ManSerNum='N' and T1.ManBtchNum='N' then 2 when T1.ManSerNum='Y' then 1 else 0 end as 'Prduct Natrue' "
                    'strsql += vbCrLf + ",case when (select top 1 T2.itmstypcod from VSN19..View_ItemProperty T2 where T2.itemcode=T1.ITEMCODE)='1' then 0 when (select top 1 T2.itmstypcod from VSN19..View_ItemProperty T2 where T2.itemcode=T1.ITEMCODE)='16'  then 1 end as 'Product Type',T1.InvntryUom 'UOM',(select CASE when ISNULL(T12.U_excode,'')='' then Convert(varchar(15),t12.ItmsGrpCod)  else T12.U_excode end  from VSN19..OITB  as T12 where T12.ItmsGrpCod=T1.ItmsGrpCod) 'Product Category', "
                    'strsql += vbCrLf + "case when left(T1.ItemName,(select len(U_DSCP)  from VSN19..OITB  where ItmsGrpCod=T1.ItmsGrpCod))= "
                    'strsql += vbCrLf + " (select (U_DSCP)  from VSN19..OITB  where ItmsGrpCod=T1.ItmsGrpCod) then replace((right(T1.ItemName,len(T1.ItemName)-1-(select len(U_DSCP)  from VSN19..OITB  where ItmsGrpCod=T1.ItmsGrpCod))),'_','|') else replace(ItemName,'_','|')  end as  "
                    'strsql += vbCrLf + "'PRODUCT_ATTRIBUTES',getdate()'LUPDate',1 'Update_Flag'"
                    'strsql += vbCrLf + "from VSN19..OITM T1 where T1.ItmsGrpCod in (select U_itemgroupcode from VSN19..[@itemgroup_udt] where U_receiver='PA'))aa "
                    'strsql += vbCrLf + "JOIN " & MDBName & "..FLAG_TABLE T10 ON T10.U_FORMCODE=aa.ITEMCODE where T10.U_FORMNAME='ITEMMASTER' AND T10.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "' and T10.U_LUPDATE>='20121117' "
                    strsql = "insert into S_oitm (PRODUCT_CODE,PRODUCT_NAME,PRODUCT_NATURE,PRODUCT_TYPE,UOM,PRODUCT_CATEGORY,PRODUCT_ATTRIBUTES,LUPDATE,UPDATE_FLAG,ADD_UPDATE)"
                    strsql += vbCrLf + "Select distinct aa.ItemCode,replace(aa.[Product Name],'&','AND'),aa.[Prduct Natrue] ,aa.[Product Type] ,isnull(aa.UOM,'Nos'),"
                    strsql += vbCrLf + "aa.[Product Category] ,replace(aa.PRODUCT_ATTRIBUTES,'&','AND'),aa.LUPDate ,aa.Update_Flag ,T10.add_update from (select T1.ItemCode,T1.ItemName 'Product Name',CASE when T1.ManSerNum='N' and T1.ManBtchNum='N' then 2 when T1.ManSerNum='Y' then 1 else 0 end as 'Prduct Natrue' "
                    strsql += vbCrLf + ",case when (select top 1 T2.itmstypcod from " & MDBName & "..View_ItemProperty T2 where T2.itemcode=T1.ITEMCODE)='1' then 0 when (select top 1 T2.itmstypcod from " & MDBName & "..View_ItemProperty T2 where T2.itemcode=T1.ITEMCODE)='16'  then 1 end as 'Product Type',T1.InvntryUom 'UOM',(select CASE when ISNULL(T12.U_excode,'')='' then Convert(varchar(15),t12.ItmsGrpCod)  else T12.U_excode end  from VSN19.dbo.OITB  as T12 where T12.ItmsGrpCod=T1.ItmsGrpCod) 'Product Category', "
                    strsql += vbCrLf + "case when left(T1.ItemName,(select len(U_DSCP)  from " & MDBName & "..OITB  where ItmsGrpCod=T1.ItmsGrpCod))= "
                    strsql += vbCrLf + "(select (U_DSCP)  from " & MDBName & "..OITB  where ItmsGrpCod=T1.ItmsGrpCod) then replace((right(T1.ItemName,len(T1.ItemName)-1-(select len(U_DSCP)  from " & MDBName & "..OITB  where ItmsGrpCod=T1.ItmsGrpCod))),'_','|') else replace(ItemName,'_','|')  end as  "
                    strsql += vbCrLf + "'PRODUCT_ATTRIBUTES',getdate()'LUPDate',1 'Update_Flag'"
                    strsql += vbCrLf + "from " & MDBName & "..OITM T1 where T1.ItmsGrpCod in (select U_itemgroupcode from " & MDBName & "..[@itemgroup_udt] where U_receiver='PA'))aa "
                    strsql += vbCrLf + "JOIN " & MDBName & "..FLAG_TABLE T10 ON T10.U_FORMCODE=aa.ITEMCODE where T10.U_FORMNAME='ITEMMASTER' AND T10.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "' and T10.U_LUPDATE>='20121117' "
                End If
                write_log("Item Master Added : " + strsql)
                'MsgBox(strsql)
                command_staging.CommandText = strsql
                command_staging.CommandTimeout = 100
                command_staging.ExecuteNonQuery()
                'objMVasan.Register_Log(1005, "ITEM_MASTER", dataset_log, "Success", objMVasan.Get_Next_Key_Value(1005), 0, Date.Now)

            End If
        Catch ex As Exception
            write_log("Item Master Issue : " + ex.ToString())
            'objMVasan.Register_Log(1005, "ITEM_MASTER", dataset_log, ex.ToString, objMVasan.Get_Next_Key_Value(1005), 1, Date.Now)
        End Try
    End Sub

    Private Sub ITEMMASTER_ADD_NEW_old()
        Try
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from s_oitm"
            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                'MsgBox(CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff"))
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    strsql = "insert into S_oitm (ItemCode,ItemName,ITEM_ATTRIBUTES,Firmname,InvntryUom,BuyUnitMsr,SalUnitMsr,"
                    strsql += vbCrLf + "ItmsGrpCod,itmsgrpname,ManSerNum,ManBtchNum,WarrntTmpl,LUPDATE)"

                    strsql += vbCrLf + "select T1.ItemCode,T1.ItemName,'',(SELECT T3.FIRMNAME  FROM " & MDBName & "..OMRC T3 WHERE T3.FIRMCODE=T1.FIRMCODE)'FirmName',T1.InvntryUom,T1.BuyUnitMsr,T1.SalUnitMsr,"
                    strsql += vbCrLf + "T1.ItmsGrpCod,(SELECT T4.ITMSGRPNAM  FROM " & MDBName & "..OITB T4 WHERE T4.ITMSGRPCOD=T1.ITMSGRPCOD)'ITMSGRPNAME',T1.ManSerNum,T1.ManBtchNum,T1.WarrntTmpl,T10.U_LUPDATE"
                    strsql += vbCrLf + "from " & MDBName & "..OITM T1 JOIN " & MDBName & "..FLAG_TABLE T10 ON T10.U_FORMCODE=T1.ITEMCODE where T10.U_FORMNAME='ITEMMASTER'"

                    strsql += vbCrLf + "insert into s_itm1 (itemcode,statecode,efctfrom,taxcode,status)"
                    strsql += vbCrLf + "SELECT T1.KeyFld_1_V 'ItemCode',T1.KeyFld_2_V 'StateCode',T2.EfctFrom'Effect From',T2.TaxCode,(select lock from " & MDBName & "..OSTC where Code=T2.TaxCode )'Status'  "
                    strsql += vbCrLf + " FROM " & MDBName & "..TCD1 T0  INNER JOIN " & MDBName & "..TCD2 T1 ON T0.AbsId = T1.Tcd1Id INNER JOIN " & MDBName & "..TCD3 T2 ON T1.AbsId = T2.Tcd2Id  join " & MDBName & "..flag_table t5 on t5.U_FORMCODE=t1.KeyFld_1_V "
                    strsql += vbCrLf + " where T0.KeyFld_1='2' and T0.KeyFld_2='6' and t5.U_FORMNAME='ItemMaster'"
                Else

                    strsql = "insert into S_oitm (ItemCode,ItemName,ITEM_ATTRIBUTES,Firmname,InvntryUom,BuyUnitMsr,SalUnitMsr,"
                    strsql += vbCrLf + "ItmsGrpCod,itmsgrpname,ManSerNum,ManBtchNum,WarrntTmpl,LUPDATE)"

                    strsql += vbCrLf + "select distinct T1.ItemCode,T1.ItemName,'',(SELECT T3.FIRMNAME  FROM " & MDBName & "..OMRC T3 WHERE T3.FIRMCODE=T1.FIRMCODE)'FirmName',T1.InvntryUom,T1.BuyUnitMsr,T1.SalUnitMsr,"
                    strsql += vbCrLf + "T1.ItmsGrpCod,(SELECT T4.ITMSGRPNAM  FROM " & MDBName & "..OITB T4 WHERE T4.ITMSGRPCOD=T1.ITMSGRPCOD)'ITMSGRPNAME',T1.ManSerNum,T1.ManBtchNum,T1.WarrntTmpl,T10.U_LUPDATE"
                    strsql += vbCrLf + "from " & MDBName & "..OITM T1 JOIN " & MDBName & "..FLAG_TABLE T10 ON T10.U_FORMCODE=T1.ITEMCODE where T10.U_FORMNAME='ITEMMASTER' AND T10.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"

                    strsql += vbCrLf + "insert into s_itm1 (itemcode,statecode,efctfrom,taxcode,status)"
                    strsql += vbCrLf + "SELECT distinct T1.KeyFld_1_V 'ItemCode',T1.KeyFld_2_V 'StateCode',T2.EfctFrom'Effect From',T2.TaxCode,(select lock from " & MDBName & "..OSTC where Code=T2.TaxCode )'Status'  "
                    strsql += vbCrLf + " FROM " & MDBName & "..TCD1 T0  INNER JOIN " & MDBName & "..TCD2 T1 ON T0.AbsId = T1.Tcd1Id INNER JOIN " & MDBName & "..TCD3 T2 ON T1.AbsId = T2.Tcd2Id  join " & MDBName & "..flag_table t5 on t5.U_FORMCODE=t1.KeyFld_1_V "
                    strsql += vbCrLf + " where T0.KeyFld_1='2' and T0.KeyFld_2='6' and t5.U_FORMNAME='ItemMaster' AND T5.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                End If

                command_staging.CommandText = strsql
                'MsgBox(strsql)
                command_staging.ExecuteNonQuery()
            End If

            'strsql = "select ItemCode,ItemName,UserText,(SELECT FIRMNAME  FROM OMRC WHERE FIRMCODE=T1.FIRMCODE)'FirmName',InvntItem,PrchseItem,SellItem,AssetItem,InvntryUom,BuyUnitMsr,SalUnitMsr,"
            'strsql += vbCrLf + "ItmsGrpCod,(SELECT ITMSGRPNAM  FROM OITB WHERE ITMSGRPCOD=T1.ITMSGRPCOD)'ITMSGRPNAME',QryGroup1,QryGroup2,QryGroup3,QryGroup4,QryGroup5,QryGroup6,QryGroup7,QryGroup8,QryGroup9,QryGroup10,"
            'strsql += vbCrLf + "QryGroup11,QryGroup12,QryGroup13,QryGroup14,QryGroup15,QryGroup16,QryGroup17,QryGroup18,QryGroup19,QryGroup20,"
            'strsql += vbCrLf + "QryGroup21,QryGroup22,QryGroup23,QryGroup24,QryGroup25,QryGroup26,QryGroup27,QryGroup28,QryGroup29,QryGroup30,"
            'strsql += vbCrLf + "QryGroup31,QryGroup32,QryGroup33,QryGroup34,QryGroup35,QryGroup36,QryGroup37,QryGroup38,QryGroup39,QryGroup40,"
            'strsql += vbCrLf + "QryGroup41,QryGroup42,QryGroup43,QryGroup44,QryGroup45,QryGroup46,QryGroup47,QryGroup48,QryGroup49,QryGroup50,"
            'strsql += vbCrLf + "QryGroup51,QryGroup52,QryGroup53,QryGroup54,QryGroup55,QryGroup56,QryGroup57,QryGroup58,QryGroup59,QryGroup60,"
            'strsql += vbCrLf + "QryGroup61,QryGroup62,QryGroup63,QryGroup64,EvalSystem,ManSerNum,ManBtchNum,WarrntTmpl"
            'strsql += vbCrLf + "from OITM T1 where itemcode in (SELECT U_FORMCODE FROM FLAG_TABLE WHERE U_LUPDATE >='" & dt.Rows(0).Item("LUPDATE") & "' AND U_formname='ITEMMASTER')"
            'dtheader = New DataTable
            'da = New SqlDataAdapter(strsql, con_SAP)
            'da.Fill(dtheader)
            'If dtheader.Rows.Count > 0 Then
            '    strsql = "insert into oitm (ItemCode,ItemName,UserText,Firmname,InvntItem,PrchseItem,SellItem,AssetItem,InvntryUom,BuyUnitMsr,SalUnitMsr,"
            '    strsql += vbCrLf + "ItmsGrpCod,ITeMGRPname,QryGroup1,QryGroup2,QryGroup3,QryGroup4,QryGroup5,QryGroup6,QryGroup7,QryGroup8,QryGroup9,QryGroup10,"
            '    strsql += vbCrLf + "QryGroup11,QryGroup12,QryGroup13,QryGroup14,QryGroup15,QryGroup16,QryGroup17,QryGroup18,QryGroup19,QryGroup20,"
            '    strsql += vbCrLf + "QryGroup21,QryGroup22,QryGroup23,QryGroup24,QryGroup25,QryGroup26,QryGroup27,QryGroup28,QryGroup29,QryGroup30,"
            '    strsql += vbCrLf + "QryGroup31,QryGroup32,QryGroup33,QryGroup34,QryGroup35,QryGroup36,QryGroup37,QryGroup38,QryGroup39,QryGroup40,"
            '    strsql += vbCrLf + "QryGroup41,QryGroup42,QryGroup43,QryGroup44,QryGroup45,QryGroup46,QryGroup47,QryGroup48,QryGroup49,QryGroup50,"
            '    strsql += vbCrLf + "QryGroup51,QryGroup52,QryGroup53,QryGroup54,QryGroup55,QryGroup56,QryGroup57,QryGroup58,QryGroup59,QryGroup60,"
            '    strsql += vbCrLf + "QryGroup61,QryGroup62,QryGroup63,QryGroup64,EvalSystem,ManSerNum,ManBtchNum,WarrntTmpl)"
            '    strsql += vbCrLf + "values"
            '    strsql += vbCrLf + "('" & dtheader.Rows(0).Item("ItemCode") & "','" & dtheader.Rows(0).Item("ItemName") & "','" & dtheader.Rows(0).Item("UserText") & "','" & dtheader.Rows(0).Item("FirmName") & "','" & dtheader.Rows(0).Item("InvntItem") & "','" & dtheader.Rows(0).Item("PrchseItem") & "','" & dtheader.Rows(0).Item("SellItem") & "','" & dtheader.Rows(0).Item("AssetItem") & "','" & dtheader.Rows(0).Item("InvntryUom") & "','" & dtheader.Rows(0).Item("BuyUnitMsr") & "','" & dtheader.Rows(0).Item("SalUnitMsr") & "',"
            '    strsql += vbCrLf + "'" & dtheader.Rows(0).Item("ItmsGrpCod") & "','" & dtheader.Rows(0).Item("ITMSGRPNAME") & "','" & dtheader.Rows(0).Item("QryGroup1") & "','" & dtheader.Rows(0).Item("QryGroup2") & "','" & dtheader.Rows(0).Item("QryGroup3") & "','" & dtheader.Rows(0).Item("QryGroup4") & "','" & dtheader.Rows(0).Item("QryGroup5") & "','" & dtheader.Rows(0).Item("QryGroup6") & "','" & dtheader.Rows(0).Item("QryGroup7") & "','" & dtheader.Rows(0).Item("QryGroup8") & "','" & dtheader.Rows(0).Item("QryGroup9") & "','" & dtheader.Rows(0).Item("QryGroup10") & "',"
            '    strsql += vbCrLf + "'" & dtheader.Rows(0).Item("QryGroup11") & "','" & dtheader.Rows(0).Item("QryGroup12") & "','" & dtheader.Rows(0).Item("QryGroup13") & "','" & dtheader.Rows(0).Item("QryGroup14") & "','" & dtheader.Rows(0).Item("QryGroup15") & "','" & dtheader.Rows(0).Item("QryGroup16") & "','" & dtheader.Rows(0).Item("QryGroup17") & "','" & dtheader.Rows(0).Item("QryGroup18") & "','" & dtheader.Rows(0).Item("QryGroup19") & "','" & dtheader.Rows(0).Item("QryGroup20") & "',"
            '    strsql += vbCrLf + "'" & dtheader.Rows(0).Item("QryGroup21") & "','" & dtheader.Rows(0).Item("QryGroup22") & "','" & dtheader.Rows(0).Item("QryGroup23") & "','" & dtheader.Rows(0).Item("QryGroup24") & "','" & dtheader.Rows(0).Item("QryGroup25") & "','" & dtheader.Rows(0).Item("QryGroup26") & "','" & dtheader.Rows(0).Item("QryGroup27") & "','" & dtheader.Rows(0).Item("QryGroup28") & "','" & dtheader.Rows(0).Item("QryGroup29") & "','" & dtheader.Rows(0).Item("QryGroup30") & "',"
            '    strsql += vbCrLf + "'" & dtheader.Rows(0).Item("QryGroup31") & "','" & dtheader.Rows(0).Item("QryGroup32") & "','" & dtheader.Rows(0).Item("QryGroup33") & "','" & dtheader.Rows(0).Item("QryGroup34") & "','" & dtheader.Rows(0).Item("QryGroup35") & "','" & dtheader.Rows(0).Item("QryGroup36") & "','" & dtheader.Rows(0).Item("QryGroup37") & "','" & dtheader.Rows(0).Item("QryGroup38") & "','" & dtheader.Rows(0).Item("QryGroup39") & "','" & dtheader.Rows(0).Item("QryGroup40") & "',"
            '    strsql += vbCrLf + "'" & dtheader.Rows(0).Item("QryGroup41") & "','" & dtheader.Rows(0).Item("QryGroup42") & "','" & dtheader.Rows(0).Item("QryGroup43") & "','" & dtheader.Rows(0).Item("QryGroup44") & "','" & dtheader.Rows(0).Item("QryGroup45") & "','" & dtheader.Rows(0).Item("QryGroup46") & "','" & dtheader.Rows(0).Item("QryGroup47") & "','" & dtheader.Rows(0).Item("QryGroup48") & "','" & dtheader.Rows(0).Item("QryGroup49") & "','" & dtheader.Rows(0).Item("QryGroup50") & "',"
            '    strsql += vbCrLf + "'" & dtheader.Rows(0).Item("QryGroup51") & "','" & dtheader.Rows(0).Item("QryGroup52") & "','" & dtheader.Rows(0).Item("QryGroup53") & "','" & dtheader.Rows(0).Item("QryGroup54") & "','" & dtheader.Rows(0).Item("QryGroup55") & "','" & dtheader.Rows(0).Item("QryGroup56") & "','" & dtheader.Rows(0).Item("QryGroup57") & "','" & dtheader.Rows(0).Item("QryGroup58") & "','" & dtheader.Rows(0).Item("QryGroup59") & "','" & dtheader.Rows(0).Item("QryGroup60") & "',"
            '    strsql += vbCrLf + "'" & dtheader.Rows(0).Item("QryGroup61") & "','" & dtheader.Rows(0).Item("QryGroup62") & "','" & dtheader.Rows(0).Item("QryGroup63") & "','" & dtheader.Rows(0).Item("QryGroup64") & "','" & dtheader.Rows(0).Item("EvalSystem") & "','" & dtheader.Rows(0).Item("ManSerNum") & "','" & dtheader.Rows(0).Item("ManBtchNum") & "','" & dtheader.Rows(0).Item("WarrntTmpl") & "')"
            '    'da = New SqlDataAdapter(strsql, con_staging)
            '    'dt = New DataTable
            '    'da.Fill(dt)
            '    command_staging.CommandText = strsql
            '    command_staging.ExecuteNonQuery()
            '    strsql = "update oitm set U_STAGING_FLAG='Y' where itemcode='" & itemcode & "'"
            '    command_sap.CommandText = strsql
            '    command_sap.ExecuteNonQuery()
            '    'da = New SqlDataAdapter(strsql, con_SAP)
            '    'dt = New DataTable
            '    'da.Fill(dt)
            'End If
        Catch ex As Exception
            'Application.Exit()
        End Try
    End Sub

    Private Sub Add_location_Master()
        Try
            write_log(vbNewLine + "-----  Location Master Check --------")
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from s_olct with (nolock)"

            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            write_log("Location Master Header Insertion : " + strsql)

            If dt.Rows.Count > 0 Then
                ' as per the chge request , done by senthil.k on 14-08' Instead of Location name passed U_exname.
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    'strsql = "INSERT INTO S_OLCT (loc_code,loc_name,location_Address,Tin_no,Region_code,State_Code,LUPDATE,Update_flag)"
                    'strsql += vbCrLf + "SELECT T1.County,T1.Location,convert(varchar,isnull(T1.Street,'')) +'|'+ convert(varchar,isnull(T1.Building ,''))+'|'+ convert(varchar,isnull(T1.City ,'')),isnull(T1.TinNo,''),isnull(T1.block,''),(SELECT top 1 Name FROM " & MDBName & "..OCST where Code=T1.STATE),T2.U_LUPDATE,1 FROM " & MDBName & "..OLCT T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON T2.U_FORMCODE=T1.Code WHERE T2.U_FORMNAME='LOCATIONMASTER' "

                    '' old code
                    'strsql += vbCrLf + "SELECT T1.County,T1.U_exname,Replace(Replace(Replace(convert(varchar,isnull(T1.Building ,''))+'|'+ convert(varchar,isnull(T1.Street,'')) +'|'+ convert(varchar,isnull(T1.City ,''))+'|'+  convert(varchar,isnull(T1.Zipcode ,'')),'""',' '),':',' '),'&',' '),isnull(T1.TinNo,''),isnull(T1.block,''),(SELECT top 1 Name FROM " & MDBName & "..OCST where Code=T1.STATE),getdate(),1,'A' 'ADD_UPDATE',U_online FROM " & MDBName & "..OLCT T1 where T1.county is not null "
                    strsql = "INSERT INTO S_OLCT (loc_code,loc_name,location_Address,Tin_no,Region_code,State_Code,LUPDATE,Update_flag,ADD_UPDATE,ONLINE,ISWHSE,ONLINEDATE)"
                    strsql += vbCrLf + "SELECT CASE when ISNULL(T1.U_excode,'')='' then T1.County else T1.U_excode end,CASE when ISNULL(T1.U_exname,'')='' then T1.Location else T1.U_exname end,Replace(Replace(Replace(convert(varchar,isnull(T1.Building ,''))+'|'+ convert(varchar,isnull(T1.Street,'')) +'|'+ convert(varchar,isnull(T1.City ,''))+'|'+  convert(varchar,isnull(T1.Zipcode ,'')),'""',' '),':',' '),'&',' '),isnull(T1.TinNo,''),isnull(T1.block,''),(SELECT top 1 Name FROM " & MDBName & "..OCST where Code=T1.STATE),getdate(),1,'A' 'ADD_UPDATE',U_online,Case When U_Branch='Y' then '0' else '1' end,U_OnlineDate FROM " & MDBName & "..OLCT T1 where T1.county is not null "
                    'strsql += vbCrLf + "SELECT T1.County,T1.Location,Replace(Replace(Replace(convert(varchar,isnull(T1.Building ,''))+'|'+ convert(varchar,isnull(T1.Street,'')) +'|'+ convert(varchar,isnull(T1.City ,''))+'|'+  convert(varchar,isnull(T1.Zipcode ,'')),'""',' '),':',' '),'&',' '),isnull(T1.TinNo,''),isnull(T1.block,''),(SELECT top 1 Name FROM " & MDBName & "..OCST where Code=T1.STATE),getdate(),1,'A' 'ADD_UPDATE' FROM " & MDBName & "..OLCT T1 where is T1.county is not null "
                Else
                    'strsql += vbCrLf + "SELECT T1.County,T1.U_exname ,Replace(Replace(Replace(convert(varchar,isnull(T1.Building ,''))+'|'+convert(varchar,isnull(T1.Street,'')) +'|'+  convert(varchar,isnull(T1.City ,''))+'|'+  convert(varchar,isnull(T1.Zipcode ,'')),'""',' '),':',' '),'&',' '),isnull(T1.TinNo,''),isnull(T1.block,''),(SELECT top 1 Name FROM " & MDBName & "..OCST where Code=T1.STATE),T2.U_LUPDATE,1,T2.ADD_UPDATE 'add_update',U_online FROM " & MDBName & "..OLCT T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON T2.U_FORMCODE=T1.Code WHERE T2.U_FORMNAME='LOCATIONMASTER'  AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                    strsql = "INSERT INTO S_OLCT (loc_code,loc_name,location_Address,Tin_no,Region_code,State_Code,LUPDATE,Update_flag,ADD_UPDATE,ONLINE,ISWHSE,ONLINEDATE)"
                    strsql += vbCrLf + "SELECT distinct CASE when ISNULL(T1.U_excode,'')='' then T1.County else T1.U_excode end,CASE when ISNULL(T1.U_exname,'')='' then T1.Location else T1.U_exname end ,Replace(Replace(Replace(convert(varchar,isnull(T1.Building ,''))+'|'+convert(varchar,isnull(T1.Street,'')) +'|'+  convert(varchar,isnull(T1.City ,''))+'|'+  convert(varchar,isnull(T1.Zipcode ,'')),'""',' '),':',' '),'&',' '),isnull(T1.TinNo,''),isnull(T1.block,''),(SELECT top 1 Name FROM " & MDBName & "..OCST where Code=T1.STATE),T2.U_LUPDATE,1,T2.ADD_UPDATE 'add_update',U_online,Case When U_Branch='Y' then '0' else '1' end,U_OnlineDate FROM " & MDBName & "..OLCT T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON T2.U_FORMCODE=T1.Code WHERE T2.U_FORMNAME='LOCATIONMASTER'  AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                    'strsql += vbCrLf + "SELECT T1.County,T1.Location,Replace(Replace(Replace(convert(varchar,isnull(T1.Building ,''))+'|'+convert(varchar,isnull(T1.Street,'')) +'|'+  convert(varchar,isnull(T1.City ,''))+'|'+  convert(varchar,isnull(T1.Zipcode ,'')),'""',' '),':',' '),'&',' '),isnull(T1.TinNo,''),isnull(T1.block,''),(SELECT top 1 Name FROM " & MDBName & "..OCST where Code=T1.STATE),T2.U_LUPDATE,1,T2.ADD_UPDATE 'add_update' FROM " & MDBName & "..OLCT T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON T2.U_FORMCODE=T1.Code WHERE T2.U_FORMNAME='LOCATIONMASTER'  AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                End If

                write_log("Location Master Added : " + strsql)
                command_staging.CommandText = strsql
                command_staging.ExecuteNonQuery()
                'objMVasan.Register_Log(1005, "LOCATION_MASTER", dataset_log, "Success", objMVasan.Get_Next_Key_Value(1005), 0, Date.Now)

            End If
        Catch ex As Exception
            write_log("Location Master Issue : " + ex.ToString())
            'objMVasan.Register_Log(1005, "LOCATION_MASTER", dataset_log, ex.ToString, objMVasan.Get_Next_Key_Value(1005), 1, Date.Now)
        End Try
    End Sub

    Private Sub Add_location_Master_old()

        Try
            Dim dtlocation As DataTable
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from s_olct"
            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    strsql = "SELECT T1.County 'loccode',T1.Location 'locName',convert(varchar,isnull(T1.Street,'')) +'_'+ convert(varchar,isnull(T1.Building ,''))+'_'+ convert(varchar,isnull(T1.City ,'')) 'locAddress',isnull(T1.TinNo,'')'TinNo',isnull(T1.block,'') 'Region',(SELECT top 1 Name FROM OCST where Code=T1.STATE) 'state',T2.U_LUPDATE 'Date' FROM OLCT T1 JOIN FLAG_TABLE T2 ON T2.U_FORMCODE=T1.Code WHERE T2.U_FORMNAME='LOCATIONMASTER'  and T1.County is not null"
                    da = New SqlDataAdapter(strsql, con_SAP)
                    dtlocation = New DataTable
                    da.Fill(dtlocation)
                    If dt.Rows.Count > 0 Then
                        For i As Integer = 0 To dtlocation.Rows.Count - 1
                            'MsgBox(CDate(dtlocation.Rows(i).Item("Date")).ToString("yyyy/MM/dd HH:mm:ss.fff"))
                            strsql = "INSERT INTO S_OLCT (loc_code,loc_name,location_Address,Tin_no,Region_code,State_Code,lupdate,Update_flag)"
                            strsql += vbCrLf + "values ('" & dtlocation.Rows(i).Item("loccode") & "','" & dtlocation.Rows(i).Item("locName") & "','" & dtlocation.Rows(i).Item("locAddress") & "','" & dtlocation.Rows(i).Item("TinNo") & "','" & dtlocation.Rows(i).Item("Region") & "','" & dtlocation.Rows(i).Item("state") & "','" & CDate(dtlocation.Rows(i).Item("Date")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "','1') "
                            command_staging.CommandText = strsql
                            command_staging.ExecuteNonQuery()
                        Next
                    End If
                Else
                    strsql = "SELECT T1.County 'loccode',T1.Location 'locName',convert(varchar,isnull(T1.Street,'')) +'_'+ convert(varchar,isnull(T1.Building ,''))+'_'+ convert(varchar,isnull(T1.City ,'')) 'locAddress',isnull(T1.TinNo,'')'TinNo',isnull(T1.block,'') 'Region',(SELECT top 1 Name FROM OCST where Code=T1.STATE) 'state',T2.U_LUPDATE 'Date' FROM OLCT T1 JOIN FLAG_TABLE T2 ON T2.U_FORMCODE=T1.Code WHERE T2.U_FORMNAME='LOCATIONMASTER'AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                    da = New SqlDataAdapter(strsql, con_SAP)
                    dtlocation = New DataTable
                    da.Fill(dtlocation)
                    If dt.Rows.Count > 0 Then
                        For i As Integer = 0 To dtlocation.Rows.Count - 1
                            strsql = "INSERT INTO S_OLCT (loc_code,loc_name,location_Address,Tin_no,Region_code,State_Code,lupdate,Update_flag)"
                            strsql += vbCrLf + "values ('" & dtlocation.Rows(i).Item("loccode") & "','" & dtlocation.Rows(i).Item("locName") & "','" & dtlocation.Rows(i).Item("locAddress") & "','" & dtlocation.Rows(i).Item("TinNo") & "','" & dtlocation.Rows(i).Item("Region") & "','" & dtlocation.Rows(i).Item("state") & "','" & CDate(dtlocation.Rows(i).Item("Date")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "','1') "
                            command_staging.CommandText = strsql
                            command_staging.ExecuteNonQuery()
                        Next
                    End If
                    'strsql = "INSERT INTO S_OLCT (loc_code,loc_name,location_Address,Tin_no,Region_code,State_Code,Update_flag)"
                    'strsql += vbCrLf + "SELECT T1.County,T1.Location,convert(varchar,isnull(T1.Street,'')) +'_'+ convert(varchar,isnull(T1.Building ,''))+'_'+ convert(varchar,isnull(T1.City ,'')),isnull(T1.TinNo,''),isnull(T1.block,''),(SELECT top 1 Name FROM OCST where Code=T1.STATE),T2.U_LUPDATE FROM " & MDBName & "..OLCT T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON T2.U_FORMCODE=T1.Code WHERE T2.U_FORMNAME='LOCATIONMASTER'  AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                End If

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Add_salesemployee_master()
        Try
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from s_oslp"
            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    strsql = "INSERT INTO S_oslp (SLPCODE,LUPDATE)"
                    strsql += vbCrLf + "SELECT T0.SlpCode,T1.U_LUPDATE   FROM " & MDBName & "..OSLP T0  JOIN " & MDBName & "..FLAG_TABLE T1 ON T0.SlpCode=T1.U_FORMCODE WHERE T1.U_FORMNAME='SALESEMPMASTER'"
                Else
                    strsql = "INSERT INTO S_oslp (SLPCODE,LUPDATE)"
                    strsql += vbCrLf + "SELECT T0.SlpCode,T1.U_LUPDATE   FROM " & MDBName & "..OSLP T0  JOIN " & MDBName & "..FLAG_TABLE T1 ON T0.SlpCode=T1.U_FORMCODE WHERE T1.U_FORMNAME='SALESEMPMASTER' AND T1.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                End If
                command_staging.CommandText = strsql
                command_staging.ExecuteNonQuery()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Add_InventoryTransfer()
        Try
            write_log(vbNewLine + "-----  Stock Transfer Check --------")
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from s_owtr with (nolock)"
            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            write_log("StockTransfer Header Insertion : " + strsql)

            If dt.Rows.Count > 0 Then
                ' before UDF
                'strsql += vbCrLf + "SELECT T1.DocEntry,T1.DocDate,(select B.county from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode=T1.Filler )'Fromlocation',(select B.county from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode in (select WhsCode from " & MDBName & "..WTR1 where DocEntry=T1.DocEntry))'Tolocation',T1.Comments,T2.U_LUPDATE,1,(select CASE  when (B.U_branch ='Y' and B.U_online ='1' and substring(REVERSE(T1.Filler),1,1)='O' ) then 'STSO' else 'STSI' end from   " & MDBName & "..OWHS A join   " & MDBName & "..olct B on A.Location= B.Code " & _
                '                   "where WhsCode = T1.Filler) 'DocType'  FROM " & MDBName & "..OWTR  T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON "
                'strsql += vbCrLf + "T1.DocEntry=T2.U_FORMCODE WHERE T2.U_FORMNAME='INVENTRYTRANSFER' AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'  and T1.Filler =( select A.WhsCode from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode=T1.Filler and B.U_branch='Y' and B.U_online=1)"

                ' as per the chge request , done by senthil.K ( filteered by Transit whse code)
                strsql = "INSERT INTO S_owtr (DOC_NO,DOC_DATE,FROM_LOCATION ,TO_LOCATION ,REMARKS,LUPDATE,UPDATE_FLAG,DOC_TYPE,SAP_KEY,rowqty)"
                strsql += vbCrLf + "SELECT Distinct (select Seriesname from " & MDBName & "..NNM1 where objectcode='67' and Series=T1.Series) + '-' + convert(varchar(15),T1.DocNum) ,T1.DocDate,(select CASE when ISNULL(B.U_excode,'')='' AND ISNULL(B.u_exname,'')='' then B.County else B.U_excode end from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode=T1.Filler )'Fromlocation',(select CASE when ISNULL(B.U_excode,'')='' AND ISNULL(B.u_exname,'')='' then B.County else B.U_excode end from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode in (select WhsCode from " & MDBName & "..WTR1 where DocEntry=T1.DocEntry))'Tolocation',T1.Comments,T2.U_LUPDATE,1,(select CASE  when (B.U_branch ='Y' and B.U_online ='1' and substring(REVERSE(T1.Filler),1,1)='O' ) then 'STSO' else 'STSI' end from   " & MDBName & "..OWHS A join   " & MDBName & "..olct B on A.Location= B.Code " & _
                                    "where WhsCode = T1.Filler) 'DocType',T1.DocEntry,(Select Sum(Quantity) from " & MDBName & "..wtr1 where DocEntry=t1.DocEntry)  FROM " & MDBName & "..OWTR  T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON "
                strsql += vbCrLf + "T1.DocEntry=T2.U_FORMCODE WHERE T2.U_FORMNAME='INVENTRYTRANSFER' AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "' and T1.DocEntry Not in(Select  Isnull(SAP_KEY,0) from S_owtr)"

                'strsql += vbCrLf + "INSERT INTO S_wtr1 (DOC_NO,PRODUCT_CODE,BARCODE,QTY,RATE,LOCATION ,MRP,VENDOR_CODE,REMARKS,UOM,LENS_TYPE,LUPDATE,SAP_KEY)"
                'strsql += vbCrLf + "SELECT Distinct AA.DocNum,AA.ItemCode,AA.IntrSerial,AA.QUANTITY,AA.StockPrice,(select Top 1 CASE when ISNULL(B.U_excode,'')='' AND ISNULL(B.u_exname,'')='' then B.County else B.U_excode end from  " & MDBName & "..olct B where B.Code=AA.LocCode )'LOcation',AA.MRP,AA.[Vendor Code]  'VendorCode',AA.Remarks,isnull(UOM,'Nos'),'' 'Lens Type',BB.U_LUPDATE,AA.DocEntry "
                'strsql += vbCrLf + "FROM (SELECT DISTINCT (select Seriesname + '-' + convert(varchar(10),W1.DocNum) from " & MDBName & "..NNM1 as N1 join " & MDBName & "..OWTR as W1 on N1.Series=W1.Series  where N1.objectcode='67' and W1.DocEntry=T1.DocEntry  ) as 'DocNum',T1.DocEntry, T1.LineNum, T1.ItemCode as ItemCode,T1.Dscription,1 QUANTITY,T1.WhsCode,T1.LocCode,T1.StockPrice,T3.IntrSerial,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,isnull(T3.U_MRP,0) 'MRP',isnull(T3.CardCode,t3.U_vendor) 'Vendor Code',1 'QUANTITY1','' 'Remarks',T6.InvntryUom as 'UOM' FROM"
                'strsql += vbCrLf + "" & MDBName & "..WTR1 T1 JOIN " & MDBName & "..SRI1 T2 ON T1.DocEntry=T2.BaseEntry AND T1.LineNum=T2.BaseLinNum"
                'strsql += vbCrLf + "JOIN " & MDBName & "..OSRI T3 ON T3.SysSerial=T2.SysSerial AND T3.ItemCode=T2.ItemCode "
                'strsql += vbCrLf + "JOIN " & MDBName & "..OSRN T4 ON T4.ItemCode=T3.ItemCode AND T4.DistNumber =T3.IntrSerial Left Outer Join " & MDBName & "..OITM as T6 on T1.ItemCode = T6.ItemCode "
                'strsql += vbCrLf + "join " & MDBName & "..FLAG_TABLE BB on T1.DocEntry=bb.U_FORMCODE  where bb.U_FORMNAME='INVENTRYTRANSFER'"
                'strsql += vbCrLf + " and T2.BaseType='67' and (substring(REVERSE(T1.WhsCode),1,1)='O' OR substring(reverse(T1.whscode),1,1)='T' or substring(reverse(T1.whscode),1,1)='A') AND bb.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                'strsql += vbCrLf + "UNION ALL"
                'strsql += vbCrLf + "SELECT DISTINCT  (select Seriesname + '-' + convert(varchar(10),W1.DocNum) from " & MDBName & "..NNM1 as N1 join " & MDBName & "..OWTR as W1 on N1.Series=W1.Series  where N1.objectcode='67' and W1.DocEntry=T3.DocEntry  ) as 'DocNum',T3.DocEntry,T3.LineNum,T3.ItemCode,T3.Dscription,T2.Quantity,T3.WhsCode,T3.LocCode,T3.StockPrice,T2.BatchNum,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,isnull(T4.U_mrp,0) 'MRP',isnull((select top 1 cardcode from " & MDBName & "..ibt1 where BatchNum=T2.BatchNum  and BaseType='20'),(select top 1 U_vendor  from " & MDBName & "..oibt where DistNumber=T4.DistNumber  and BaseType='20' and U_vendor is not null)) 'Vendor Code',T2.Quantity 'QUANTITY1','' 'Remarks',(Select top 1 Invntryuom from " & MDBName & "..OITM where ItemCode =T3.ItemCode)"
                'strsql += vbCrLf + "FROM " & MDBName & "..IBT1 T2 "
                'strsql += vbCrLf + "JOIN " & MDBName & "..WTR1 T3 ON T3.DocEntry=T2.BaseEntry AND T3.LineNum=T2.BaseLinNum "
                'strsql += vbCrLf + "JOIN " & MDBName & "..OBTN T4 ON T4.DistNumber=T2.BatchNum and T4.ItemCode=T2.ItemCode"
                'strsql += vbCrLf + "join " & MDBName & "..FLAG_TABLE BB on T3.DocEntry=bb.U_FORMCODE  where bb.U_FORMNAME='INVENTRYTRANSFER'"
                'strsql += vbCrLf + "and T2.BaseType='67' AND bb.U_LUPDATE>'" & CDate& MDBName &(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                'strsql += vbCrLf + "Union ALl"
                'strsql += vbCrLf + "select (select Seriesname + '-' + convert(varchar(10),W1.DocNum) from " & MDBName & "..NNM1 as N1 join " & MDBName & "..OWTR as W1 on N1.Series=W1.Series  where N1.objectcode='67' and W1.DocEntry=T1.DocEntry  ) as 'DocNum',T1.DocEntry,T1.LineNum, T1.ItemCode as ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.loccode,T1.stockprice,'' 'Intrserial','' 'Mnfserial','' 'LotNumber','' 'ExpDate','' 'MnfDate','' 'Indate',"
                'strsql += vbCrLf + "0 'MRP','' 'Vendor Code',0 'Quantity1','' 'Remarks',T2.InvntryUom from " & MDBName & "..wtr1 T1 join " & MDBName & "..OITM T2 on T2.ItemCode=T1.ItemCode join " & MDBName & "..FLAG_TABLE BB on T1.DocEntry=bb.U_FORMCODE where T2.ManSerNum='N' and T2.ManBtchNum='N'  and bb.U_FORMNAME='INVENTRYTRANSFER' AND bb.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                'strsql += vbCrLf + ")AA join " & MDBName & "..FLAG_TABLE BB on aa.DocEntry=bb.U_FORMCODE  where bb.U_FORMNAME='INVENTRYTRANSFER' AND bb.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "' and aa.DocEntry Not in(Select Isnull(SAP_KEY,0) from S_wtr1)"

                strsql += vbCrLf + "INSERT INTO"
                strsql += vbCrLf + "S_wtr1(DOC_NO, PRODUCT_CODE, BARCODE, QTY, Rate, LOCATION, MRP, VENDOR_CODE, REMARKS, UOM, LENS_TYPE, LUPDATE, SAP_KEY)"
                strsql += vbCrLf + "SELECT DISTINCT AA.docnum,"
                strsql += vbCrLf + "AA.itemcode,"
                strsql += vbCrLf + "AA.intrserial,"
                strsql += vbCrLf + "AA.quantity,"
                strsql += vbCrLf + "AA.stockprice,"
                strsql += vbCrLf + "(SELECT TOP 1 CASE"
                strsql += vbCrLf + "WHEN Isnull(B.u_excode, '') = ''"
                strsql += vbCrLf + "AND Isnull(B.u_exname, '') = '' THEN"
                strsql += vbCrLf + "B.county"
                strsql += vbCrLf + "Else  B.u_excode"
                strsql += vbCrLf + "End"
                strsql += vbCrLf + "FROM   " & MDBName & "..olct B"
                strsql += vbCrLf + "WHERE  B.code = AA.loccode) 'LOcation',"
                strsql += vbCrLf + "AA.mrp,"
                strsql += vbCrLf + "AA.[vendor code]             'VendorCode',"
                strsql += vbCrLf + "AA.remarks,"
                strsql += vbCrLf + "Isnull(invntryuom, 'Nos')    'InvntryUom',"
                strsql += vbCrLf + "''                           'Lens Type',"
                strsql += vbCrLf + "'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'         U_LUPDATE,"
                strsql += vbCrLf + "AA.docentry"
                strsql += vbCrLf + "--into StckTrns"
                strsql += vbCrLf + "FROM   (SELECT DISTINCT (SELECT seriesname + '-'"
                strsql += vbCrLf + "+ CONVERT(VARCHAR(10), W1.docnum)"
                strsql += vbCrLf + "FROM   " & MDBName & "..nnm1 AS N1"
                strsql += vbCrLf + "JOIN " & MDBName & "..owtr AS W1"
                strsql += vbCrLf + "ON N1.series = W1.series"
                strsql += vbCrLf + "WHERE  N1.objectcode = '67'"
                strsql += vbCrLf + "AND W1.docentry = T1.docentry) AS 'DocNum',"
                strsql += vbCrLf + "T1.docentry,"
                strsql += vbCrLf + "T1.linenum,"
                strsql += vbCrLf + "T1.itemcode                            AS ItemCode,"
                strsql += vbCrLf + "T1.dscription,"
                strsql += vbCrLf + "1              QUANTITY,"
                strsql += vbCrLf + "T1.whscode,"
                strsql += vbCrLf + "T1.loccode,"
                strsql += vbCrLf + "T1.stockprice,"
                strsql += vbCrLf + "T3.intrserial,"
                strsql += vbCrLf + "T4.mnfserial,"
                strsql += vbCrLf + "T4.lotnumber,"
                strsql += vbCrLf + "T4.expdate,"
                strsql += vbCrLf + "T4.mnfdate,"
                strsql += vbCrLf + "T4.indate,"
                strsql += vbCrLf + "Isnull(T3.u_mrp, 0)                    'MRP',"
                strsql += vbCrLf + "Isnull(T3.cardcode, t3.u_vendor)       'Vendor Code',"
                strsql += vbCrLf + "1              'QUANTITY1',"
                strsql += vbCrLf + "''                                     'Remarks',"
                strsql += vbCrLf + "T6.invntryuom                          AS 'InvntryUom'"
                strsql += vbCrLf + "FROM   " & MDBName & "..wtr1 T1"
                strsql += vbCrLf + "JOIN " & MDBName & "..sri1 T2"
                strsql += vbCrLf + "ON T1.docentry = T2.baseentry"
                strsql += vbCrLf + "AND T1.linenum = T2.baselinnum"
                strsql += vbCrLf + "JOIN " & MDBName & "..osri T3"
                strsql += vbCrLf + "ON T3.sysserial = T2.sysserial"
                strsql += vbCrLf + "AND T3.itemcode = T2.itemcode"
                strsql += vbCrLf + "JOIN " & MDBName & "..osrn T4"
                strsql += vbCrLf + "  ON T4.itemcode = T3.itemcode"
                strsql += vbCrLf + "AND T4.distnumber = T3.intrserial"
                strsql += vbCrLf + "LEFT OUTER JOIN " & MDBName & "..oitm AS T6"
                strsql += vbCrLf + "ON T1.itemcode = T6.itemcode"
                strsql += vbCrLf + "--join"
                strsql += vbCrLf + "--   " & MDBName & "..FLAG_TABLE BB"
                strsql += vbCrLf + "--   on T1.DocEntry = bb.U_FORMCODE"
                strsql += vbCrLf + "WHERE"
                strsql += vbCrLf + "--bb.U_FORMNAME = 'INVENTRYTRANSFER'"
                strsql += vbCrLf + "T1.docentry IN (SELECT u_formcode"
                strsql += vbCrLf + "FROM   " & MDBName & "..flag_table"
                strsql += vbCrLf + "WHERE  u_formname = 'INVENTRYTRANSFER'"
                strsql += vbCrLf + "AND u_lupdate > '2020-12-21 06:34:56.237')"
                strsql += vbCrLf + "AND T2.basetype = '67'"
                strsql += vbCrLf + "AND ( Substring(Reverse(T1.whscode), 1, 1) = 'O'"
                strsql += vbCrLf + "OR Substring(Reverse(T1.whscode), 1, 1) = 'T'"
                strsql += vbCrLf + "OR Substring(Reverse(T1.whscode), 1, 1) = 'A' )"
                strsql += vbCrLf + "--AND bb.U_LUPDATE > '" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                strsql += vbCrLf + "UNION ALL"
                strsql += vbCrLf + "SELECT DISTINCT (SELECT seriesname + '-'"
                strsql += vbCrLf + "+ CONVERT(VARCHAR(10), W1.docnum)"
                strsql += vbCrLf + "FROM   " & MDBName & "..nnm1 AS N1"
                strsql += vbCrLf + "JOIN " & MDBName & "..owtr AS W1"
                strsql += vbCrLf + "ON N1.series = W1.series"
                strsql += vbCrLf + "WHERE  N1.objectcode = '67'"
                strsql += vbCrLf + "AND W1.docentry = T3.docentry) AS 'DocNum',"
                strsql += vbCrLf + "T3.docentry,"
                strsql += vbCrLf + "T3.linenum,"
                strsql += vbCrLf + "T3.itemcode,"
                strsql += vbCrLf + "T3.dscription,"
                strsql += vbCrLf + "T2.quantity,"
                strsql += vbCrLf + "T3.whscode,"
                strsql += vbCrLf + "T3.loccode,"
                strsql += vbCrLf + "T3.stockprice,"
                strsql += vbCrLf + "T2.batchnum,"
                strsql += vbCrLf + "T4.mnfserial,"
                strsql += vbCrLf + "T4.lotnumber,"
                strsql += vbCrLf + "T4.expdate,"
                strsql += vbCrLf + "T4.mnfdate,"
                strsql += vbCrLf + "T4.indate,"
                strsql += vbCrLf + "Isnull(T4.u_mrp, 0)                    'MRP',"
                strsql += vbCrLf + "("
                strsql += vbCrLf + "select top 1"
                strsql += vbCrLf + "s0.CardCode"
                strsql += vbCrLf + "from"
                strsql += vbCrLf + "" & MDBName & "..OITL S0"
                strsql += vbCrLf + "inner Join"
                strsql += vbCrLf + "" & MDBName & "..ITL1 S1"
                strsql += vbCrLf + "on S0.LogEntry = S1.LogEntry"
                strsql += vbCrLf + "--inner join " & MDBName & "..obtn b on"
                strsql += vbCrLf + "and s1.SysNumber=t4.SysNumber               "
                strsql += vbCrLf + "where"
                strsql += vbCrLf + "--b.DistNumber= t2.BatchNum  and"
                strsql += vbCrLf + "s0.DocType ='20' and s0.ItemCode =t3.ItemCode"
                strsql += vbCrLf + ")"
                strsql += vbCrLf + "'Vendor Code',"
                strsql += vbCrLf + "T2.quantity                            'QUANTITY1',"
                strsql += vbCrLf + "''                                     'Remarks',"
                strsql += vbCrLf + "(SELECT TOP 1 invntryuom"
                strsql += vbCrLf + "FROM   " & MDBName & "..oitm"
                strsql += vbCrLf + "WHERE  itemcode = T3.itemcode)        'InvntryUom'"
                strsql += vbCrLf + "FROM   " & MDBName & "..ibt1 T2"
                strsql += vbCrLf + "JOIN " & MDBName & "..wtr1 T3"
                strsql += vbCrLf + "ON T3.docentry = T2.baseentry"
                strsql += vbCrLf + "AND T3.linenum = T2.baselinnum"
                strsql += vbCrLf + "JOIN " & MDBName & "..obtn T4"
                strsql += vbCrLf + "ON T4.distnumber = T2.batchnum"
                strsql += vbCrLf + "AND T4.itemcode = T2.itemcode"
                strsql += vbCrLf + "--join"
                strsql += vbCrLf + "--   " & MDBName & "..FLAG_TABLE BB"
                strsql += vbCrLf + "--   on T3.DocEntry = bb.U_FORMCODE"
                strsql += vbCrLf + "WHERE"
                strsql += vbCrLf + "--bb.U_FORMNAME = 'INVENTRYTRANSFER'"
                strsql += vbCrLf + "t3.docentry IN (SELECT u_formcode"
                strsql += vbCrLf + "FROM   " & MDBName & "..flag_table"
                strsql += vbCrLf + "WHERE  u_formname = 'INVENTRYTRANSFER'"
                strsql += vbCrLf + "AND u_lupdate > '2020-12-21 06:34:56.237')"
                strsql += vbCrLf + "AND T2.basetype = '67'"
                strsql += vbCrLf + "--AND bb.U_LUPDATE > '" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                strsql += vbCrLf + "UNION ALL"
                strsql += vbCrLf + "SELECT (SELECT seriesname + '-'"
                strsql += vbCrLf + "+ CONVERT(VARCHAR(10), W1.docnum)"
                strsql += vbCrLf + "FROM   " & MDBName & "..nnm1 AS N1"
                strsql += vbCrLf + "JOIN " & MDBName & "..owtr AS W1"
                strsql += vbCrLf + "ON N1.series = W1.series"
                strsql += vbCrLf + "WHERE  N1.objectcode = '67'"
                strsql += vbCrLf + "AND W1.docentry = T1.docentry) AS 'DocNum',"
                strsql += vbCrLf + "T1.docentry,"
                strsql += vbCrLf + "T1.linenum,"
                strsql += vbCrLf + "T1.itemcode                            AS ItemCode,"
                strsql += vbCrLf + "T1.dscription,"
                strsql += vbCrLf + "T1.quantity,"
                strsql += vbCrLf + "T1.whscode,"
                strsql += vbCrLf + "T1.loccode,"
                strsql += vbCrLf + "T1.stockprice,"
                strsql += vbCrLf + "''                                     'Intrserial',"
                strsql += vbCrLf + "''                                     'Mnfserial',"
                strsql += vbCrLf + "''                                     'LotNumber',"
                strsql += vbCrLf + "''                                     'ExpDate',"
                strsql += vbCrLf + "''                                     'MnfDate',"
                strsql += vbCrLf + "''                                     'Indate',"
                strsql += vbCrLf + "0              'MRP',"
                strsql += vbCrLf + "''                                     'Vendor Code',"
                strsql += vbCrLf + "0              'Quantity1',"
                strsql += vbCrLf + "''                                     'Remarks',"
                strsql += vbCrLf + "T2.invntryuom"
                strsql += vbCrLf + "FROM   " & MDBName & "..wtr1 T1"
                strsql += vbCrLf + "JOIN " & MDBName & "..oitm T2"
                strsql += vbCrLf + "ON T2.itemcode = T1.itemcode"
                strsql += vbCrLf + "WHERE  T2.mansernum = 'N'"
                strsql += vbCrLf + "AND T2.manbtchnum = 'N'"
                strsql += vbCrLf + "AND T1.docentry IN (SELECT u_formcode"
                strsql += vbCrLf + "                                   FROM   " & MDBName & "..flag_table"
                strsql += vbCrLf + "WHERE  u_formname = 'INVENTRYTRANSFER'"
                strsql += vbCrLf + "AND u_lupdate >'2020-12-21 06:34:56.237'))"
                strsql += vbCrLf + "AA"
                strsql += vbCrLf + "WHERE  aa.docentry IN (SELECT u_formcode"
                strsql += vbCrLf + "FROM   " & MDBName & "..flag_table"
                strsql += vbCrLf + "WHERE  u_formname = 'INVENTRYTRANSFER'"
                strsql += vbCrLf + "AND u_lupdate > '2020-12-21 06:34:56.237')"
                strsql += vbCrLf + "AND aa.docentry NOT IN (SELECT Isnull(sap_key, 0)"
                strsql += vbCrLf + "FROM   s_wtr1)"


                'tran = con_staging.BeginTransaction("INVENTOEY")
                'command_staging.CommandText = strsql
                'command_staging.CommandTimeout = 1000
                'command_staging.Transaction = tran
                'command_staging.ExecuteNonQuery()
                'strsql = ""

                Try
                    '' as per the change request , done by senthil.k ( for STSI , STSO update)

                    'strsql = "INSERT INTO S_owtr (DOC_NO,DOC_DATE,FROM_LOCATION ,TO_LOCATION ,REMARKS,LUPDATE,UPDATE_FLAG,DOC_TYPE,SAP_KEY,rowqty)"
                    'strsql += vbCrLf + "SELECT Distinct (select Seriesname from " & MDBName & "..NNM1 where objectcode='67' and Series=T1.Series) + '-' + convert(varchar(15),T1.DocNum),T1.DocDate,(select CASE when ISNULL(B.U_excode,'')='' AND ISNULL(B.u_exname,'')='' then B.County else B.U_excode end from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode=T1.Filler  )'Fromlocation',(select CASE when ISNULL(B.U_excode,'')='' AND ISNULL(B.u_exname,'')='' then B.County else B.U_excode end from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode in (select WhsCode from " & MDBName & "..WTR1 where DocEntry=T1.DocEntry))'Tolocation',T1.Comments,T2.U_LUPDATE,1,(select CASE  when (B.U_branch ='Y' and B.U_online ='1' and substring(REVERSE(T1.Filler),1,1)='O' ) then 'STSO' else 'STSI' end from   " & MDBName & "..OWHS A join   " & MDBName & "..olct B on A.Location= B.Code " & _
                    '                    " where WhsCode = T1.Filler) 'DocType',T1.DocEntry,(Select Sum(Quantity) from " & MDBName & "..wtr1 where DocEntry=t1.DocEntry)   FROM " & MDBName & "..OWTR  T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON "
                    'strsql += vbCrLf + "T1.DocEntry=T2.U_FORMCODE WHERE T2.U_FORMNAME='INVENTRYTRANSFER' AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'  and  T1.DocEntry Not in(Select  Isnull(SAP_KEY,0) from S_owtr)"
                    'WriteLog_INV(strsql)
                    'command_staging.CommandText = strsql
                    'command_staging.ExecuteNonQuery()

                Catch ex As Exception

                End Try

            End If
            'tran.Commit()
            'command_staging.CommandText = strsql
            'command_staging.ExecuteNonQuery()
            'End If
            write_log("Stock Transfer Added : " + strsql)
            command_staging.CommandText = strsql
            command_staging.CommandTimeout = 600
            command_staging.ExecuteNonQuery()
            'objMVasan.Register_Log(1005, "STOCK_OUT", dataset_log, "Success", objMVasan.Get_Next_Key_Value(1005), 1, Date.Now)

        Catch ex As Exception
            'tran.Rollback()
            write_log("Stock Transfer Issue : " + ex.ToString())
            'objMVasan.Register_Log(1005, "STOCK_OUT", dataset_log, ex.ToString, objMVasan.Get_Next_Key_Value(1005), 1, Date.Now)
        End Try
        'If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
        '    'strsql = "INSERT INTO S_owtr (DOC_NO,DOC_DATE,FROM_LOCATION ,TO_LOCATION ,REMARKS,LUPDATE,UPDATE_FLAG)"
        '    'strsql += vbCrLf + "SELECT T1.DocEntry,T1.DocDate,(select B.county from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode=T1.Filler)'Fromlocation',(select B.county from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode=T1.Filler)'tolocation',T1.Comments,T2.U_LUPDATE,1  FROM " & MDBName & "..OWTR  T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON "
        '    'strsql += vbCrLf + "T1.DocEntry=T2.U_FORMCODE WHERE T2.U_FORMNAME='INVENTRYTRANSFER'"
        '    strsql = "INSERT INTO S_owtr (DOC_NO,DOC_DATE,FROM_LOCATION ,TO_LOCATION ,REMARKS,LUPDATE,UPDATE_FLAG)"
        '    strsql += vbCrLf + "SELECT T1.DocEntry,T1.DocDate,(select B.county from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode=T1.Filler)'Fromlocation',(select B.county from " & MDBName & "..OWHS A join " & MDBName & "..olct B on A.Location=B.Code  where WhsCode=T1.Filler)'tolocation',T1.Comments,getdate(),1  FROM " & MDBName & "..OWTR  T1"

        '    'strsql += vbCrLf + "INSERT INTO S_wtr1 (DOC_NO,PRODUCT_CODE,BARCODE,QTY,RATE,LOCATION ,MRP,VENDOR_CODE,REMARKS,UOM,LENS_TYPE,LUPDATE)"
        '    'strsql += vbCrLf + "SELECT AA.DocEntry,AA.ItemCode,AA.IntrSerial,AA.QUANTITY,AA.StockPrice,(select top 1 county  from " & MDBName & "..olct where code=AA.LocCode)'Location',AA.MRP,'' 'VendorCode',AA.Remarks,(select top 1 InvntryUom from " & MDBName & "..OITM where ItemCode=AA.ItemCode) 'UOM','' 'Lens Type',BB.U_LUPDATE "
        '    'strsql += vbCrLf + "FROM (SELECT DISTINCT T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.LocCode,T1.StockPrice,T3.IntrSerial,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',1 'QUANTITY1','' 'Remarks' FROM"
        '    'strsql += vbCrLf + "" & MDBName & "..WTR1 T1 JOIN " & MDBName & "..SRI1 T2 ON T1.DocEntry=T2.BaseEntry AND T1.LineNum=T2.BaseLinNum"
        '    'strsql += vbCrLf + "JOIN " & MDBName & "..OSRI T3 ON T3.SysSerial=T2.SysSerial AND T3.ItemCode=T2.ItemCode "
        '    'strsql += vbCrLf + "JOIN " & MDBName & "..OSRN T4 ON T4.ItemCode=T3.ItemCode AND T4.DistNumber =T3.IntrSerial "
        '    'strsql += vbCrLf + " WHERE T2.BaseType='67' "
        '    'strsql += vbCrLf + "UNION ALL"
        '    'strsql += vbCrLf + "SELECT DISTINCT  T3.DocEntry,T3.LineNum,T3.ItemCode,T3.Dscription,T3.Quantity,T3.WhsCode,T3.LocCode,T3.StockPrice,T2.BatchNum,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',T2.Quantity 'QUANTITY1','' 'Remarks'   "
        '    'strsql += vbCrLf + "FROM " & MDBName & "..IBT1 T2 "
        '    'strsql += vbCrLf + "JOIN " & MDBName & "..WTR1 T3 ON T3.DocEntry=T2.BaseEntry AND T3.LineNum=T2.BaseLinNum "
        '    'strsql += vbCrLf + "JOIN " & MDBName & "..OBTN T4 ON T4.DistNumber=T2.BatchNum "
        '    'strsql += vbCrLf + "WHERE T2.BaseType='67'"
        '    'strsql += vbCrLf + "Union All"
        '    'strsql += vbCrLf + "select T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.loccode,T1.stockprice,'' 'Intrserial','' 'Mnfserial','' 'LotNumber','' 'ExpDate','' 'MnfDate','' 'Indate',"
        '    'strsql += vbCrLf + "0 'MRP',0 'Quantity1','' 'Remarks' from " & MDBName & "..wtr1 T1 join " & MDBName & "..OITM T2 on T2.ItemCode=T1.ItemCode where T2.ManSerNum='N' and T2.ManBtchNum='N'"
        '    'strsql += vbCrLf + ")AA join " & MDBName & "..FLAG_TABLE BB on aa.DocEntry=bb.U_FORMCODE  where bb.U_FORMNAME='INVENTRYTRANSFER'"
        '    strsql += vbCrLf + "INSERT INTO S_wtr1 (DOC_NO,PRODUCT_CODE,BARCODE,QTY,RATE,LOCATION ,MRP,VENDOR_CODE,REMARKS,UOM,LENS_TYPE,LUPDATE,UPDATE_FLAG)"
        '    strsql += vbCrLf + "SELECT AA.DocEntry,AA.ItemCode,AA.IntrSerial,AA.QUANTITY,AA.StockPrice,(select top 1 county  from " & MDBName & "..olct where code=AA.LocCode)'Location',AA.MRP,AA.[Vendor Code]  'VendorCode',AA.Remarks,(select top 1 InvntryUom from " & MDBName & "..OITM where ItemCode=AA.ItemCode) 'UOM','' 'Lens Type',getdate(),1 "
        '    strsql += vbCrLf + "FROM (SELECT DISTINCT T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.LocCode,T1.StockPrice,T3.IntrSerial,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,isnull(T3.U_MRP,0) 'MRP',T3.CardCode 'Vendor Code',1 'QUANTITY1','' 'Remarks' FROM"
        '    strsql += vbCrLf + "" & MDBName & "..WTR1 T1 JOIN " & MDBName & "..SRI1 T2 ON T1.DocEntry=T2.BaseEntry AND T1.LineNum=T2.BaseLinNum"
        '    strsql += vbCrLf + "JOIN " & MDBName & "..OSRI T3 ON T3.SysSerial=T2.SysSerial AND T3.ItemCode=T2.ItemCode "
        '    strsql += vbCrLf + "JOIN " & MDBName & "..OSRN T4 ON T4.ItemCode=T3.ItemCode AND T4.DistNumber =T3.IntrSerial "
        '    strsql += vbCrLf + " WHERE T2.BaseType='67' "
        '    strsql += vbCrLf + "UNION ALL"
        '    strsql += vbCrLf + "SELECT DISTINCT  T3.DocEntry,T3.LineNum,T3.ItemCode,T3.Dscription,T3.Quantity,T3.WhsCode,T3.LocCode,T3.StockPrice,T2.BatchNum,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,isnull(T4.U_mrp,0) 'MRP',(select top 1 cardcode from ibt1 where BatchNum=T2.BatchNum  and BaseType='18') 'Vendor Code',T2.Quantity 'QUANTITY1','' 'Remarks'   "
        '    strsql += vbCrLf + "FROM " & MDBName & "..IBT1 T2 "
        '    strsql += vbCrLf + "JOIN " & MDBName & "..WTR1 T3 ON T3.DocEntry=T2.BaseEntry AND T3.LineNum=T2.BaseLinNum "
        '    strsql += vbCrLf + "JOIN " & MDBName & "..OBTN T4 ON T4.DistNumber=T2.BatchNum "
        '    strsql += vbCrLf + "WHERE T2.BaseType='67'"
        '    strsql += vbCrLf + "Union ALl"
        '    strsql += vbCrLf + "select T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.loccode,T1.stockprice,'' 'Intrserial','' 'Mnfserial','' 'LotNumber','' 'ExpDate','' 'MnfDate','' 'Indate',"
        '    strsql += vbCrLf + "0 'MRP','' 'Vendor Code',0 'Quantity1','' 'Remarks' from " & MDBName & "..wtr1 T1 join " & MDBName & "..OITM T2 on T2.ItemCode=T1.ItemCode where T2.ManSerNum='N' and T2.ManBtchNum='N'"
        '    strsql += vbCrLf + ")AA"
        'Else
    End Sub

    Private Sub Add_InventoryTransfer_old()
        Try
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from s_owtr"
            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    strsql = "INSERT INTO S_owtr (DOCNUM,SERIES,DOCDATE,DOCCUR,COMMENTS,GROUPNUM,SLPCODE,FILLER,LUPDATE)"
                    strsql += vbCrLf + "SELECT T1.DocEntry,T1.Series,T1.DocDate,T1.DocCur,T1.Comments,T1.GroupNum,T1.SlpCode,T1.Filler,T2.U_LUPDATE  FROM " & MDBName & "..OWTR  T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON "
                    strsql += vbCrLf + "T1.DocEntry=T2.U_FORMCODE WHERE T2.U_FORMNAME='INVENTRYTRANSFER'"

                    strsql += vbCrLf + "INSERT INTO S_wtr1 (DOCNUM,LINENUM,ITEMCODE,DSCRIPTION,QUANTITY,WHSCODE,LOCCODE,STOCKPRICE,INTRSERIAL,LOTNUMBER,MNFSERIAL,EXPDATE,MNFDATE,INDATE,MRP,QUANTITY1)"
                    strsql += vbCrLf + "SELECT AA.DocEntry,AA.LineNum,AA.ItemCode,AA.Dscription,AA.QUANTITY,AA.WhsCode,AA.LocCode,AA.StockPrice,AA.IntrSerial,AA.MnfSerial,AA.LotNumber,AA.ExpDate,AA.MnfDate ,AA.InDate ,AA.MRP,AA.QUANTITY1 "
                    strsql += vbCrLf + "FROM (SELECT DISTINCT T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.LocCode,T1.StockPrice,T3.IntrSerial,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',1 'QUANTITY1' FROM"
                    strsql += vbCrLf + "" & MDBName & "..WTR1 T1 JOIN " & MDBName & "..SRI1 T2 ON T1.DocEntry=T2.BaseEntry AND T1.LineNum=T2.BaseLinNum"
                    strsql += vbCrLf + "JOIN " & MDBName & "..OSRI T3 ON T3.SysSerial=T2.SysSerial AND T3.ItemCode=T2.ItemCode "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OSRN T4 ON T4.ItemCode=T3.ItemCode AND T4.DistNumber =T3.IntrSerial "
                    strsql += vbCrLf + " WHERE T2.BaseType='67' "
                    strsql += vbCrLf + "UNION ALL"
                    strsql += vbCrLf + "SELECT DISTINCT  T3.DocEntry,T3.LineNum,T3.ItemCode,T3.Dscription,T2.Quantity,T3.WhsCode,T3.LocCode,T3.StockPrice,T2.BatchNum,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',T2.Quantity 'QUANTITY1' "
                    strsql += vbCrLf + "FROM " & MDBName & "..IBT1 T2 "
                    strsql += vbCrLf + "JOIN " & MDBName & "..WTR1 T3 ON T3.DocEntry=T2.BaseEntry AND T3.LineNum=T2.BaseLinNum "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OBTN T4 ON T4.DistNumber=T2.BatchNum "
                    strsql += vbCrLf + "WHERE T2.BaseType='67'"
                    strsql += vbCrLf + "Union ALl"
                    strsql += vbCrLf + "select T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.loccode,T1.stockprice,'' 'Intrserial','' 'Mnfserial','' 'LotNumber','' 'ExpDate','' 'MnfDate','' 'Indate',"
                    strsql += vbCrLf + "0 'MRP',0 'Quantity1' from " & MDBName & "..wtr1 T1 join " & MDBName & "..OITM T2 on T2.ItemCode=T1.ItemCode where T2.ManSerNum='N' and T2.ManBtchNum='N'"
                    strsql += vbCrLf + ")AA join " & MDBName & "..FLAG_TABLE BB on aa.DocEntry=bb.U_FORMCODE  where bb.U_FORMNAME='INVENTRYTRANSFER'"
                Else
                    strsql = "INSERT INTO S_owtr (DOCNUM,SERIES,DOCDATE,DOCCUR,COMMENTS,GROUPNUM,SLPCODE,FILLER,LUPDATE)"
                    strsql += vbCrLf + "SELECT T1.DocEntry,T1.Series,T1.DocDate,T1.DocCur,T1.Comments,T1.GroupNum,T1.SlpCode,T1.Filler,T2.U_LUPDATE  FROM " & MDBName & "..OWTR  T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON "
                    strsql += vbCrLf + "T1.DocEntry=T2.U_FORMCODE WHERE T2.U_FORMNAME='INVENTRYTRANSFER' AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"

                    strsql += vbCrLf + "INSERT INTO S_wtr1 (DOCNUM,LINENUM,ITEMCODE,DSCRIPTION,QUANTITY,WHSCODE,LOCCODE,STOCKPRICE,INTRSERIAL,LOTNUMBER,MNFSERIAL,EXPDATE,MNFDATE,INDATE,MRP,QUANTITY1)"
                    strsql += vbCrLf + "SELECT AA.DocEntry,AA.LineNum,AA.ItemCode,AA.Dscription,AA.QUANTITY,AA.WhsCode,AA.LocCode,AA.StockPrice,AA.IntrSerial,AA.MnfSerial,AA.LotNumber,AA.ExpDate,AA.MnfDate ,AA.InDate ,AA.MRP,AA.QUANTITY1 "
                    strsql += vbCrLf + "FROM (SELECT DISTINCT T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.LocCode,T1.StockPrice,T3.IntrSerial,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',1 'QUANTITY1' FROM"
                    strsql += vbCrLf + "" & MDBName & "..WTR1 T1 JOIN " & MDBName & "..SRI1 T2 ON T1.DocEntry=T2.BaseEntry AND T1.LineNum=T2.BaseLinNum"
                    strsql += vbCrLf + "JOIN " & MDBName & "..OSRI T3 ON T3.SysSerial=T2.SysSerial AND T3.ItemCode=T2.ItemCode "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OSRN T4 ON T4.ItemCode=T3.ItemCode AND T4.DistNumber =T3.IntrSerial "
                    strsql += vbCrLf + " WHERE T2.BaseType='67' "
                    strsql += vbCrLf + "UNION ALL"
                    strsql += vbCrLf + "SELECT DISTINCT  T3.DocEntry,T3.LineNum,T3.ItemCode,T3.Dscription,T3.Quantity,T3.WhsCode,T3.LocCode,T3.StockPrice,T2.BatchNum,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',T2.Quantity 'QUANTITY1'   "
                    strsql += vbCrLf + "FROM " & MDBName & "..IBT1 T2 "
                    strsql += vbCrLf + "JOIN " & MDBName & "..WTR1 T3 ON T3.DocEntry=T2.BaseEntry AND T3.LineNum=T2.BaseLinNum "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OBTN T4 ON T4.DistNumber=T2.BatchNum "
                    strsql += vbCrLf + "WHERE T2.BaseType='67'"
                    strsql += vbCrLf + "Union ALl"
                    strsql += vbCrLf + "select T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T1.loccode,T1.stockprice,'' 'Intrserial','' 'Mnfserial','' 'LotNumber','' 'ExpDate','' 'MnfDate','' 'Indate',"
                    strsql += vbCrLf + "0 'MRP',0 'Quantity1' from " & MDBName & "..wtr1 T1 join " & MDBName & "..OITM T2 on T2.ItemCode=T1.ItemCode where T2.ManSerNum='N' and T2.ManBtchNum='N'"
                    strsql += vbCrLf + ")AA join " & MDBName & "..FLAG_TABLE BB on aa.DocEntry=bb.U_FORMCODE  where bb.U_FORMNAME='INVENTRYTRANSFER' AND bb.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                End If
                command_staging.CommandText = strsql
                command_staging.ExecuteNonQuery()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub add_Grpo()
        Try
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from s_opdn"
            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    ''Header Table
                    strsql = "INSERT INTO S_opdn (DOCNUM,SERIES,DOCDATE,DOCCUR,COMMENTS,LUPDATE)"
                    strsql += vbCrLf + "SELECT T1.DocEntry,T1.Series,T1.DocDate,T1.DocCur,T1.Comments,T2.U_LUPDATE  FROM " & MDBName & "..opdn  T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON "
                    strsql += vbCrLf + "T1.DocEntry=T2.U_FORMCODE WHERE T2.U_FORMNAME='GRPO'"

                    strsql += vbCrLf + "INSERT INTO S_pdn1 (DOCNUM,LINENUM,ITEMCODE,DSCRIPTION,QUANTITY,WHSCODE,INTRSERIAL,LOTNUMBER,MNFSERIAL,EXPDATE,MNFDATE,INDATE,MRP,QUANTITY1)"
                    strsql += vbCrLf + "SELECT AA.DocEntry,AA.LineNum,AA.ItemCode,AA.Dscription,AA.Quantity,AA.WhsCode,AA.IntrSerial,AA.MnfSerial,AA.LotNumber,AA.ExpDate,AA.MnfDate,AA.InDate,AA.MRP,AA.QUANTITY1 "
                    strsql += vbCrLf + "FROM(SELECT DISTINCT T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T3.IntrSerial,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',1 'QUANTITY1' FROM"
                    strsql += vbCrLf + " " & MDBName & "..PDN1  T1 JOIN " & MDBName & "..SRI1 T2 ON T1.DocEntry=T2.BaseEntry AND T1.LineNum=T2.BaseLinNum"
                    strsql += vbCrLf + " JOIN " & MDBName & "..OSRI T3 ON T3.SysSerial=T2.SysSerial AND T3.ItemCode=T2.ItemCode "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OSRN T4 ON T4.ItemCode=T3.ItemCode AND T4.DistNumber =T3.IntrSerial "
                    strsql += vbCrLf + " WHERE T2.BaseType='20' "
                    strsql += vbCrLf + "UNION ALL"
                    strsql += vbCrLf + "SELECT DISTINCT  T3.DocEntry,T3.LineNum,T3.ItemCode,T3.Dscription,T3.Quantity,T3.WhsCode,T2.BatchNum,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',T2.Quantity 'QUANTITY1'   FROM " & MDBName & "..IBT1 T2 "
                    strsql += vbCrLf + "JOIN " & MDBName & "..PDN1  T3 ON T3.DocEntry=T2.BaseEntry AND T3.LineNum=T2.BaseLinNum "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OBTN T4 ON T4.DistNumber=T2.BatchNum "
                    strsql += vbCrLf + "WHERE T2.BaseType='20'"
                    strsql += vbCrLf + "Union ALl"
                    strsql += vbCrLf + "select T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,'' 'Intrserial','' 'Mnfserial','' 'LotNumber','' 'ExpDate','' 'MnfDate','' 'Indate',"
                    strsql += vbCrLf + "0 'MRP',0 'Quantity1' from " & MDBName & "..pdn1 T1 join " & MDBName & "..OITM T2 on T2.ItemCode=T1.ItemCode where T2.ManSerNum='N' and T2.ManBtchNum='N'"
                    strsql += vbCrLf + ") AA JOIN " & MDBName & "..FLAG_TABLE BB ON AA.DocEntry=BB.U_FORMCODE WHERE BB.U_FORMNAME='GRPO'"
                Else
                    ''Header Table
                    strsql = "INSERT INTO S_opdn (DOCNUM,SERIES,DOCDATE,DOCCUR,COMMENTS,LUPDATE)"
                    strsql += vbCrLf + "SELECT T1.DocEntry,T1.Series,T1.DocDate,T1.DocCur,T1.Comments,T2.U_LUPDATE  FROM " & MDBName & "..opdn  T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON "
                    strsql += vbCrLf + "T1.DocEntry=T2.U_FORMCODE WHERE T2.U_FORMNAME='GRPO' AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"

                    strsql += vbCrLf + "INSERT INTO S_pdn1 (DOCNUM,LINENUM,ITEMCODE,DSCRIPTION,QUANTITY,WHSCODE,INTRSERIAL,LOTNUMBER,MNFSERIAL,EXPDATE,MNFDATE,INDATE,MRP,QUANTITY1)"
                    strsql += vbCrLf + "SELECT AA.DocEntry,AA.LineNum,AA.ItemCode,AA.Dscription,AA.Quantity,AA.WhsCode,AA.IntrSerial,AA.MnfSerial,AA.LotNumber,AA.ExpDate,AA.MnfDate,AA.InDate,AA.MRP,AA.QUANTITY1 "
                    strsql += vbCrLf + "FROM(SELECT DISTINCT T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T3.IntrSerial,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',1 'QUANTITY1' FROM"
                    strsql += vbCrLf + " " & MDBName & "..PDN1  T1 JOIN " & MDBName & "..SRI1 T2 ON T1.DocEntry=T2.BaseEntry AND T1.LineNum=T2.BaseLinNum"
                    strsql += vbCrLf + " JOIN " & MDBName & "..OSRI T3 ON T3.SysSerial=T2.SysSerial AND T3.ItemCode=T2.ItemCode "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OSRN T4 ON T4.ItemCode=T3.ItemCode AND T4.DistNumber =T3.IntrSerial "
                    strsql += vbCrLf + " WHERE T2.BaseType='20' "
                    strsql += vbCrLf + "UNION ALL"
                    strsql += vbCrLf + "SELECT DISTINCT  T3.DocEntry,T3.LineNum,T3.ItemCode,T3.Dscription,T3.Quantity,T3.WhsCode,T2.BatchNum,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',T2.Quantity 'QUANTITY1'   FROM " & MDBName & "..IBT1 T2 "
                    strsql += vbCrLf + "JOIN " & MDBName & "..PDN1  T3 ON T3.DocEntry=T2.BaseEntry AND T3.LineNum=T2.BaseLinNum "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OBTN T4 ON T4.DistNumber=T2.BatchNum "
                    strsql += vbCrLf + "WHERE T2.BaseType='20'"
                    strsql += vbCrLf + "Union ALl"
                    strsql += vbCrLf + "select T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,'' 'Intrserial','' 'Mnfserial','' 'LotNumber','' 'ExpDate','' 'MnfDate','' 'Indate',"
                    strsql += vbCrLf + "0 'MRP',0 'Quantity1' from " & MDBName & "..pdn1 T1 join " & MDBName & "..OITM T2 on T2.ItemCode=T1.ItemCode where T2.ManSerNum='N' and T2.ManBtchNum='N'"
                    strsql += vbCrLf + ") AA JOIN " & MDBName & "..FLAG_TABLE BB ON AA.DocEntry=BB.U_FORMCODE WHERE BB.U_FORMNAME='GRPO' AND bb.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                End If
                command_staging.CommandText = strsql
                command_staging.ExecuteNonQuery()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub add_GoodsReturn()
        Try
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from S_ORPD"
            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    ''Header Table
                    strsql = "INSERT INTO S_ORPD (DOCNUM,SERIES,DOCDATE,DOCCUR,COMMENTS,LUPDATE)"
                    strsql += vbCrLf + "SELECT T1.DocEntry,T1.Series,T1.DocDate,T1.DocCur,T1.Comments,T2.U_LUPDATE  FROM " & MDBName & "..orpd  T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON "
                    strsql += vbCrLf + "T1.DocEntry=T2.U_FORMCODE WHERE T2.U_FORMNAME='GOODSRETURN'"

                    strsql += vbCrLf + "INSERT INTO S_RPD1 (DOCNUM,LINENUM,ITEMCODE,DSCRIPTION,QUANTITY,WHSCODE,INTRSERIAL,LOTNUMBER,MNFSERIAL,EXPDATE,MNFDATE,INDATE,MRP,QUANTITY1)"
                    strsql += vbCrLf + "SELECT AA.DocEntry,AA.LineNum,AA.ItemCode,AA.Dscription,AA.Quantity,AA.WhsCode,AA.IntrSerial,AA.MnfSerial,AA.LotNumber,AA.ExpDate,AA.MnfDate,AA.InDate,AA.MRP,AA.QUANTITY1 "
                    strsql += vbCrLf + "FROM(SELECT DISTINCT T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T3.IntrSerial,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',1 'QUANTITY1' FROM"
                    strsql += vbCrLf + " " & MDBName & "..rpd1  T1 JOIN " & MDBName & "..SRI1 T2 ON T1.DocEntry=T2.BaseEntry AND T1.LineNum=T2.BaseLinNum"
                    strsql += vbCrLf + " JOIN " & MDBName & "..OSRI T3 ON T3.SysSerial=T2.SysSerial AND T3.ItemCode=T2.ItemCode "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OSRN T4 ON T4.ItemCode=T3.ItemCode AND T4.DistNumber =T3.IntrSerial "
                    strsql += vbCrLf + " WHERE T2.BaseType='21' "
                    strsql += vbCrLf + "UNION ALL"
                    strsql += vbCrLf + "SELECT DISTINCT  T3.DocEntry,T3.LineNum,T3.ItemCode,T3.Dscription,T3.Quantity,T3.WhsCode,T2.BatchNum,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',T2.Quantity 'QUANTITY1'   FROM " & MDBName & "..IBT1 T2 "
                    strsql += vbCrLf + "JOIN " & MDBName & "..rpd1  T3 ON T3.DocEntry=T2.BaseEntry AND T3.LineNum=T2.BaseLinNum "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OBTN T4 ON T4.DistNumber=T2.BatchNum "
                    strsql += vbCrLf + "WHERE T2.BaseType='21'"
                    strsql += vbCrLf + "Union ALl"
                    strsql += vbCrLf + "select T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,'' 'Intrserial','' 'Mnfserial','' 'LotNumber','' 'ExpDate','' 'MnfDate','' 'Indate',"
                    strsql += vbCrLf + "0 'MRP',0 'Quantity1' from " & MDBName & "..RPD1 T1 join " & MDBName & "..OITM T2 on T2.ItemCode=T1.ItemCode where T2.ManSerNum='N' and T2.ManBtchNum='N'"
                    strsql += vbCrLf + ") AA JOIN " & MDBName & "..FLAG_TABLE BB ON AA.DocEntry=BB.U_FORMCODE WHERE BB.U_FORMNAME='GOODSRETURN'"
                Else
                    ''Header Table
                    strsql = "INSERT INTO S_ORPD (DOCNUM,SERIES,DOCDATE,DOCCUR,COMMENTS,LUPDATE)"
                    strsql += vbCrLf + "SELECT T1.DocEntry,T1.Series,T1.DocDate,T1.DocCur,T1.Comments,T2.U_LUPDATE  FROM " & MDBName & "..orpd  T1 JOIN " & MDBName & "..FLAG_TABLE T2 ON "
                    strsql += vbCrLf + "T1.DocEntry=T2.U_FORMCODE WHERE T2.U_FORMNAME='GOODSRETURN' AND T2.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"

                    strsql += vbCrLf + "INSERT INTO S_RPD1 (DOCNUM,LINENUM,ITEMCODE,DSCRIPTION,QUANTITY,WHSCODE,INTRSERIAL,LOTNUMBER,MNFSERIAL,EXPDATE,MNFDATE,INDATE,MRP,QUANTITY1)"
                    strsql += vbCrLf + "SELECT AA.DocEntry,AA.LineNum,AA.ItemCode,AA.Dscription,AA.Quantity,AA.WhsCode,AA.IntrSerial,AA.MnfSerial,AA.LotNumber,AA.ExpDate,AA.MnfDate,AA.InDate,AA.MRP,AA.QUANTITY1 "
                    strsql += vbCrLf + "FROM(SELECT DISTINCT T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,T3.IntrSerial,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',1 'QUANTITY1' FROM"
                    strsql += vbCrLf + " " & MDBName & "..rpd1  T1 JOIN " & MDBName & "..SRI1 T2 ON T1.DocEntry=T2.BaseEntry AND T1.LineNum=T2.BaseLinNum"
                    strsql += vbCrLf + " JOIN " & MDBName & "..OSRI T3 ON T3.SysSerial=T2.SysSerial AND T3.ItemCode=T2.ItemCode "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OSRN T4 ON T4.ItemCode=T3.ItemCode AND T4.DistNumber =T3.IntrSerial "
                    strsql += vbCrLf + " WHERE T2.BaseType='21' "
                    strsql += vbCrLf + "UNION ALL"
                    strsql += vbCrLf + "SELECT DISTINCT  T3.DocEntry,T3.LineNum,T3.ItemCode,T3.Dscription,T3.Quantity,T3.WhsCode,T2.BatchNum,T4.MnfSerial,T4.LotNumber,T4.ExpDate,T4.MnfDate,T4.InDate,0 'MRP',T2.Quantity 'QUANTITY1'   FROM " & MDBName & "..IBT1 T2 "
                    strsql += vbCrLf + "JOIN " & MDBName & "..rpd1  T3 ON T3.DocEntry=T2.BaseEntry AND T3.LineNum=T2.BaseLinNum "
                    strsql += vbCrLf + "JOIN " & MDBName & "..OBTN T4 ON T4.DistNumber=T2.BatchNum "
                    strsql += vbCrLf + "WHERE T2.BaseType='21'"
                    strsql += vbCrLf + "Union ALl"
                    strsql += vbCrLf + "select T1.DocEntry,T1.LineNum,T1.ItemCode,T1.Dscription,T1.Quantity,T1.WhsCode,'' 'Intrserial','' 'Mnfserial','' 'LotNumber','' 'ExpDate','' 'MnfDate','' 'Indate',"
                    strsql += vbCrLf + "0 'MRP',0 'Quantity1' from " & MDBName & "..RPD1 T1 join " & MDBName & "..OITM T2 on T2.ItemCode=T1.ItemCode where T2.ManSerNum='N' and T2.ManBtchNum='N'"
                    strsql += vbCrLf + ") AA JOIN " & MDBName & "..FLAG_TABLE BB ON AA.DocEntry=BB.U_FORMCODE WHERE BB.U_FORMNAME='GOODSRETURN' AND bb.U_LUPDATE>'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                End If
                command_staging.CommandText = strsql
                command_staging.ExecuteNonQuery()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TaxCode_determiantion()
        Try
            ' as per the chage request , done by senthil.k
            write_log(vbNewLine + "-----  Tax Master Check --------")
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from S_OTAX"
            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            write_log("Tax Master Header Insertion : " + strsql)

            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    strsql = "insert S_OTAX (LOCATION_CODE,EFFECTIVE_DATE,VAT_PER,PRODUCT_GROUP,LUPDATE,UPDATE_FLAG) "
                    strsql += vbCrLf + " Select case when isnull(T3.U_excode,'')='' then T3.county else T3.U_excode end as 'LocationCode',EfctFrom 'Effect From',(select top 1 Rate from " & MDBName & "..Ostc where Code=TaxCode)'Vatper',"
                    strsql += vbCrLf + " case when isnull(T4.U_excode,'')='' then convert(nvarchar(50),T4.ItmsGrpCod) else T4.U_excode end as 'ItemGroup',GETDATE(),1"
                    strsql += vbCrLf + " FROM " & MDBName & "..TCD1 T0  INNER JOIN " & MDBName & "..TCD2 T1 ON T0.AbsId = T1.Tcd1Id INNER JOIN " & MDBName & "..TCD3 T2 ON T1.AbsId = T2.Tcd2Id join " & MDBName & "..OLCT T3 on T3.State=T1.KeyFld_2_V join " & MDBName & "..OITB T4 on T1.KeyFld_1_V = T4.ItmsGrpCod "
                    strsql += vbCrLf + " where T3.U_branch='Y' and KeyFld_1='9' and KeyFld_2='6' and T3.county is not null and KeyFld_1_V in ( select U_ITEMGROUPCODE  from " & MDBName & "..[@ITEMGROUP_UDT] where U_RECEIVER='PA')"
                Else
                    'strsql = "insert S_OTAX (LOCATION_CODE,EFFECTIVE_DATE,VAT_PER,PRODUCT_GROUP,LUPDATE,UPDATE_FLAG) "
                    'strsql += vbCrLf + " Select distinct case when isnull(T3.U_excode,'')='' then T3.county else T3.U_excode end as 'LocationCode',EfctFrom 'Effect From',(select top 1 Rate from " & MDBName & "..Ostc where Code=TaxCode)'Vatper',"
                    'strsql += vbCrLf + " case when isnull(T4.U_excode,'')='' then convert(nvarchar(50),T4.ItmsGrpCod) else T4.U_excode end as 'ItemGroup',GETDATE(),1"
                    'strsql += vbCrLf + " FROM " & MDBName & "..TCD1 T0  INNER JOIN " & MDBName & "..TCD2 T1 ON T0.AbsId = T1.Tcd1Id INNER JOIN " & MDBName & "..TCD3 T2 ON T1.AbsId = T2.Tcd2Id join " & MDBName & "..OLCT T3 on T3.State=T1.KeyFld_2_V  join " & MDBName & "..OITB T4 on T1.KeyFld_1_V = T4.ItmsGrpCod JOIN " & MDBName & "..FLAG_TABLE T5 ON T5.U_FORMCODE=T0.AbsId"
                    'strsql += vbCrLf + " where T3.U_branch='Y' and  KeyFld_1='9' and KeyFld_2='6' and T3.county is not null and KeyFld_1_V in ( select U_ITEMGROUPCODE  from " & MDBName & "..[@ITEMGROUP_UDT] where U_RECEIVER='PA') and T5.U_FORMNAME='TAX' AND T5.U_LUPDATE >'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                    strsql = "insert S_OTAX(LOCATION_CODE, EFFECTIVE_DATE, VAT_PER, PRODUCT_GROUP, LUPDATE, UPDATE_FLAG)"
                    strsql += vbCrLf + "Select distinct case when isnull(T3.U_excode,'')='' then T3.county else T3.U_excode end as 'LocationCode',EfctFrom 'Effect From',(select top 1 Rate from " & MDBName & "..Ostc where Code=TaxCode)'Vatper',"
                    strsql += vbCrLf + "case when isnull(T4.U_excode,'')='' then convert(nvarchar(50),T4.ItmsGrpCod) else T4.U_excode end as 'ItemGroup',GETDATE(),1"
                    strsql += vbCrLf + "FROM " & MDBName & "..TCD1 T0  INNER JOIN VSN19.dbo.TCD2 T1 ON T0.AbsId = T1.Tcd1Id INNER JOIN " & MDBName & "..TCD3 T2 ON T1.AbsId = T2.Tcd2Id join " & MDBName & "..OLCT T3 on T3.State=T1.KeyFld_2_V  join " & MDBName & "..OITB T4 on T1.KeyFld_1_V = T4.ItmsGrpCod JOIN " & MDBName & "..FLAG_TABLE T5 ON T5.U_FORMCODE=T0.AbsId"
                    strsql += vbCrLf + "where T3.U_branch='Y' and  KeyFld_1='9' and KeyFld_2='6' and T3.county is not null and KeyFld_1_V in ( select U_ITEMGROUPCODE  from " & MDBName & "..[@ITEMGROUP_UDT] where U_RECEIVER='PA') and T5.U_FORMNAME='TAX' AND T5.U_LUPDATE >'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                End If
            End If
           
            'strsql = "insert S_OTAX (LOCATION_CODE,EFFECTIVE_DATE,VAT_PER,PRODUCT_GROUP,LUPDATE,UPDATE_FLAG) "
            'strsql += vbCrLf + " Select T3.county 'LocationCode',EfctFrom 'Effect From',(select top 1 Rate from " & MDBName & "..Ostc where Code=TaxCode)'Vatper',"
            'strsql += vbCrLf + " (select top 1 ItmsGrpNam  from " & MDBName & "..OITB where ItmsGrpCod=KeyFld_1_V) 'ItemGroup',GETDATE(),1"
            'strsql += vbCrLf + " FROM " & MDBName & "..TCD1 T0  INNER JOIN " & MDBName & "..TCD2 T1 ON T0.AbsId = T1.Tcd1Id INNER JOIN " & MDBName & "..TCD3 T2 ON T1.AbsId = T2.Tcd2Id join " & MDBName & "..OLCT T3 on T3.State=T1.KeyFld_2_V "
            'strsql += vbCrLf + " where KeyFld_1='9' and KeyFld_2='6' and T3.county is not null and KeyFld_1_V in ( select U_ITEMGROUPCODE  from " & MDBName & "..[@ITEMGROUP_UDT] where U_RECEIVER='PA')"
            write_log("Tax Master Added : " + strsql)
            command_staging.CommandText = strsql
            command_staging.CommandTimeout = 100
            command_staging.ExecuteNonQuery()
            'objMVasan.Register_Log(1005, "TAX_MASTER", dataset_log, "Success", objMVasan.Get_Next_Key_Value(1005), 1, Date.Now)

        Catch ex As Exception
            write_log("Tax Master Issue : " + ex.ToString())
            'objMVasan.Register_Log(1005, "TAX_MASTER", dataset_log, ex.ToString, objMVasan.Get_Next_Key_Value(1005), 1, Date.Now)
        End Try
    End Sub

    Private Sub item_Group()
        Try
            write_log(vbNewLine + "-----  Item Group Check --------")
            strsql = "select isnull(MAX(ISNULL(LUPDATE,'')),'') 'LUPDATE' from S_OITB"
            da = New SqlDataAdapter(strsql, con_staging)
            dt = New DataTable
            da.Fill(dt)
            write_log("Item Group Header Insertion : " + strsql)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item("LUPDATE") = "1900-01-01 00:00:00.000" Then
                    ''strsql = "insert into S_ocrd (VENDOR_CODE,VENDOR_NAME,VENDOR_ADDRESS,CURRENCY_CODE,LUPDATE,UPDATE_FLAG)"
                    ''strsql += vbCrLf + "select T1.cardcode,T1.CardName,T1.Address + '_' + T1.Zipcode'Vendor Address',T1.currency,T2.U_LUPDATE,1  from " & MDBName & "..OCRD T1 join " & MDBName & "..flag_table1 T2 on T1.CardCode=T2.U_formcode "
                    ''strsql += vbCrLf + " where T2.U_formname='VENDORMASTER'"
                    'strsql = "insert into S_oitb (itmsgrpcode,itmsgrpname,ITEMGRPPARENTNAME,lupdate,update_flag)"
                    'strsql += vbCrLf + "select ItmsGrpCod ,ItmsGrpNam,U_ITEMCATEGORY,getdate(),1  from " & MDBName & "..oitb where ItmsGrpCod in (select U_ITEMGROUPCODE from " & MDBName & "..[@ITEMGROUP_UDT] where U_RECEIVER='HO')"
                    ' 
                    'strsql += vbCrLf + "select ItmsGrpCod ,ItmsGrpNam,U_ITEMCATEGORY,'A',getdate(),1  from " & MDBName & "..oitb where ItmsGrpCod in (select U_ITEMGROUPCODE from " & MDBName & "..[@ITEMGROUP_UDT] where U_RECEIVER='PA')"

                    strsql = "insert into S_oitb (ITEM_GROUP_CODE,ITEM_GROUP_NAME,ITEM_MAIN_GROUP,ADD_UPDATE,lupdate,update_flag)"
                    strsql += vbCrLf + "select CASE when ISNULL(t1.U_excode,'')='' then convert(varchar(15),t1.ItmsGrpCod)  else T1.U_excode end,CASE when ISNULL(t1.U_exname,'')='' then t1.ItmsGrpNam  else T1.U_exname end,U_ITEMCATEGORY,'A',getdate(),1  from " & MDBName & "..oitb T1 where ItmsGrpCod in (select U_ITEMGROUPCODE from " & MDBName & "..[@ITEMGROUP_UDT] where U_RECEIVER='PA')"
                Else
                    'strsql += vbCrLf + "select t1.ItmsGrpCod ,T1.ItmsGrpNam,U_ITEMCATEGORY,T2.ADD_UPDATE,T2.U_LUPDATE,1  from " & MDBName & "..oitb T1 join " & MDBName & "..FLAG_TABLE T2 on T1.ItmsGrpCod=T2.U_FORMCODE "
                    strsql = "insert into S_oitb (ITEM_GROUP_CODE,ITEM_GROUP_NAME,ITEM_MAIN_GROUP,ADD_UPDATE,lupdate,update_flag)"
                    strsql += vbCrLf + "select distinct CASE when ISNULL(t1.U_excode,'')='' then convert(varchar(15),t1.ItmsGrpCod)  else T1.U_excode end,CASE when ISNULL(t1.U_exname,'')='' then t1.ItmsGrpNam  else T1.U_exname end,U_ITEMCATEGORY,T2.ADD_UPDATE,T2.U_LUPDATE,1  from " & MDBName & "..oitb T1 join " & MDBName & "..FLAG_TABLE T2 on T1.ItmsGrpCod=T2.U_FORMCODE "
                    strsql += vbCrLf + "where  T2.U_FORMNAME='ITEMGROUP' AND T2.U_LUPDATE >'" & CDate(dt.Rows(0).Item("LUPDATE")).ToString("yyyy-MM-dd HH:mm:ss.fff") & "'"
                End If
            End If
            write_log("Item Group Added : " + strsql)
            command_staging.CommandText = strsql
            command_staging.CommandTimeout = 100
            command_staging.ExecuteNonQuery()
            'objMVasan.Register_Log(1005, "ITEM_GROUP", dataset_log, "Success", objMVasan.Get_Next_Key_Value(1007), 0, Date.Now)
        Catch ex As Exception
            write_log("Item Group Issue : " + ex.ToString())
            'objMVasan.Register_Log(1005, "ITEM_GROUP", dataset_log, ex.ToString, objMVasan.Get_Next_Key_Value(1007), 1, Date.Now)
        End Try
    End Sub

#Region "New Connection"

#Region "Staging DB"
    Private Sub connection_StagingDB()
        Try
            Dim Staginginfo As String '= objMVasan.Get_Staging_Conn_String(1001)
            ' Staginginfo = "Server=BSSTEAM7-LAPTOP;uid=sa;pwd=Mipl@1234;database=MIPL_Outbound_Proactive"
            Staginginfo = "Server=192.168.250.6;uid=sa;pwd=Pass4sqlsa;database=MIPL_Outbound_Proactive"
            con_staging = New SqlConnection
            con_staging = New SqlConnection(Staginginfo)
            If con_staging.State = ConnectionState.Closed Then
                con_staging.Open()
            End If
            command_staging = New SqlCommand
            command_staging.Connection = con_staging
        Catch ex As Exception
            MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace)
            ' Return False
        End Try
    End Sub
#End Region

#Region "SAP DB"
    Private Sub connection_SAPDB()
        Dim Staginginfo As String '= objMVasan.Get_Staging_Conn_String(1005)
        Try
            'Staginginfo = "Server=BSSTEAM7-LAPTOP;uid=sa;pwd=Mipl@1234;database=VHC_LIVE"
            Staginginfo = "Server=192.168.250.6;uid=sa;pwd=Pass4sqlsa;database=VSN19"
            con_SAP = New SqlConnection
            con_SAP = New SqlConnection(Staginginfo)
            If con_SAP.State = ConnectionState.Closed Then
                con_SAP.Open()
            End If
            command_sap = New SqlCommand
            command_sap.Connection = con_SAP
        Catch ex As Exception
            MsgBox("Error in SAPDB: " + ex.Message + vbCrLf + "Position : " + ex.StackTrace)
            ' Return False
        End Try
    End Sub
#End Region

    Public Sub getting_SAP_information()
        Try

            'Dim objMVasan As New MIPL_Vasan.Common_Module
            'Dim oMIPL_Vasan_MIPL_Vasan_SAP_Profile As New MIPL_Vasan.MIPL_Vasan_SAP_Profile
            'oMIPL_Vasan_MIPL_Vasan_SAP_Profile = objMVasan.Get_SAP_Profile(1005)
            'MServerName = oMIPL_Vasan_MIPL_Vasan_SAP_Profile.Db_Server_Name
            'MDBName = "VHC_LIVE" ' oMIPL_Vasan_MIPL_Vasan_SAP_Profile.Db_Name
            MDBName = "VSN19" ' oMIPL_Vasan_MIPL_Vasan_SAP_Profile.Db_Name
            'MUID = oMIPL_Vasan_MIPL_Vasan_SAP_Profile.Db_User_Name
            'MPWD = oMIPL_Vasan_MIPL_Vasan_SAP_Profile.Db_Password
            'MLoginType = "S"
            'MSAPUID = oMIPL_Vasan_MIPL_Vasan_SAP_Profile.SAP_UserName
            'MSAPPWd = oMIPL_Vasan_MIPL_Vasan_SAP_Profile.SAP_Password

        Catch ex As Exception
            write_log("Error in GettingSapInfo" & ex.Message)
            'MsgBox(ex.Message)
        End Try
        '' local DB
        ''oMIPL_Vasan_MIPL_Vasan_SAP_Profile = objMVasan.Get_SAP_Profile(1005)
        'MServerName = "BSSTEAM7-LAPTOP" ' oMIPL_Vasan_MIPL_Vasan_SAP_Profile.Db_Server_Name
        'MDBName = "VHC_LIVE" ' oMIPL_Vasan_MIPL_Vasan_SAP_Profile.Db_Name
        'MUID = "manager" 'oMIPL_Vasan_MIPL_Vasan_SAP_Profile.Db_User_Name
        'MPWD = "test" ' oMIPL_Vasan_MIPL_Vasan_SAP_Profile.Db_Password
        'MLoginType = "S"
        'MSAPUID = "sa" 'oMIPL_Vasan_MIPL_Vasan_SAP_Profile.SAP_UserName
        'MSAPPWd = "Mipl@1234" 'oMIPL_Vasan_MIPL_Vasan_SAP_Profile.SAP_Password
    End Sub
#End Region

#Region "old_connection"

#Region "SAPDB"
    'Private Sub connection_SAPDB()
    '    info_notepad("DBInfo_SAP")
    '    Try
    '        con_SAP = New SqlConnection
    '        If MLoginType.ToUpper.Trim <> "S" Then
    '            con_SAP = New SqlConnection("DATA SOURCE = " + MServerName + ";INITIAL CATALOG = " + MDBName + "; INTEGRATED SECURITY = TRUE;")
    '        Else
    '            con_SAP = New SqlConnection("DATA SOURCE = " + MServerName + ";INITIAL CATALOG = " + MDBName + "; USER ID=" + MUID + "; PASSWORD=" + MPWD + ";")
    '        End If
    '        If con_SAP.State = ConnectionState.Closed Then
    '            con_SAP.Open()
    '        End If
    '        command_sap = New SqlCommand
    '        command_sap.Connection = con_SAP
    '    Catch ex As Exception
    '        MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace)
    '        ' Return False
    '    End Try
    'End Sub
#End Region

#Region "Stagingdb"
    'Private Sub connection_StagingDB()
    '    info_notepad("DBInfo_Staging")
    '    Try
    '        con_staging = New SqlConnection
    '        If MLoginType.ToUpper.Trim <> "S" Then
    '            con_staging = New SqlConnection("DATA SOURCE = " + MServerName + ";INITIAL CATALOG = " + MDBName + "; INTEGRATED SECURITY = TRUE;")
    '        Else
    '            con_staging = New SqlConnection("DATA SOURCE = " + MServerName + ";INITIAL CATALOG = " + MDBName + "; USER ID=" + MUID + "; PASSWORD=" + MPWD + ";")
    '        End If
    '        If con_staging.State = ConnectionState.Closed Then
    '            con_staging.Open()
    '        End If
    '        command_staging = New SqlCommand
    '        command_staging.Connection = con_staging
    '    Catch ex As Exception
    '        MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace)
    '        ' Return False
    '    End Try
    'End Sub
#End Region

#Region "Getnotepad info"
    'Private Sub info_notepad(ByVal notepad_filename As String)
    '    Dim objMUtil As New MIPLUtil.GlobalMethods(Application.StartupPath + "\" & notepad_filename & ".ini", 20)
    '    MServerName = objMUtil.xServerName
    '    MDBName = objMUtil.xDBName
    '    MUID = objMUtil.xUID
    '    MPWD = objMUtil.xPWD
    '    MLoginType = objMUtil.xLoginType
    '    MSAPUID = objMUtil.xSAPUser
    '    MSAPPWd = objMUtil.xSAPPwd
    'End Sub
#End Region

#End Region

#Region "Price List Query"
    Private Sub PriceList_PA()
        Try
            'strsql = "insert into S_PRICELIST (LOCATIONCD,PRDCATEGORY,PRODUCT_CODE,VENDOR,MRP,DIS,COSTPRICE,LUPDATE,UPDATE_FLAG)"
            'strsql += vbCrLf + " select (select T2.ListName  from " & MDBName & "..opln T2 where T2.ListNum=T3.PriceList and (T2.U_dealwith ='PA' or T2.U_dealwith='BH') )'Location Code',(( select top 1 ItmsGrpCod  from " & MDBName & "..oitm where itemcode=T3.ItemCode))'ItemGroup',"
            'strsql += vbCrLf + " T3.ItemCode'Item Code',isnull(T1.CardCode,'')'Vendor Code',isnull(T3.price,0)  'MRP',"
            'strsql += vbCrLf + " isnull(T1.Discount,0)'Discount',isnull(T1.Price,0)'Cost Price',getdate(),1"
            'strsql += vbCrLf + " from " & MDBName & " ..ITM1 T3 left outer join " & MDBName & "..ospp T1 on T3.PriceList=T1.ListNum and T3.ItemCode=T1.ItemCode"
            'strsql += vbCrLf + " Where  --T3.itemcode='43979'"
            'strsql += vbCrLf + " (select ItmsGrpCod  from " & MDBName & "..oitm where itemcode=T3.ItemCode) =  '144'  and "
            'strsql += vbCrLf + " isnull(T3.price,0)>0 AND ISNULL(T1.CardCode,'') <> '' and"
            'strsql += vbCrLf + "T1.cardcode in (select T5.cardcode  from " & MDBName & "..OCRD T5 where T5.Groupcode in (select U_BPGROUP from " & MDBName & "..[@miplBPGROUP] where U_PA='Y'))"
            'strsql += vbCrLf + " order by T3.ItemCode "

            write_log(vbNewLine + "-----  Price Master Check --------")
            ' as per the chage request , done by senthil.k , the above query is original query
            strsql = "insert into S_PRICELIST (LOCATIONCD,PRDCATEGORY,PRODUCT_CODE,VENDOR,MRP,DIS,COSTPRICE,LUPDATE,UPDATE_FLAG)"
            strsql += vbCrLf + "Select B.LOC_CODE,A.* from (select case When ISNULL(Tb.U_excode,'')='' then TB.ItmsGrpCod else Tb.U_excode end as 'ItemGroup', "
            strsql += vbCrLf + "T3.ItemCode 'Item Code',isnull(T1.CardCode,'')'Vendor Code',isnull(T3.price,0)  'MRP', "
            strsql += vbCrLf + "  isnull(T1.Discount, 0) 'Discount',isnull(T1.Price,0)'Cost Price',getdate() as 'Date',1 ' Flg'"
            strsql += vbCrLf + "from  " & MDBName & "  ..ITM1 T3 left outer join  " & MDBName & " ..ospp T1 on "
            strsql += vbCrLf + "  T3.PriceList = T1.ListNum And T3.ItemCode = T1.ItemCode"
            strsql += vbCrLf + "Left Outer Join " & MDBName & "  ..OITM as T4 on T3.ItemCode = T4.ItemCode"
            strsql += vbCrLf + " Where (T4.ItmsGrpCod  in (select U_ITEMGROUPCODE from " & MDBName & ".. [@ITEMGROUP_UDT] where u_lensmrp='y' )) and"
            strsql += vbCrLf + "isnull(T3.price,0)>0 AND ISNULL(T1.CardCode,'') <> '' and T1.cardcode in (Select T5.cardcode  from  " & MDBName & " ..OCRD T5 where T5.Groupcode in"
            strsql += vbCrLf + "(select U_BPGROUP from  " & MDBName & " ..[@miplBPGROUP] where U_PA='Y'))  ) as A, S_OLCT as B  where B.UPDATE_FLAG='3'  Order by B.LOC_CODE"

            'strsql = "insert into S_PRICELIST (LOCATIONCD,PRDCATEGORY,PRODUCT_CODE,VENDOR,MRP,DIS,COSTPRICE,LUPDATE,UPDATE_FLAG)"
            'strsql += vbCrLf + " select (select T2.ListName  from " & MDBName & "..opln T2 where T2.ListNum=T3.PriceList and (T2.U_dealwith ='PA' or T2.U_dealwith='Both') )'Location Code',(select top 1 ItmsGrpNam  from " & MDBName & "..oitb A where A.ItmsGrpCod =( select top 1 ItmsGrpCod  from " & MDBName & "..oitm where itemcode=T3.ItemCode))'ItemGroup',"
            'strsql += vbCrLf + " T3.ItemCode'Item Code',isnull(T1.CardCode,'')'Vendor Code',isnull(T3.price,0)  'MRP',"
            'strsql += vbCrLf + " isnull(T1.Discount,0)'Discount',isnull(T1.Price,0)'Cost Price',getdate(),1"
            'strsql += vbCrLf + " from " & MDBName & " ..ITM1 T3 left outer join " & MDBName & "..ospp T1 on T3.PriceList=T1.ListNum and T3.ItemCode=T1.ItemCode"
            'strsql += vbCrLf + " Where --T3.itemcode='43979'"
            'strsql += vbCrLf + " (select ItmsGrpCod  from " & MDBName & "..oitm where itemcode=T3.ItemCode)in(select U_itemgroupcode from " & MDBName & " ..[@ITEMGROUP_UDT ] where U_receiver='PA') and "
            'strsql += vbCrLf + " isnull(T3.price,0)>0 AND ISNULL(T1.CardCode,'') <> '' and"
            'strsql += vbCrLf + "T1.cardcode in (select T5.cardcode  from " & MDBName & "..OCRD T5 where T5.Groupcode in (select U_BPGROUP from " & MDBName & "..[@miplBPGROUP] where U_PA='Y'))"
            'strsql += vbCrLf + " order by T3.ItemCode "
            write_log("Price Master Added : " + strsql)
            command_staging.CommandText = strsql
            command_staging.CommandTimeout = 100
            command_staging.ExecuteNonQuery()
            'objMVasan.Register_Log(1005, "PRICE_LIST", dataset_log, "Success", objMVasan.Get_Next_Key_Value(1005), 1, Date.Now)

        Catch ex As Exception
            write_log("Price Master Issue : " + ex.ToString())
            'objMVasan.Register_Log(1005, "PRICE_LIST", dataset_log, "Success", objMVasan.Get_Next_Key_Value(1005), 1, Date.Now)
        End Try
    End Sub
#End Region

#Region "Write Log"
    Private Sub WriteLog_Item(ByVal Str As String)
        Try

      
        Dim chatlog As String = "C:\MIPL_APPS\PA_Out_Item.txt"
        ' Dim objReader As New System.IO.StreamReader(chatlog)
        Dim sdate As String
        sdate = Now
        'objReader.Close()
        If System.IO.File.Exists(chatlog) = True Then
            Dim objWriter As New System.IO.StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
            objWriter.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            ' MsgBox("Failed to send message!")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub WriteLog_BP(ByVal Str As String)
        Dim chatlog As String = "C:\MIPL_APPS\PA_Out_BP.txt"
        ' Dim objReader As New System.IO.StreamReader(chatlog)
        Dim sdate As String
        sdate = Now
        'objReader.Close()
        If System.IO.File.Exists(chatlog) = True Then
            Dim objWriter As New System.IO.StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
            objWriter.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            ' MsgBox("Failed to send message!")
        End If
    End Sub

    Private Sub WriteLog_INV(ByVal Str As String)
        Dim chatlog As String = "C:\MIPL_APPS\PA_Out_INV.txt"
        ' Dim objReader As New System.IO.StreamReader(chatlog)
        Dim sdate As String
        sdate = Now
        'objReader.Close()
        If System.IO.File.Exists(chatlog) = True Then
            Dim objWriter As New System.IO.StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
            objWriter.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            ' MsgBox("Failed to send message!")
        End If
    End Sub

    Private Sub WriteLog_General(ByVal Str As String)
        Dim chatlog As String = "C:\MIPL_APPS\PA_Out_Gen.txt"
        ' Dim objReader As New System.IO.StreamReader(chatlog)
        Dim sdate As String
        sdate = Now
        'objReader.Close()
        If System.IO.File.Exists(chatlog) = True Then
            Dim objWriter As New System.IO.StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
            objWriter.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            ' MsgBox("Failed to send message!")
        End If
    End Sub

    Private Sub write_log(ByVal status As String)
        Dim fs As FileStream
        Dim objWriter As System.IO.StreamWriter
        Dim chatlog As String
        Try
            If time = "" Then time = "\Log_" & Now.ToString("yyyyMMdd HHmmss tt")
            Dim di As DirectoryInfo = New DirectoryInfo("C:\QL\PASync_log")
            If di.Exists Then
            Else
                di.Create()
            End If
            chatlog = "C:\QL\PASync_log" & time & ".txt"
            If File.Exists(chatlog) Then
            Else
                fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
                fs.Close()
            End If
            objWriter = New System.IO.StreamWriter(chatlog, True)
            If status <> "" Then objWriter.WriteLine("PA OutBound" & " : " & status)
            objWriter.Close()
        Catch ex As Exception
            ' MsgBox("createlog :" + ex.ToString)
        End Try
    End Sub

#End Region
    Public Function FunCheckLicense() As Boolean
        Try
            Dim TDate As String = Date.Today.ToString("yyyyMMdd")
            If TDate < "20220321" Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try
    End Function
End Module
