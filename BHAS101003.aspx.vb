Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports B2BEntity
Imports B2BService
Imports IntelliSys.NetExpress.Entity
Imports IntelliSys.NetExpress.Web
Imports Newtonsoft.Json
Imports PUBEntity
Imports RestSharp


Public Class BHAS101003
    Inherits MasterDetailPageModel

    Dim sbErrAll As String = String.Empty
    Dim IOFFICER As String = ""
    Dim K077Token As String = ""
    Dim HeaderToken As String = ""
    Dim K077NewFlag As Boolean = False
    Dim btnQueryClick As Boolean = False


    Protected UcRocDobjbgn As ucRocDate
    Protected UcRocDobjend As ucRocDate
    Protected Birtdy_IINS As ucRocDate
    Protected UcRocDbirth As ucRocDate 'As Global.PUBControl.ucRocDateJS


    '是否為新壽通路K077
    ReadOnly Property bolK077() As Boolean
        Get
            Return Me.txtCompanyid.Text = "K077"
        End Get
    End Property

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim script As String

        script = "<script language='javascript' src='" & ResolveUrl("js/BHAS101002.js") & "'></script>" & Environment.NewLine
        ClientScript.RegisterClientScriptBlock(Me.GetType, "HASCommon", script)
        script = "<script language='javascript' src='" & ResolveUrl("js/BHASCommon.js") & "'></script>" & Environment.NewLine
        ClientScript.RegisterClientScriptBlock(Me.GetType, "BHASCommon", script)

        If Not ClientScript.IsClientScriptBlockRegistered("ucCmnZipCdoe") Then
            script = "<script language='javascript' src='" & Me.ResolveUrl("~/PUBControl/CMN/ucCmnZipCode.js") & "'></script>" & Environment.NewLine
            ClientScript.RegisterClientScriptBlock(Me.GetType, "ucCmnZipCode", script)
        End If
        'ASP.NET control在chrome下的相容性處理
        'Page.ClientTarget = "uplevel"

        If Not Me.IsPostBack Then
            '###################################################################################
            '##如果有新增[通路代號]or[帳號]的相關卡控，需留意LogicAfterAssignACC()是否一併調整##
            '###################################################################################

            txtIsSPEChannel.Text = ""
            txtIsSPEChannel.Visible = False
            MyBase.UISetConfirmMessage(UICommands.csDelete, "確認[要保取消]？要保取消後可能無法依原保期承接，請確認是否仍要繼續執行要保取消？")
            'F115031101' MyBase.UISetConfirmMessage("SaveTemp", "將會暫存一筆新的單號，確認[資料暫存]?")
            MyBase.UISetConfirmMessage("SaveTemp_New", "將會暫存一筆新的單號，確認[資料暫存]?")
            MyBase.UISetConfirmMessage("UpdateTemp", "確認[資料更新]?")
            MyBase.UISetConfirmMessage("SaveEdit", "確認[修改]?")
            MyBase.UISetConfirmMessage("PrintManager5", "本保險配合收費出單，須提供繳費資料憑證，待本公司核保審核通過後，將於起保日前（含）進行扣款，本保險始生效力")
            MyBase.UISetConfirmMessage("PayMoney", "本保險配合收費出單，須於保期生效前完成繳費，待本公司核保審核通過後，本保險始生效力")         '是否使用EKYC
            MyBase.UISetConfirmMessage("ModifyWhenY", "調整要、被保人的電話、Email、地址等基本資訊，僅限要保確認狀態可行。")
            'If CommonFunc.EKYCOrNot(Me, "BHAS101003", Me.ScsOrgID) Then
            '	MyBase.UISetConfirmMessage("SaveCheck", "將自動轉跳至EKYC頁面。請留意瀏覽器視窗右上方有無【封鎖彈跳\快顯視窗】的訊息，若有，請點選【永遠允許彈跳\快顯視窗】。")
            'End If

            '點選「線上繳費」出現提示語：「本保險配合收費出單，須提供繳費資料憑證，待本公司核保審核通過後，將於起保日前（含）進行扣款，本保險始生效力」，並有按鈕供保戶按「確認」或「取消」
            '點選「多元化繳費單」出現提示語：「本保險配合收費出單，須於保期生效前完成繳費，待本公司核保審核通過後，本保險始生效力」，並有按鈕供保戶按「確認」或「取消」  

            '''MyBase.UISetConfirmMessage("PayMoney", "若使用多元化繳費單繳費，請將經辦聯與要保文件一併交出")
            Me.BaseDoClear()
            GetWsURL_IQP(CBool(ConfigurationManager.AppSettings("bHASB2CRequest")))

            If Me.ScsOrgID = "SKI" Then
                '檢核登錄證字號
                '取得SK員工管理人EMAIL與經辦名稱
                If Me.ScsUserID.Substring(0, 4) = "SKI\" Then
                    Dim MONEY_PAPERS As String = ""
                    Dim PSNiofficer As String = ""
                    PSNiofficer = Me.ScsUserID.Substring(4, 6)
                    Dim dsRtn As DataSet = Me.CallCommonService("PUBService.bsSingleSQL", New String() {"PSNV_SALES", "IOFFICER ='" + PSNiofficer + "' And sysdate between DMONEY_START And DMONEY_END", "*"})
                    If dsRtn.Tables(0).DefaultView.Count > 0 Then
                        MONEY_PAPERS = dsRtn.Tables(0).Rows(0)("MONEY_PAPERS")
                    Else
                        InsertUP_BTN.Text = "X"
                        UIAlertMessage("業務員 查無有效登錄證字號")
                    End If
                    dsRtn.Dispose()
                End If
            End If
            'B2B作業程式，當通路別=KKSB增加功能 A112072500 #6577
            If Me.ScsUserID.Substring(0, 5) = "KKSB\" Then
                Dim ScsUserIDS As String() = Me.ScsUserID.Split("\"c)
                If ScsUserIDS(1).Trim.Substring(0, 4) = "KKSB" Then
                    '帳號前4碼非[KKSB]者,將B2B帳號傳入內控資料[通路人員代號]欄位。
                Else
                    txtIRECPESN.Text = ScsUserIDS(1).ToString.Trim
                End If
            End If

            'B2B作業程式，當通路別=KKEU增加功能 F114101300 通路單位預設070705
            If Me.ScsUserID.Substring(0, 5) = "KKEU\" Then
                '帳號前4碼[KKEU]通路單位預設070705。
                txtIrecunit.Text = "070705"
            End If

            ''因應颱風來襲，10/01~10/4不受理韓國旅遊投保、 10/01~10/5不受理日本旅遊投保
            'Me.b2b_title.Text = "因應颱風來襲，10/01~10/4不受理韓國旅遊投保、 10/01~10/5不受理日本旅遊投保"
            Dim HAS_B2B_TITLE As String = ""
            Dim dsTitle As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_B2B_TITLE' AND CODE ='HAS_B2B_TITLE'", "DATA1"})

            If Not IsNothing(dsTitle) AndAlso dsTitle.Tables.Count > 0 AndAlso dsTitle.Tables(0).Rows.Count > 0 Then
                HAS_B2B_TITLE = dsTitle.Tables(0).Rows(0)("DATA1").ToString.Trim
                Me.b2b_title.Text = HAS_B2B_TITLE
            End If

            '===有權限才顯示 HAS_SET_ACCOUNT
            Dim sUserid As String() = Me.ScsUserID.Split("\")
            Me.txtCompanyid.Text = sUserid(0)
            Me.txtUserid.Text = sUserid(1)

            ' F115030200 註記下列 (原HAS_SET_ACCOUNT信律保經可以調整通路代號與帳號，但早在此需求單之前已無提供信律保經使用)
            'Dim HAS_SET_ACCOUNT As String = ""
            'HAS_SET_ACCOUNT = Me.txtCompanyid.Text.Trim & "\" & Me.txtUserid.Text.Trim
            'Dim dsAccount As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_SET_ACCOUNT' AND CODE ='" + HAS_SET_ACCOUNT + "'", "CODE,NAME"})

            'If dsAccount.Tables(0).Rows.Count > 0 Then
            '    Me.txtCompanyid.Enabled = False
            '    Me.txtUserid.Enabled = True
            'Else
            '    Me.txtCompanyid.Enabled = False
            '    Me.txtUserid.Enabled = False
            'End If
            '===有權限才顯示 HAS_SET_ACCOUNT

            Dim dsOfficer As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.getiofficer", Me.ScsUserID)
            If dsOfficer.Tables("Organization").Rows.Count > 0 Then
                IOFFICER = dsOfficer.Tables("Organization").Rows(0).Item("Iofficer")
                Me.labIOFFICER.Text = IOFFICER
                '"KEJ07600"
                '農經資料查詢
                If Me.ScsOrgID = "KEJ0" Then
                    GetIrecItems(IOFFICER)
                End If

            Else
                'Throw New Exception("找不到本組織的[單位代號]或[經辦代號]")
                dllGetIrec.Visible = False
            End If


            If Not Me.StateTransfer Is Nothing Then
                BaseDoClear()

                Dim basic As String = CType(StateTransfer.Args(0), String)

                If basic.Length = 14 Then
                    Me.txtIQPMainOid.Text = basic
                    Me.BaseDoQuery()

                    If Me.txtIQPMainOid.Text.Trim <> "" Then

                        Dim Rtype As String = CType(StateTransfer.Args(1), String)
                        Me.txtRtype.Text = Rtype
                        '資料顯示類別    E=編輯  V=預覽
                        If Me.txtRtype.Text = "V" Then

                            InsertUP_BTN.Text = "X"
                            'IISSU_N.Enabled = False
                            CK2.Checked = True

                        ElseIf Me.txtRtype.Text = "E" Then
                            'CheckS.Text = "1"
                            'CheckSNNN.Text = "請[重新計算保費]"
                            'ReCaculate.Enabled = True
                        Else
                            InsertUP_BTN.Text = "X"
                            'IISSU_N.Enabled = False
                            CK2.Checked = True

                        End If

                        '設定保費顯示
                        If txtPESTATUS_NAME.Text = "[送審中]" Or
                           txtPESTATUS_NAME.Text = "[檢核中]" Or
                           txtPESTATUS_NAME.Text = "[要保確認]" Or
                           txtPESTATUS_NAME.Text = "[要保取消]" Or
                           txtPESTATUS_NAME.Text = "[已轉檔]" Then
                            ShowPremInfoOrNot(True)
                        Else
                            ShowPremInfoOrNot(False)
                        End If

                    End If
                End If

            End If

            '要、被保人國別下拉選單給值
            InitCountryDDL()

            ''海突是否含傳染病的RadioBtn初始設定
            'rad_Disease_CheckedChanged(sender, e)

            '20251218 特殊通路才可投保方案【海外標準型】&【海外豪華型】不提供海突不含傳染病
            If txtCase.Text <> "J 海外標準型" AndAlso txtCase.Text <> "K 海外豪華型" AndAlso txtCase.Text <> "L 一般型或申根型" Then
                titIT64.Attributes.Add("style", "display:none")
                ddlIT64.Attributes.Add("style", "display:none")
                BT8.Attributes.Add("style", "display:none")
                If txtCase.Text = "僅人身保險" Or txtCase.Text = "A 國內海陸型" Then
                    ddlIT6E.Attributes.Add("style", "display:none")
                ElseIf txtCase.Text = "M 海外安心型" Or txtCase.Text = "N 海外超值型" Then
                    ddlIT63.Attributes.Add("style", "display:none")
                    ddlIT6E.Attributes.Add("style", "display:none")
                    ddlIT64.Attributes.Add("style", "display:inline")
                Else
                    ddlIT63.Attributes.Add("style", "display:none")
                End If
            End If
            If checkSpecialChannel().Item1 = True Then
                BT6.Attributes.Add("style", "display:inline")
                BT7.Attributes.Add("style", "display:inline")
                BT9.Attributes.Add("style", "display:none")
                BT10.Attributes.Add("style", "display:none")
                BT11.Attributes.Add("style", "display:none")
                'Me.BT6.Visible = True
                'Me.BT7.Visible = True
            ElseIf checkChannelO().Item1 = True Then
                BT3.Attributes.Add("style", "display:inline")
                BT5.Attributes.Add("style", "display:inline")
                BT6.Attributes.Add("style", "display:inline")
                BT7.Attributes.Add("style", "display:inline")
                BT9.Attributes.Add("style", "display:none")
                BT10.Attributes.Add("style", "display:none")
                BT11.Attributes.Add("style", "display:inline")
            ElseIf checkChannelMN().Item1 Then
                BT3.Attributes.Add("style", "display:none")
                BT5.Attributes.Add("style", "display:none")
                BT6.Attributes.Add("style", "display:none")
                BT7.Attributes.Add("style", "display:none")
                BT9.Attributes.Add("style", "display:inline")
                BT10.Attributes.Add("style", "display:inline")
                BT11.Attributes.Add("style", "display:none")
            Else
                BT6.Attributes.Add("style", "display:none")
                BT7.Attributes.Add("style", "display:none")
                BT9.Attributes.Add("style", "display:none")
                BT10.Attributes.Add("style", "display:none")
                BT11.Attributes.Add("style", "display:none")
                'Me.BT6.Visible = False
                'Me.BT7.Visible = False
            End If


            '航班快速理賠服務區塊
            If StateMain.Tables("HASM_FLIGHT_DETAIL").Rows.Count > 0 Then
                btnOpenFlight.Text = "-"
                Me.Show5.Visible = True
                Me.Show6.Visible = True
                Me.Show7.Visible = True
            Else
                btnOpenFlight.Text = "+"
                Me.Show5.Visible = False
                Me.Show6.Visible = False
                Me.Show7.Visible = False
            End If
        End If
        '在這裡放置使用者程式碼以初始化網頁
        AjaxService.RegisterAjaxService(Me, New String() {"B2BService.bsITP101000GetOracleValue"})

        Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csUpdate, "").Caption = "要保確認"
        'UICommands.cmdDelete
        Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csDelete, "").Caption = "要保取消"

        MyBase.UISetConfirmMessage(UICommands.csUpdate, "「本案僅完成投保資料輸入，尚未完成核保作業，請於" & Me.UcRocDobjbgn.Text & "(即保期起日)前將要保書回傳至本公司並完成繳費，逾期未回傳者，視同未投保成功，如發生理賠事故，本公司不負理賠之責。」")



        If txtTobjbgn.Text = "" Then
            txtTobjbgn.Text = "12"
        End If
        If txtTobjend.Text = "" Then
            txtTobjend.Text = "12"
        End If
        If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
            Benefi.Text = "法定繼承人"
        End If

        CK2.Enabled = False

        If ddlCountry_B1.SelectedValue <> "" Or ddlCountry_C1.SelectedValue <> "" Or ddlCountry_D1.SelectedValue <> "" Or
           ddlCountry_E1.SelectedValue <> "" Or ddlCountry_F1.SelectedValue <> "" Or ddlCountry_G1.SelectedValue <> "" Or ddlCountry_H1.SelectedValue <> "" Then
            MoreCountryVisible(True)
        Else
            MoreCountryVisible(False)
        End If

        If Benefi2.Text = "" Then
            Me.ShowS22.Visible = False
            Me.ShowS23.Visible = False
        Else
            Me.ShowS22.Visible = True
            Me.ShowS23.Visible = True
        End If

        '初始設定 不是查詢進來的
        If ddlCountry_A1.SelectedValue = "" And txtIQPMainOid.Text = "" Then
            Me.TRAVEL_hidden.Visible = False
            Me.ShowS1.Visible = True
        End If

        '取得PASSPOR權限
        Dim PassportAccess As String = ""
        Dim dsSing As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_PASSPORT' AND CODE ='" + Me.ScsUserID + "'", "CODE"})

        If Not IsNothing(dsSing) AndAlso dsSing.Tables.Count > 0 AndAlso dsSing.Tables(0).Rows.Count > 0 Then
            PassportAccess = dsSing.Tables(0).Rows(0)("CODE").ToString.Trim
        End If

        '1140616 開啟供使用者勾選
        PassportShow.Text = "1"
        'Me.CHK_P1.Visible = False
        Me.CHK_P1.Attributes.Add("style", "display:none;")

        '初始設定 有權限才能看到(若要復原，需要使用Attributes設定，否則前端沒有CHK_P1 & CHK_P2標籤，JS會出錯)
        If PassportAccess <> "" Then
            'Me.CHK_P1.Visible = True
            'Me.CHK_P2.Visible = True
            'PassportShow.Text = "1"
        Else
            'Me.CHK_P1.Visible = False
            'Me.CHK_P2.Visible = False
            'PassportShow.Text = "0"
        End If

        '抓通路代號的權限 通路檢核
        If ChannelColsShowOrNot() Then
            Me.Shows3.Visible = True
            Me.Shows4.Visible = True
            'If Me.ScsOrgID = "K077" Or Me.ScsOrgID = "KA77" Then
            '    txtIREC.Text = "2017000XXXXXX"
            'End If
        Else
            If txtIREC.Text = "" Then
                Me.Shows3.Visible = False
                Me.Shows4.Visible = False
            Else
                Me.Shows3.Visible = True
                Me.Shows4.Visible = True
            End If
        End If

        txtIiofficer.Text = ScsOrgID '在JS分辨是否鎖定團費 K077一率不開團費

        '取得鎖定集體費率權限
        Dim HAS_GROUP As String = ""
        Dim dsSingG As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_GROUP' AND CODE ='" + Me.ScsUserID + "'", "CODE"})

        If Not IsNothing(dsSingG) AndAlso dsSingG.Tables.Count > 0 AndAlso dsSingG.Tables(0).Rows.Count > 0 Then
            HAS_GROUP = dsSingG.Tables(0).Rows(0)("CODE").ToString.Trim
        End If
        If HAS_GROUP <> "" Then
            CHK5.Checked = True
            CHK5_G.Text = "直費限定"
            CHK5.Enabled = False
        Else
            CHK5.Checked = False
            CHK5_G.Text = ""
            CHK5.Enabled = False
        End If
        '航班是否顯示
        Dim dsNew As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_IFLIGHT'  AND STATUS='1'", "DATA1,DATA2,DATA3,NREMARK"})
        If dsNew.Tables(0).Rows.Count > 0 Then
            '有資料要顯示
            flight_title.Visible = True
            rad_flight_Y.Visible = True
            rad_flight_N.Visible = True
        Else
            flight_title.Visible = False
            rad_flight_Y.Visible = False
            rad_flight_N.Visible = False
        End If
        '20221011目前不提供集體費率投保，請取消勾選集體費率。
        'CHK5.Checked = False
        'CHK5.Enabled = False

        AjaxPro.Utility.RegisterTypeForAjax(GetType(BCAR101000))

    End Sub

    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
        BasePreRender()
        Dim js As String = ""
        'Me.UcRocDbirth.TextBox.Attributes("onblur") = Me.UcRocDbirth.TextBox.Attributes("onblur") + "GetAgeInLaw(" + "'" + Me.UcRocDbirth.TextBox.ClientID + "'" + "," + "'" + Me.UcRocDobjbgn.TextBox.ClientID + "'" + "," + "'" + Me.txtDBIRTH.ClientID + "'" + ");"
        Me.UcRocDbirth.TextBox.Attributes("onblur") = Me.UcRocDbirth.TextBox.Attributes("onblur") + "GetAgeInLaw(" + "'" + Me.UcRocDbirth.TextBox.ClientID + "'" + "," + "'" + Me.UcRocDobjbgn.TextBox.ClientID + "'" + "," + "'" + Me.txtDBIRTH.ClientID + "'" + ");"
        UTF.AppendUniqueAttributeValue(Me.UcRocDobjend.TextBox, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(Me.UcRocDobjbgn.TextBox, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(txtTobjbgn, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(txtTobjend, "onblur", "JS301DisplayIQP_Days();")

        'UTF.AppendUniqueAttributeValue(Me.UcRocDobjbgn.TextBox, "onblur", "checkdate3();")

        UTF.AppendUniqueAttributeValue(ddlCountry_A2, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(ddlCountry_B2, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(ddlCountry_C2, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(ddlCountry_D2, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(ddlCountry_E2, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(ddlCountry_F2, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(ddlCountry_G2, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(ddlCountry_H2, "onblur", "JS301DisplayIQP_Days();")
        UTF.AppendUniqueAttributeValue(TRAVEL_BEGIN, "onblur", "JS301DisplayIQP_Days();")
        '郵遞區號
        UTF.AppendUniqueAttributeValue(IINS_AREA, "onblur", "GetAddressByZipcode(" + IINS_AREA.ClientID + ", " + IINS_ADDR.ClientID + ");")
        UTF.AppendUniqueAttributeValue(IISSU_AREA, "onblur", "GetAddressByZipcode(" + IISSU_AREA.ClientID + ", " + IISSU_ADDR.ClientID + ");")
        UTF.AppendUniqueAttributeValue(txtZIPREL, "onblur", "GetAddressByZipcode(" + txtZIPREL.ClientID + ", " + txtAREL.ClientID + ");")
        UTF.AppendUniqueAttributeValue(IINS_ADDR, "onBlur", "GetZipcodeByAddress(" & IINS_ADDR.ClientID & ", " & IINS_AREA.ClientID & ");")
        UTF.AppendUniqueAttributeValue(IISSU_ADDR, "onBlur", "GetZipcodeByAddress(" & IISSU_ADDR.ClientID & ", " & IISSU_AREA.ClientID & ");")
        UTF.AppendUniqueAttributeValue(txtAREL, "onBlur", "GetZipcodeByAddress(" & txtAREL.ClientID & ", " & txtZIPREL.ClientID & ");")


        If InsertUP_BTN.Text = "X" Then
            Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csUpdate, "").Enabled = False
            Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csQuery, "").Enabled = False
            'F115031101'Me.StateCommandBar.CommandState.FindByNameArgument("SaveTemp", "").Enabled = False
            Me.StateCommandBar.CommandState.FindByNameArgument("SaveTemp_New", "").Enabled = False
            Me.StateCommandBar.CommandState.FindByNameArgument("UpdateTemp", "").Enabled = False
            'IISSU_N.Enabled = False
        Else
            If txtPESTATUS.Text <> "" Then
                Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csUpdate, "").Enabled = True
                Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csQuery, "").Enabled = True
                'IISSU_N.Enabled = True
            End If
        End If

        'F115031101 重啟複製功能 且不論案件狀態 \ 20181115 要拿掉了 因為新舊險種代號不同的關係
        MyBase.StateCommandBar.CommandState.FindByNameArgument("SCOPY", "").Visible = False

        '修改按鈕顯示邏輯
        If String.IsNullOrEmpty(txtIQPMainOid.Text) = False AndAlso txtPESTATUS.Text = "S" AndAlso labIOFFICER.Text = "K4560000" Then
            MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveEdit", "").Visible = True
        Else
            MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveEdit", "").Visible = False
        End If
        ' 除了Pestatus = 'Y'以外，一律不顯示此按鈕（基本資料修改）(調整要被保人)
        MyBase.StateCommandBar.CommandState.FindByNameArgument("ModifyWhenY", "").Visible = False

        '各種按鈕
        If Not StateMain Is Nothing Then
            If AppGetDataGridBoundDataView(dgdMain, StateMain).Count > 0 Then
                If txtPESTATUS.Text = "Y" Or txtPESTATUS.Text = "C" Then
                    If Me.txtCompanyid.Text = "KEN0" Then '和泰飽經
                        MyBase.StateCommandBar.CommandState.FindByNameArgument("PayMoney", "").Visible = False
                    Else
                        MyBase.StateCommandBar.CommandState.FindByNameArgument("PayMoney", "").Visible = True
                    End If
                    'If ScsOrgID = "SKI" Then
                    '	MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager2", "").Visible = False
                    '	MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager4", "").Visible = True
                    'Else
                    '	'非SKI
                    '	'要保完成刪除投保證明、收據、申根簽證文件列印功能(須完全隱藏)
                    '	MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager2", "").Visible = False
                    '	MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager4", "").Visible = False
                    'End If

                    MyBase.StateCommandBar.CommandState.FindByNameArgument(UICommands.csDelete, "").Visible = True
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager", "").Visible = True
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager5", "").Visible = True
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp_New", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("UpdateTemp", "").Visible = False
                    'F115031101'MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveCheck", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("Print_Combine", "").Visible = True '20240709整合列印
                    If txtPESTATUS.Text = "Y" Then
                        MyBase.StateCommandBar.CommandState.FindByNameArgument("ModifyWhenY", "").Visible = True
                    End If
                    'Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csDelete, "").Caption = "要保取消"
                    'UICommands.cmdDelete

                    '要保確認 資料轉檔看得到 複製 
                    'F115031101 重啟複製功能 \ 20181115 要拿掉了 因為新舊險種代號不同的關係
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SCOPY", "").Visible = True
                ElseIf txtPESTATUS.Text = "R" Or txtPESTATUS.Text = "W" Or txtPESTATUS.Text = "E" Then
                    '要保取消 只能暫存檔
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("PayMoney", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument(UICommands.csDelete, "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument(UICommands.csUpdate, "").Visible = False

                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager", "").Visible = False
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager2", "").Visible = False
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager3", "").Visible = False
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager4", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager5", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp_New", "").Visible = True
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("UpdateTemp", "").Visible = False
                    'F115031101'MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp", "").Visible = True
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveCheck", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("Print_Combine", "").Visible = False '20240709整合列印
                    txtPAY_NUM.Text = ""
                    txtIQUO.Text = ""
                    txtIKYC.Text = ""
                Else
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("PayMoney", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument(UICommands.csDelete, "").Visible = False
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager", "").Visible = False
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager2", "").Visible = False
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager3", "").Visible = False
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager4", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager5", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("Print_Combine", "").Visible = False '20240709整合列印
                End If
                If txtPESTATUS.Text.Trim = "" Then
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp_New", "").Visible = True
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("UpdateTemp", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveCheck", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument(UICommands.csUpdate, "").Visible = False
                Else
                    If (txtPESTATUS.Text = "Y" Or txtPESTATUS.Text = "R" Or txtPESTATUS.Text = "C") Then
                        MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveCheck", "").Visible = False
                    Else
                        MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveCheck", "").Visible = True
                        If txtPESTATUS.Text = "S" Then
                            If btnQueryClick = True Then
                                MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveCheck", "").Visible = False
                                'F115031101'MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp", "").Enabled = True
                                btnQueryClick = False
                            End If
                            MyBase.StateCommandBar.CommandState.FindByNameArgument("UpdateTemp", "").Visible = True
                            MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp_New", "").Visible = False
                        End If
                    End If

                    'F115031101 重啟複製功能 且不論案件狀態
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SCOPY", "").Visible = True

                End If
            Else
                If txtPESTATUS.Text = "S" Then
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("UpdateTemp", "").Visible = True
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp_New", "").Visible = False
                    'F115031101 重啟複製功能 且不論案件狀態
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SCOPY", "").Visible = True
                ElseIf txtPESTATUS.Text = "" Then
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("UpdateTemp", "").Visible = False
                End If
                MyBase.StateCommandBar.CommandState.FindByNameArgument("PayMoney", "").Visible = False
                MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveCheck", "").Visible = False
                MyBase.StateCommandBar.CommandState.FindByNameArgument(UICommands.csUpdate, "").Visible = False
                MyBase.StateCommandBar.CommandState.FindByNameArgument(UICommands.csDelete, "").Visible = False
                'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager", "").Visible = False
                'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager2", "").Visible = False
                'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager3", "").Visible = False
                'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager4", "").Visible = False
                MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager5", "").Visible = False
                MyBase.StateCommandBar.CommandState.FindByNameArgument("Print_Combine", "").Visible = False '20240709整合列印
            End If
        Else
            MyBase.StateCommandBar.CommandState.FindByNameArgument("PayMoney", "").Visible = False
            MyBase.StateCommandBar.CommandState.FindByNameArgument("UpdateTemp", "").Visible = False
            MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveCheck", "").Visible = False
            MyBase.StateCommandBar.CommandState.FindByNameArgument(UICommands.csDelete, "").Visible = False
            'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager", "").Visible = False
            'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager2", "").Visible = False
            'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager3", "").Visible = False
            'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager4", "").Visible = False
            MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager5", "").Visible = False
            MyBase.StateCommandBar.CommandState.FindByNameArgument("Print_Combine", "").Visible = False '20240709整合列印
        End If

        If Me.txtCompanyid.Text = "KKFH" Then '886威盛保經不顯示線上繳費
            MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager5", "").Visible = False
        End If

        'Y要鎖欄位
        If txtPESTATUS.Text = "Y" Or txtPESTATUS.Text = "C" Then
            LockPage()
        End If

        Dim sJsUcase As String = "this.value=this.value.toUpperCase();"
        'txtUserid
        UTF.AppendUniqueAttributeValue(txtUserid, "onblur", sJsUcase)
        UTF.AppendUniqueAttributeValue(Benefi_ID, "onblur", sJsUcase)
        UTF.AppendUniqueAttributeValue(Benefi, "onblur", "BidChange();")
        UTF.AppendUniqueAttributeValue(Benefi_ID, "onblur", "BidChange();")

        UTF.AppendUniqueAttributeValue(Benefi_ID2, "onblur", sJsUcase)
        UTF.AppendUniqueAttributeValue(Benefi2, "onblur", "BidChange();")
        UTF.AppendUniqueAttributeValue(Benefi_ID2, "onblur", "BidChange();")

        UTF.AppendUniqueAttributeValue(Benefi_ID3, "onblur", sJsUcase)
        UTF.AppendUniqueAttributeValue(Benefi3, "onblur", "BidChange();")
        UTF.AppendUniqueAttributeValue(Benefi_ID3, "onblur", "BidChange();")
        UTF.AppendUniqueAttributeValue(txtFlightNum, "onblur", sJsUcase)


        Me.txtTobjbgn.Style("Width") = "25px"
        Me.txtTobjend.Style("Width") = "25px"
        Me.txtIQPITP_Days.Style("Width") = "30px"
        Me.IINS_AREA.Style("Width") = "50px"
        Me.IINS_ADDR.Style("Width") = "300px"
        Me.txtDBIRTH.Style("Width") = "20px"
        Me.IISSU_AREA.Style("Width") = "50px"
        Me.IISSU_ADDR.Style("Width") = "300px"
        Me.NAME_IINS.Style("Width") = "100px"
        Me.NAME_IISSU.Style("Width") = "100px"
        Me.txtISEQ.Style("Width") = "0px"
        Me.txtIISSU_Ori.Style("Width") = "0px"
        Me.txtZIPREL.Style("Width") = "50px"
        Me.txtAREL.Style("Width") = "300px"
        Me.NAME_IINS.Style("Width") = "100px"
        Me.NAME_IISSU.Style("Width") = "100px"
        Me.txtIins_IINS2.Style("Width") = "80px"
        Me.txtIins_ISSU2.Style("Width") = "80px"
        Me.IINSMAIL.Style("Width") = "200px"
        Me.EMAIL.Style("Width") = "200px"

        Me.sum01.Style("Width") = "50px"
        Me.sum02.Style("Width") = "50px"
        Me.txtTotal.Style("Width") = "50px"
        Me.txtCase.Style("Width") = "100px"
        Me.txtIrecunit.Style("Width") = "130px"
        Me.txtIREC.Style("Width") = "130px"
    End Sub

    Protected Overrides Sub AppInitMainRow(ByVal mainDataRow As System.Data.DataRow)

    End Sub

    ' 必須覆寫，將畫面資料存入 mainDataRow 裡面
    Protected Overrides Function AppCollectMainFormData(ByVal mainDataRow As DataRow) As Boolean
        Return Custom_AppCollectMainFormData(mainDataRow)
    End Function

    Private Function Custom_AppCollectMainFormData(ByVal mainDataRow As DataRow, Optional ACTFrom As String = "") As Boolean
        Dim iindex As Integer = StateMain.Tables("iqpissu").DefaultView.Count
        If CHK5_G.Text = "" Then
            If CHK5.Checked = True And iindex < 5 And Not checklock() Then 'Me.ScsOrgID <> "K105" Then
                Throw New Exception("集體費必須要5人以上")
            End If
        End If

        '套印文件(禁治產詢問)
        If rad_MARKX03_Y.Checked = False And rad_MARKX03_N.Checked = False Then
            Throw New Exception("請選擇傷害險醫療問項!!!")
        ElseIf rad_MARKX03_Y.Checked = True Then
            UTF.SetColumnText(mainDataRow, "MARKX03", "Y")
        ElseIf rad_MARKX03_N.Checked = True Then
            UTF.SetColumnText(mainDataRow, "MARKX03", "N")
        End If

        UTF.SetColumnText(mainDataRow, "MARKX04", "N")
        'End If
        If rad_MARKX05_Y.Checked = True Then
            Throw New Exception("茲被保險人受有監護宣告，請洽核保人員!")
        ElseIf rad_MARKX05_N.Checked = True Then
            UTF.SetColumnText(mainDataRow, "MARKX05", "N")
        End If

        If ACTFrom.ToLower() <> "savetemp_new" Then
            If iindex > 0 Then
            Else
                Throw New Exception("下方被保人資訊並無資料!!!")
            End If
        End If

        If sbErrAll <> "" Then
            Throw New Exception(sbErrAll)
        End If
        'txtTotal IT41 IT4Y IT11
        Dim TOTAMT As Integer = 0
        Dim TOTPREM As Integer = 0

        'IQPMAIN.TOTAMT= OBJ全部被保人的保額(IT11+IT41+IT4Y)加總。
        For Each dr As DataRow In StateMain.Tables("IQPOBJ").Select("instyp in ('IT41','IT4Y','IT11' , 'IQP01', 'IQP0Y', 'IQP02')")
            'drD.Delete()
            TOTAMT += CInt(dr("INSAMT"))
        Next

        For Each dr As DataRow In StateMain.Tables("IQPOBJ").Select("1=1")
            TOTPREM += CInt(dr("INSPREM"))
        Next

        '單張累計保額幫我開放至三億
        '另外其累計的幾付項目為IT41及IT4Y
        Dim TOTAMT_IT4Y4Y As Integer = 0
        Dim TOTAMT_IT41 As Integer = 0

        For Each dr_IT41 As DataRow In StateMain.Tables("IQPOBJ").Select(" instyp in ('IT41', 'IQP01')")
            'drD.Delete()
            TOTAMT_IT41 += CInt(dr_IT41("INSAMT"))

        Next

        Dim TOTAMT_IT4Y As Integer = 0

        For Each dr_IT4Y As DataRow In StateMain.Tables("IQPOBJ").Select(" instyp in ( 'IT4Y', 'IQP0Y')")
            'drD.Delete()
            TOTAMT_IT4Y += CInt(dr_IT4Y("INSAMT"))

        Next

        TOTAMT_IT4Y4Y = TOTAMT_IT41 + TOTAMT_IT4Y

        If TOTAMT_IT4Y4Y > 300000000 Then
            Throw New Exception("同一保單總保額不得超過 NT$ 3億 ")
        End If

        UTF.SetColumnText(mainDataRow, "IOFFICER", labIOFFICER.Text.Trim)
        UTF.SetColumnText(mainDataRow, "DPLY_BGN", UcRocDobjbgn.DateText)
        UTF.SetColumnText(mainDataRow, "DPLY_END", UcRocDobjend.DateText)
        UTF.SetColumnText(mainDataRow, "TPLY_BGN", txtTobjbgn.Text.Trim)
        UTF.SetColumnText(mainDataRow, "TPLY_END", txtTobjend.Text.Trim)

        UTF.SetColumnText(mainDataRow, "TrvCount", txtIQPITP_Days.Text.Trim)
        UTF.SetColumnText(mainDataRow, "NTRVSTART", TRAVEL_BEGIN.Text.Trim)

        TRAVEL_END.Text = Replace(TRAVEL_END.Text.Trim, "(申根區)", " ")
        UTF.SetColumnText(mainDataRow, "NTRVDEST", TRAVEL_END.Text)
        UTF.SetColumnText(mainDataRow, "NRMK7", "1") '旅行目的 1 2 3
        UTF.SetColumnText(mainDataRow, "NRMK5", TRAVEL_A.SelectedValue)
        'If rad_flight_Y.Checked = True Then
        '    UTF.SetColumnText(mainDataRow, "AITEM", "Y")
        'ElseIf rad_flight_N.Checked = True Then
        '    UTF.SetColumnText(mainDataRow, "AITEM", "N")
        'Else
        '    UTF.SetColumnText(mainDataRow, "AITEM", "")
        'End If


        'ITP費率檔(個人集體直接) 1啟用 0停用 
        If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

            If CHK5.Checked = True Then
                If TRAVEL_A.SelectedValue = "400" Then
                    UTF.SetColumnText(mainDataRow, "Ftrvarea", "5") '國外5 外6 別
                Else
                    UTF.SetColumnText(mainDataRow, "Ftrvarea", "6") '國外5 外6 別
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    UTF.SetColumnText(mainDataRow, "Ftrvarea", "1") '國內1 外2 別
                Else
                    UTF.SetColumnText(mainDataRow, "Ftrvarea", "2") '國內1 外2 別
                End If
            End If
        Else
            If TRAVEL_A.SelectedValue = "400" Then
                UTF.SetColumnText(mainDataRow, "Ftrvarea", "1") '國內1 外2 別
            Else
                UTF.SetColumnText(mainDataRow, "Ftrvarea", "2") '國內1 外2 別
            End If
        End If

        UTF.SetColumnText(mainDataRow, "Ipeople", iindex.ToString)
        UTF.SetColumnText(mainDataRow, "NINS", NAME_IINS.Text.Trim)

        '要保人編號,1=身分證， 2=統編，3=居留證
        UTF.SetColumnText(mainDataRow, "FinsCls", ddl_IINS.SelectedValue)

        If ddl_IINS.SelectedValue.ToString = "5" Then
            UTF.SetColumnText(mainDataRow, "FinsTyp", "7")

            '要保人國別
            UTF.SetColumnText(mainDataRow, "IINS_COUNTRY", ddlIINS_COUNTRY.SelectedValue)
        Else
            UTF.SetColumnText(mainDataRow, "FinsTyp", ddl_IINS2.SelectedValue)
        End If
        UTF.SetColumnText(mainDataRow, "IINS", txtIins_IINS2.Text.Trim)

        UTF.SetColumnText(mainDataRow, "INSBD", Birtdy_IINS.DateText)
        UTF.SetColumnText(mainDataRow, "IINS_ZIP", IINS_AREA.Text)
        UTF.SetColumnText(mainDataRow, "AINS", IINS_ADDR.Text.Trim)
        UTF.SetColumnText(mainDataRow, "NINS_TEL", IINS_TEL.Text.Trim)


        ''''是否索取電子保單
        '''If ChkMailYN.Checked = True Then
        '''    UTF.SetColumnText(mainDataRow, "FEP", "Y")
        '''Else
        '''    UTF.SetColumnText(mainDataRow, "FEP", "N")
        '''End If
        UTF.SetColumnText(mainDataRow, "FEP", "Y") '鎖定Y

        'IINSMAIL
        'INSMOBILE
        UTF.SetColumnText(mainDataRow, "IINSMAIL", IINSMAIL.Text.Trim)
        UTF.SetColumnText(mainDataRow, "INSMOBILE", INSMOBILE.Text.Trim)
        'NINS_CHARGER 負責人
        UTF.SetColumnText(mainDataRow, "NINS_CHARGER", txtPrincipal.Text.Trim)
        '通路內控資料
        UTF.SetColumnText(mainDataRow, "ICHANNEL", txtICHANNEL.Text.Trim)
        UTF.SetColumnText(mainDataRow, "IBIG_BR", txtIBIG_BR.Text.Trim)
        UTF.SetColumnText(mainDataRow, "IRECUNIT", txtIrecunit.Text.Trim)
        UTF.SetColumnText(mainDataRow, "NRECUNIT", txtNRECUNIT.Text.Trim)
        UTF.SetColumnText(mainDataRow, "IRECPESN", txtIRECPESN.Text.Trim)
        UTF.SetColumnText(mainDataRow, "NRECPESN", txtNrecpesn.Text.Trim)
        UTF.SetColumnText(mainDataRow, "IREC", txtIREC.Text)
        UTF.SetColumnText(mainDataRow, "IREC_OFFICR", txtIREC_OFFICR.Text.Trim)
        '團號 電傳人
        UTF.SetColumnText(mainDataRow, "TrlCod", txtTrlCod.Text.Trim)
        UTF.SetColumnText(mainDataRow, "TrlCal", txtTrlCal.Text.Trim)

        '抓主被保人
        Dim MNISSU As String = ""
        Dim MIISSU As String = ""
        Dim MIPLANNO As String = ""
        For Each dr_ISSU As DataRow In StateMain.Tables("IQPISSU").Select(" ISEQ = '1'")
            MIISSU = dr_ISSU("IISSU").ToString.Trim
            MNISSU = dr_ISSU("NISSU").ToString.Trim
            MIPLANNO = dr_ISSU("IPLANNO").ToString.Trim
        Next

        '因F115071600 改為被保人一併複製，故移除預設主被保人
        ''F115031101 資料暫存改為有[旅遊資訊]&[要保人資訊]即可暫存，故主被保人先預設為要保人
        'If ACTFrom.ToLower() = "savetemp_new" Then
        '    MIISSU = txtIins_IINS2.Text.Trim
        '    MNISSU = NAME_IINS.Text.Trim
        'End If

        UTF.SetColumnText(mainDataRow, "MNISSU", MNISSU)
        UTF.SetColumnText(mainDataRow, "MIISSU", MIISSU)

        '保費試算時MAIN_B_NRMK1再重給一次值（procChangedTargetProject()），因為同一要保單ITP與IQP轉換；MAIN_B_NRMK1也須調整
        If ddlTargetProject.SelectedValue = "ITP" Then
            'B2C(ITP的MAIN_B_NRMK1將會是3)
            'B2B ITP 的MAIN_B_NRMK1  則依照頁面費率別欄位進行變動，個人費率=2 團費率=3
            If CHK5.Checked = True Then
                UTF.SetColumnText(mainDataRow, "MAIN_B_NRMK1", "3") '集體投保否2=個人費率,3=團體費率			
            Else
                UTF.SetColumnText(mainDataRow, "MAIN_B_NRMK1", "2") '集體投保否2=個人費率,3=團體費率
            End If

            If (checkSpecialChannel().Item1 = True AndAlso (MIPLANNO = "J" Or MIPLANNO = "K" Or MIPLANNO = "L")) Or
               (checkChannelO().Item1 = True AndAlso (MIPLANNO = "J" Or MIPLANNO = "K" Or MIPLANNO = "L" Or MIPLANNO = "O")) Or
               (checkChannelMN().Item1 = True AndAlso (MIPLANNO = "M" Or MIPLANNO = "N")) Then
                UTF.SetColumnText(mainDataRow, "MAIN_B_NRMK1", "9")
            End If
        Else
            'B2C(ITP的MAIN_B_NRMK1將會是3)
            'B2B IQP 的MAIN_B_NRMK1  則依照頁面費率別欄位進行變動，個人費率=0 團費率=1
            If CHK5.Checked = True Then
                UTF.SetColumnText(mainDataRow, "MAIN_B_NRMK1", "1")
            Else
                UTF.SetColumnText(mainDataRow, "MAIN_B_NRMK1", "0")
            End If
        End If


        'IINSKIND
        UTF.SetColumnText(mainDataRow, "IINSKIND", ddlTargetProject.SelectedValue.Trim)
        UTF.SetColumnText(mainDataRow, "TOTAMT", TOTAMT.ToString) '總保額萬
        UTF.SetColumnText(mainDataRow, "TOTPREM", TOTPREM.ToString) '總保費元

        '國別
        UTF.SetColumnText(mainDataRow, "COUNTRY_A1", ddlCountry_A1.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_A2", ddlCountry_A2.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_B1", ddlCountry_B1.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_B2", ddlCountry_B2.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_C1", ddlCountry_C1.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_C2", ddlCountry_C2.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_D1", ddlCountry_D1.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_D2", ddlCountry_D2.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_E1", ddlCountry_E1.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_E2", ddlCountry_E2.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_F1", ddlCountry_F1.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_F2", ddlCountry_F2.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_G1", ddlCountry_G1.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_G2", ddlCountry_G2.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_H1", ddlCountry_H1.SelectedValue)
        UTF.SetColumnText(mainDataRow, "COUNTRY_H2", ddlCountry_H2.SelectedValue)

        If TRAVEL_A.SelectedValue = "750" Then
            UTF.SetColumnText(mainDataRow, "IS750", "1")
        Else
            UTF.SetColumnText(mainDataRow, "IS750", "0")
        End If

        UTF.SetColumnText(mainDataRow, "USERID", GetUserIDFromTextBox())

        Dim bolChanged As Boolean = False

        'If bolChanged Then
        bolChanged = bolChanged Or UTF.SetColumnText(mainDataRow, "Icreate", Me.ScsUserID)
        'bolChanged = bolChanged Or UTF.SetColumnText(mainDataRow, "Icreate", Me.txtCompanyid.Text.Trim & "\" & Me.txtUserid.Text.Trim)
        'bolChanged = bolChanged Or UTF.SetColumnText(mainDataRow, "Dcreate", Now.ToString("yyyy/MM/dd"))
        'End If
        Return bolChanged
    End Function

    Protected Overrides Function AppCollectDataGridItemData(ByVal currentDataGrid As DataGrid, ByVal currentItem As DataGridItem, ByVal currentRow As DataRow) As Boolean

        Dim bolChanged As Boolean = False
        Select Case currentDataGrid.ID

            Case dgdMain.ID

                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "ISEQ", CType(currentItem.FindControl("labISEQ"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "NISSU", CType(currentItem.FindControl("labNISSU"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "IISSU", CType(currentItem.FindControl("labIISSU"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "NISSUBD", CType(currentItem.FindControl("ucNISSUBD"), ucRocDate).DateText)

                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "IISSU_TP", CType(currentItem.FindControl("labIISSU_TP"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "IISSU_ZIP", CType(currentItem.FindControl("labIISSU_ZIP"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "AISSU", CType(currentItem.FindControl("labAISSU"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "NISSU_TEL", CType(currentItem.FindControl("labNISSU_TEL"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "NISSUMAIL", CType(currentItem.FindControl("labNISSUMAIL"), TextBox).Text)

                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "NINSREL", CType(currentItem.FindControl("labNINSREL"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "NAMEREL", CType(currentItem.FindControl("labNAMEREL"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "IDREL", CType(currentItem.FindControl("labIDREL"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "TYPEREL", CType(currentItem.FindControl("labTYPEREL"), TextBox).Text)

                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "NAMEREL2", CType(currentItem.FindControl("labNAMEREL2"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "IDREL2", CType(currentItem.FindControl("labIDREL2"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "TYPEREL2", CType(currentItem.FindControl("labTYPEREL2"), TextBox).Text)

                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "NAMEREL3", CType(currentItem.FindControl("labNAMEREL3"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "IDREL3", CType(currentItem.FindControl("labIDREL3"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "TYPEREL3", CType(currentItem.FindControl("labTYPEREL3"), TextBox).Text)
                'RELMOTHED
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "RELMOTHED", CType(currentItem.FindControl("labRELMOTHED"), TextBox).Text)
                'iPlanNo
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "iPlanNo", CType(currentItem.FindControl("labiPlanNo"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "TELREL", CType(currentItem.FindControl("labTELREL"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "ZIPREL", CType(currentItem.FindControl("labZIPREL"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "AREL", CType(currentItem.FindControl("labAREL"), TextBox).Text)
                'IISSU_SEX
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "IISSU_SEX", CType(currentItem.FindControl("labIISSU_SEX"), TextBox).Text)

            Case dgdMain2.ID

                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "ISEQ", CType(currentItem.FindControl("labISEQ"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "NISSU", CType(currentItem.FindControl("labNISSU"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "IISSU", CType(currentItem.FindControl("labIISSU"), TextBox).Text)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "NISSUBD", CType(currentItem.FindControl("ucNISSUBD"), ucRocDate).DateText)
                bolChanged = bolChanged Or UTF.SetColumnText(currentRow, "MDEDUCT", CType(currentItem.FindControl("MDEDUCT"), TextBox).Text)

        End Select
        Return bolChanged
    End Function


    Protected Overrides Function BaseGetCommandBarCommands() As UICommand()
        '選擇性覆寫:設定本網頁CommandBar所要使用的按鈕  
        'Return New UICommand() {UICommands.cmdNew, UICommands.cmdUpdate, UICommands.cmdQuery, New UICommand("PrintManager", "要保書列印"), New UICommand("PrintManager3", "名冊列印"), New UICommand("PrintManager2", "證明書列印"), New UICommand("PrintManager4", "申根簽證文件列印"), UICommands.cmdDelete, UICommands.cmdExit}

        ' Return New UICommand() {New UICommand("SaveTemp", "資料暫存"), New UICommand("SCOPY", "資料複製"), New UICommand("SaveCheck", "保費試算"), UICommands.cmdUpdate, UICommands.cmdQuery, New UICommand("PrintManager", "要保文件列印"), New UICommand("PrintManager2", "證明書列印"), New UICommand("PrintManager4", "申根簽證文件列印"), New UICommand("PayMoney", "多元化繳費平台"), New UICommand("PrintManager5", "線上付款"), UICommands.cmdDelete, New UICommand("MainClear", "清除"), UICommands.cmdExit}

        Return New UICommand() {New UICommand("SCOPY", "資料複製"), New UICommand("SaveTemp_New", "資料暫存"), New UICommand("UpdateTemp", "資料更新"), New UICommand("SaveCheck", "保費試算"), UICommands.cmdUpdate, New UICommand("SaveEdit", "修改"), UICommands.cmdQuery, New UICommand("Print_Combine", "列印整合"), New UICommand("PayMoney", "多元化繳費平台"), New UICommand("PrintManager5", "線上付款"), New UICommand("ModifyWhenY", "基本資料修改"), UICommands.cmdDelete, New UICommand("MainClear", "清除"), UICommands.cmdExit}
        'Return New UICommand() {New UICommand("SCOPY", "資料複製"), New UICommand("SaveTemp", "資料暫存"), New UICommand("SaveCheck", "保費試算"), UICommands.cmdUpdate, New UICommand("SaveEdit", "修改"), UICommands.cmdQuery, New UICommand("Print_Combine", "列印整合"), New UICommand("PayMoney", "多元化繳費平台"), New UICommand("PrintManager5", "線上付款"), UICommands.cmdDelete, New UICommand("MainClear", "清除"), UICommands.cmdExit}

    End Function
    Protected Overrides Function AppGetChildDataGrids(ByVal parentDataGrid As System.Web.UI.WebControls.DataGrid) As System.Web.UI.WebControls.DataGrid()
        If parentDataGrid Is Nothing Then
            Return New DataGrid() {dgdMain, dgdMain2, dgdFlight}
        End If
        Return New DataGrid() {}
    End Function

    Protected Overrides Function AppGetDataGridBoundDataView(ByVal currentDataGrid As System.Web.UI.WebControls.DataGrid, ByVal mainDataSet As System.Data.DataSet) As System.Data.DataView
        Select Case currentDataGrid.ID
            Case dgdMain.ID
                Dim dv As DataView = CType(mainDataSet, beITP101000).IQPISSU.DefaultView
                Me.BaseRestoreDataViewState(currentDataGrid.UniqueID, dv)
                Return dv
            Case dgdMain2.ID
                Dim dv As DataView = CType(mainDataSet, beITP101000).IQPOBJ.DefaultView
                Me.BaseRestoreDataViewState(currentDataGrid.UniqueID, dv)

                Return dv
            Case dgdFlight.ID
                Dim dv As DataView = CType(mainDataSet, beITP101000).HASM_FLIGHT_DETAIL.DefaultView
                Me.BaseRestoreDataViewState(currentDataGrid.UniqueID, dv)

                Return dv
            Case Else
                Throw New Exception("Invalid DataGrid!")
        End Select
    End Function

    Protected Overrides Function AppGetQueryParameterDataSet() As System.Data.DataSet

        Dim dsQuery As New beITP101000
        Dim row As DataRow = dsQuery.IQPMAIN.NewIQPMAINRow

        UTF.SetColumnText(row, "OID", Me.txtIQPMainOid.Text)
        dsQuery.IQPMAIN.AddIQPMAINRow(row)
        Return dsQuery

    End Function

    Protected Overrides Sub AppInitDataGridItemRow(ByVal currentDataGrid As System.Web.UI.WebControls.DataGrid, ByVal currentRow As System.Data.DataRow)
        Dim dr As DataRow = Me.BaseGetMainRow(Me.StateMain)
    End Sub


    Protected Overrides Sub BaseInitQueryForm()


    End Sub

    Protected Overrides Sub AppInitMainForm()
        '清空判斷[可保特殊方案(豪華型、標準型)]的變數，清空後會抓通路代碼重新判斷
        txtIsSPEChannel.Text = ""

        'lblTargetPlan.Text = String.Empty 
        ddlIT41.SelectedValue = ""
        ddlIT42.SelectedValue = ""
        ddlIT61.SelectedValue = ""
        ddlIT63.SelectedValue = ""
        ddlIT6E.SelectedValue = ""
        ddlIT44.SelectedValue = ""
        ddlIT4Y.SelectedValue = ""
        ddlIT43.SelectedValue = ""
        ddlIT64.SelectedValue = ""
        'txtCase.Text = ""
        sum02.Text = ""
        txtTotal.Text = ""
    End Sub

    Protected Overrides Function AppNewMainDataSet() As System.Data.DataSet
        Return New beITP101000
    End Function

    Protected Overrides Function AppNewQueryDataSet() As System.Data.DataSet
        Return New beITP101000
    End Function

    Protected Overrides ReadOnly Property BaseBrowseServiceName() As String
        Get
            Return "B2BService.bsBHAS101001Load"
        End Get
    End Property

    Protected Overrides ReadOnly Property AppQueryServiceName() As String
        Get
            Return "B2BService.bsBHAS101001Load"
        End Get
    End Property

    Protected Overrides ReadOnly Property AppSaveServiceName() As String
        Get
            Return "B2BService.bsBHAS101002Save"
        End Get
    End Property

    Protected Overrides Sub AppSetMainFormData(ByVal mainDataRow As System.Data.DataRow)

        txtIQPMainOid.Text = UTF.GetColumnText(mainDataRow, "oid")
        'ddlTargetProject.SelectedValue = UTF.GetColumnText(mainDataRow, "IINSKIND")
        Dim IQPMAINOID As String = ""
        Dim instyp As String = ""
        IQPMAINOID = txtIQPMainOid.Text.Trim
        If IQPMAINOID <> "" Then
            Dim dsSing As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"IQPOBJ", "IQPMAINOID ='" + IQPMAINOID + "' AND instyp in ('IT38')", "instyp"})

            If Not IsNothing(dsSing) AndAlso dsSing.Tables.Count > 0 AndAlso dsSing.Tables(0).Rows.Count > 0 Then
                'IT24 OR IT2F
                '20231020 改成IT38
                instyp = dsSing.Tables(0).Rows(0)("instyp").ToString.Trim
            Else
                '國內陸上行
            End If
        End If

        '日期先給，下方方案按鈕才有值可作業
        UcRocDobjbgn.DateText = UTF.GetColumnText(mainDataRow, "DPLY_BGN")
        UcRocDobjend.DateText = UTF.GetColumnText(mainDataRow, "DPLY_END")
        txtTobjbgn.Text = UTF.GetColumnText(mainDataRow, "TPLY_BGN")
        txtTobjend.Text = UTF.GetColumnText(mainDataRow, "TPLY_END")

        '若使用者查詢要保號，依照要保號給予ddlTargetProject為ITP or IQP，以此避免後續使用者欲修改要保人資料時，可能造成錯誤
        If Not String.IsNullOrEmpty(IQPMAINOID) Then
            ddlTargetProject.SelectedValue = UTF.GetColumnText(mainDataRow, "IINSKIND")

            '取主被保人的方案(IPlanNo)資料並給值，為了滿足當使用者查詢要保號後直接點﹝保費試算﹞
            Dim strMIISSU As String = UTF.GetColumnText(mainDataRow, "MIISSU")
            Dim dsMIISSU As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"IQPISSU", "IQPMAINOID ='" + IQPMAINOID + "' AND IISSU = '" + strMIISSU + "'"})
            Dim MiPlanNo As String = ""

            If Not IsNothing(dsMIISSU) AndAlso dsMIISSU.Tables.Count > 0 AndAlso dsMIISSU.Tables(0).Rows.Count > 0 Then
                MiPlanNo = dsMIISSU.Tables(0).Rows(0)("IPLANNO").ToString.Trim

                Me.txtCase.Text = GetiPlanNo(MiPlanNo.ToUpper())
                Me.lblTargetPlan.Text = Me.txtCase.Text

                Select Case MiPlanNo.ToUpper()
                    Case "A"
                        BT1_BUTTON("")
                    Case "C"
                        BT3_BUTTON("")
                    Case "D"
                        BT5_BUTTON("")
                    Case "J"
                        BT6_BUTTON("")
                    Case "K"
                        BT7_BUTTON("")
                    Case "L"
                        BT8_BUTTON("")
                    Case "M"
                        BT9_BUTTON("")
                    Case "N"
                        BT10_BUTTON("")
                    Case "O"
                        BT11_BUTTON("")
                End Select

            End If


        End If

        '要保人
        ddl_IINS.SelectedValue = UTF.GetColumnText(mainDataRow, "FinsCls")
        ddl_IINS2.SelectedValue = UTF.GetColumnText(mainDataRow, "FinsTyp")

        '20250620 移除此行，因為將CHK_P1改為都不顯示
        'If PassportShow.Text = "1" Then
        '	If ddl_IINS.SelectedValue = "5" Then
        '		CHK_P1.Checked = True
        '	End If
        'End If

        Dim IINS As String = UTF.GetColumnText(mainDataRow, "IINS")
        Me.txtIins_IINS2.Text = IINS
        If Not String.IsNullOrEmpty(IINS) Then
            txtIins_IINS2_TextChanged(Nothing, Nothing)
        End If
        Me.NAME_IINS.Text = UTF.GetColumnText(mainDataRow, "NINS")
        Birtdy_IINS.DateText = UTF.GetColumnText(mainDataRow, "INSBD")
        IINS_AREA.Text = UTF.GetColumnText(mainDataRow, "IINS_ZIP")
        IINS_ADDR.Text = UTF.GetColumnText(mainDataRow, "AINS")
        IINS_TEL.Text = UTF.GetColumnText(mainDataRow, "NINS_TEL")
        txtIQPITP_Days.Text = UTF.GetColumnText(mainDataRow, "TrvCount")
        TRAVEL_A.SelectedValue = UTF.GetColumnText(mainDataRow, "NRMK5")
        Check750.Text = TRAVEL_A.SelectedValue.Trim

        TRAVEL_BEGIN.Text = UTF.GetColumnText(mainDataRow, "NTRVSTART")
        TRAVEL_END.Text = UTF.GetColumnText(mainDataRow, "NTRVDEST")

        'IINSMAIL
        'INSMOBILE
        txtPESTATUS.Text = UTF.GetColumnText(mainDataRow, "PESTATUS")
        IINSMAIL.Text = UTF.GetColumnText(mainDataRow, "IINSMAIL")
        INSMOBILE.Text = UTF.GetColumnText(mainDataRow, "INSMOBILE")
        '通路內控資料
        txtICHANNEL.Text = UTF.GetColumnText(mainDataRow, "ICHANNEL")
        txtIBIG_BR.Text = UTF.GetColumnText(mainDataRow, "IBIG_BR")
        txtIrecunit.Text = UTF.GetColumnText(mainDataRow, "IRECUNIT")
        txtNRECUNIT.Text = UTF.GetColumnText(mainDataRow, "NRECUNIT")
        txtIRECPESN.Text = UTF.GetColumnText(mainDataRow, "IRECPESN")
        txtNrecpesn.Text = UTF.GetColumnText(mainDataRow, "NRECPESN")
        txtIREC.Text = UTF.GetColumnText(mainDataRow, "IREC")
        txtIREC_OFFICR.Text = UTF.GetColumnText(mainDataRow, "IREC_OFFICR")
        '團號 電傳人 
        txtTrlCod.Text = UTF.GetColumnText(mainDataRow, "TrlCod")
        txtTrlCal.Text = UTF.GetColumnText(mainDataRow, "TrlCal")
        'NINS_CHARGER 負責人
        txtPrincipal.Text = UTF.GetColumnText(mainDataRow, "NINS_CHARGER")
        txtMIISSU.Text = UTF.GetColumnText(mainDataRow, "MIISSU")
        txtMNIISSU.Text = UTF.GetColumnText(mainDataRow, "MNISSU")
        txtIQUO.Text = UTF.GetColumnText(mainDataRow, "IEST")
        txtPAY_NUM.Text = UTF.GetColumnText(mainDataRow, "PAY_NUM")
        txtIKYC.Text = UTF.GetColumnText(mainDataRow, "IKYC")

        'txtTOTPREM
        txtTOTPREM_NEW.Text = UTF.GetColumnText(mainDataRow, "TOTPREM")
        'FlightNum
        'If UTF.GetColumnText(mainDataRow, "AITEM") = "Y" Then
        '    rad_flight_Y.Checked = True
        'ElseIf UTF.GetColumnText(mainDataRow, "AITEM") = "N" Then
        '    rad_flight_N.Checked = True
        'End If
        'txtFlightNum.Text = UTF.GetColumnText(mainDataRow, "AITEM")
        'txtTOTPREM_NEW
        txtTOTPREM_NEW.Text = UTF.GetColumnText(mainDataRow, "TOTPREM")
        'txtPESTATUS_NAME

        '除了新增，其他狀態都不可勾選chkAssignACC
        chkAssignACC.Enabled = False

        Select Case txtPESTATUS.Text
            Case ""
                txtPESTATUS_NAME.Text = ""

                '特定通路才可以指定要保的通路代號與帳號
                If checkAssignACC().Item1 = True Then
                    chkAssignACC.Enabled = True
                    labChkMsg.InnerText = "確認通路代碼與帳號(若要調整請取消勾選)"
                    labChkMsg.Style("color") = "red"
                End If
                ShowPremInfoOrNot(False)
            Case "S"
                txtPESTATUS_NAME.Text = "[資料暫存]"
                ShowPremInfoOrNot(False)
            Case "W"
                txtPESTATUS_NAME.Text = "[送審中]"
                ShowPremInfoOrNot(True)
            Case "E"
                txtPESTATUS_NAME.Text = "[檢核中]"
                ShowPremInfoOrNot(True)
            Case "Y"
                txtPESTATUS_NAME.Text = "[要保確認]"
                ShowPremInfoOrNot(True)
            Case "R"
                txtPESTATUS_NAME.Text = "[要保取消]"
                ShowPremInfoOrNot(True)
            Case "C"
                txtPESTATUS_NAME.Text = "[已轉檔]"
                ShowPremInfoOrNot(True)
        End Select

        If UTF.GetColumnText(mainDataRow, "FEP") <> "" Then
            ddlChkMailYN.SelectedValue = "Y" 'UTF.GetColumnText(mainDataRow, "FEP")
        End If

        If UTF.GetColumnText(mainDataRow, "MAIN_B_NRMK1") = "3" OrElse UTF.GetColumnText(mainDataRow, "MAIN_B_NRMK1") = "1" Then
            CHK5.Checked = True
        Else
            CHK5.Checked = False
        End If

        '國別
        'A1
        ddlCountry_A1.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_A1")
        If ddlCountry_A1.SelectedValue <> "'" Then
            SetCountriesToDropDownList(ddlCountry_A1)
        End If
        ddlCountry_A2.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_A2")
        'B1
        ddlCountry_B1.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_B1")
        If ddlCountry_B1.SelectedValue <> "" Then
            SetCountriesToDropDownList(ddlCountry_B1)
        End If
        ddlCountry_B2.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_B2")
        'C1
        ddlCountry_C1.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_C1")
        If ddlCountry_C1.SelectedValue <> "" Then
            SetCountriesToDropDownList(ddlCountry_C1)
        End If
        ddlCountry_C2.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_C2")
        'D1
        ddlCountry_D1.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_D1")
        If ddlCountry_D1.SelectedValue <> "" Then
            SetCountriesToDropDownList(ddlCountry_D1)
        End If
        ddlCountry_D2.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_D2")
        'E1
        ddlCountry_E1.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_E1")
        If ddlCountry_E1.SelectedValue <> "" Then
            SetCountriesToDropDownList(ddlCountry_E1)
        End If
        ddlCountry_E2.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_E2")
        'F1
        ddlCountry_F1.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_F1")
        If ddlCountry_F1.SelectedValue <> "" Then
            SetCountriesToDropDownList(ddlCountry_F1)
        End If
        ddlCountry_F2.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_F2")
        'G1
        ddlCountry_G1.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_G1")
        If ddlCountry_G1.SelectedValue <> "" Then
            SetCountriesToDropDownList(ddlCountry_G1)
        End If
        ddlCountry_G2.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_G2")
        'H1
        ddlCountry_H1.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_H1")
        If ddlCountry_H1.SelectedValue <> "" Then
            SetCountriesToDropDownList(ddlCountry_H1)
        End If
        ddlCountry_H2.SelectedValue = UTF.GetColumnText(mainDataRow, "COUNTRY_H2")

        'IS750
        If UTF.GetColumnText(mainDataRow, "IS750") = "1" Then
            Checkbox2.Checked = True

        Else
            Checkbox2.Checked = False
        End If

        '舊資料查詢進來顯示 舊頁面 + 新頁面
        If UTF.GetColumnText(mainDataRow, "COUNTRY_A1") = "" Then
            Me.TRAVEL_hidden.Visible = True
            Me.ShowS1.Visible = True
        Else
            Me.TRAVEL_hidden.Visible = False
            Me.ShowS1.Visible = True
        End If

        '套印文件(禁治產詢問)
        If UTF.GetColumnText(mainDataRow, "MARKX03").ToString.Trim = "Y" Then
            rad_MARKX03_Y.Checked = True
        ElseIf UTF.GetColumnText(mainDataRow, "MARKX03").ToString.Trim = "N" Then
            rad_MARKX03_Y.Checked = False
        End If
        'F111122000拿掉此問項
        'If UTF.GetColumnText(mainDataRow, "MARKX04").ToString.Trim = "Y" Then
        '    rad_MARKX04_Y.Checked = True
        'ElseIf UTF.GetColumnText(mainDataRow, "MARKX04").ToString.Trim = "N" Then
        '    rad_MARKX04_Y.Checked = False
        'End If
        rad_MARKX04_Y.Checked = False

        If UTF.GetColumnText(mainDataRow, "MARKX05").ToString.Trim = "Y" Then
            rad_MARKX05_Y.Checked = True
        ElseIf UTF.GetColumnText(mainDataRow, "MARKX05").ToString.Trim = "N" Then
            rad_MARKX05_Y.Checked = False
        End If

        '外國人護照投保 給值
        If Not String.IsNullOrWhiteSpace(UTF.GetColumnText(mainDataRow, "FinsCls").ToString.Trim) Then
            ddl_IINS.SelectedValue = UTF.GetColumnText(mainDataRow, "FinsCls").ToString.Trim
        End If
        If Not String.IsNullOrWhiteSpace(UTF.GetColumnText(mainDataRow, "FinsTyp").ToString.Trim) Then
            ddl_IINS2.SelectedValue = UTF.GetColumnText(mainDataRow, "FinsTyp").ToString.Trim
        End If

        'F115030200 將要保帳號的資料填入txtCompany通路代碼、txtUserId帳號、labIOfficer經辦代號、txtIRECPESN_TxtIn登錄證號 
        labIOFFICER.Text = UTF.GetColumnText(mainDataRow, "IOFFICER").ToString.Trim
        txtIRECPESN_TxtIn.Text = UTF.GetColumnText(mainDataRow, "IRECPESN")

        Dim tempUserId = UTF.GetColumnText(mainDataRow, "USERID").ToString.Trim
        If tempUserId.Contains("\") Then
            txtCompanyid.Text = (tempUserId.Split("\"))(0)
            txtUserid.Text = (tempUserId.Split("\"))(1)
        End If

        If txtCompanyid.Text.ToUpper = "KEJ0" Then
            GetIrecItems(labIOFFICER.Text)
        End If

        '航班列表給值
        If StateMain.Tables.Contains("HASM_FLIGHT_DETAIL") Then
            BaseSetChildFormData(dgdFlight, StateMain)
        End If
    End Sub
    Protected Overrides ReadOnly Property BaseBrowseResultHiddenColumnNames() As String()
        Get
            Return New String() {"OID"}
        End Get
    End Property
    Private Sub dgdMain_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgdMain.DataBinding
        '選擇性覆寫:預設的實作是-如果目前正在編輯就隱藏所有的排序與刪除按鈕
        MyBase.BaseDataGridDataBinding(dgdMain, -1)
    End Sub

    Protected Overrides ReadOnly Property BaseIsBatchSave() As Boolean
        Get
            Return True
        End Get
    End Property

    Protected Overrides Function BaseGetMainTable(ByVal mainDataSet As System.Data.DataSet) As System.Data.DataTable
        Return mainDataSet.Tables(0)
    End Function
    ''事件
    Protected Overloads Overrides Sub BaseDoUICommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.CommandEventArgs)
        Try
            Select Case e.CommandName
                Case "Print_Combine"
                    '整合列印
                    Me.CallOpenModel("HASPrint_Combine_Model.aspx", "HASPrint_Combine_Model", "", "SELECT", New Object() {txtIQPMainOid.Text.Trim.HtmlEncode, "HAS101"})
                Case "New"
                    '暫存
                    Throw New Exception("系統錯誤，公會通報系統連線失敗，請再重新傳送NEW")

                Case UICommands.csUpdate
                    Try
                        Dim Q As String = Me.txtIQPMainOid.Text.Trim
                        If txtPESTATUS.Text = "W" Or txtPESTATUS.Text = "E" Then
                            Throw New Exception("資料送審中，無需點選[要保確認]")
                        End If

                        If txtPESTATUS.Text = "Y" Or txtPESTATUS.Text = "C" Then
                            Throw New Exception("已要保確認成功，無需再點選[要保確認]")
                        End If

                        '檢核帳號是否有誤
                        Dim dsAccount As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.account_check", Me.txtCompanyid.Text.Trim & "\" & Me.txtUserid.Text.Trim)
                        If dsAccount.Tables(0).DefaultView.Count > 0 Then
                            '有找到帳號
                        Else
                            Throw New Exception("[帳號]輸入錯誤，請確認")
                        End If

                        Dim iindex As Integer = dgdMain.Items.Count

                        If CHK5.Checked = True And iindex < 5 And Me.txtCompanyid.Text <> "K105" Then 'Not checklock() Then '
                            If CHK5_G.Text = "" Then
                                Throw New Exception("集體費必須要5人以上")
                            End If
                        End If

                        If INSMOBILE.Text.Trim = "" Then
                            Throw New Exception("索取電子保單 請輸入要保人行動電話.")
                        End If

                        If INSMOBILE.Text.Trim <> "" Then
                            If IsNumeric(INSMOBILE.Text.Trim) = True Then
                                If INSMOBILE.Text.Trim.Length <> 10 Then
                                    Throw New Exception("行動電話請輸入數字10碼.")
                                End If
                                If INSMOBILE.Text.Substring(0, 2).Trim <> "09" Then
                                    Throw New Exception("請輸入正確 行動電話 以方便簡訊傳送.")
                                End If
                            Else
                                Throw New Exception("行動電話請輸入數字.")
                            End If
                        End If

                        Dim checkValue As String = Request.Form("CheckSNNN")
                        If checkValue <> "檢核資料正確" Then
                            CheckSNNN.Text = ""
                            Throw New Exception("要保確認前，前請先點選[保費試算]")
                        End If

                        '檢核IQPMain & IQPOBJ總保額是否符合
                        Dim TotPremGood As String = checkTOTPREM()
                        If String.IsNullOrWhiteSpace(TotPremGood) = False Then
                            ErrMail("保費加總金額不符"， TotPremGood)
                            Throw New Exception("保費加總異常，請重新[保費試算]再執行。")
                        End If

                        '檢核IQPIISSU中【方案代號IPlanNo是否皆一致】 &【方案代號IPlanNo & IProjectNo是否符合】
                        '這個檢查放在checkOIDNIissu()、checkMAIN_B_NRMK()之前，先確保被保人IPlanNo皆一致
                        Dim ChkProNo As String = checkProjectNo()
                        If String.IsNullOrWhiteSpace(ChkProNo) = False Then
                            ErrMail("方案代號不符"， ChkProNo)
                            Throw New Exception("方案代號異常，請重新[保費試算]再執行，或通知資訊部同仁協助。")
                        End If

                        '檢核IQPIISSU方案是否與OID是否符合
                        Dim ChkPlanNoNOID As String = checkOIDNIissu()
                        If String.IsNullOrWhiteSpace(ChkPlanNoNOID) = False Then
                            ErrMail("方案代號與要保號不符"， ChkProNo)
                            Throw New Exception("方案代號與要保號不符，請重新登打此筆要保資料。")
                        End If

                        '檢核MAIN_B_NRMK與方案是否符合
                        Dim ChkB_NRMK As String = checkMAIN_B_NRMK()
                        If String.IsNullOrWhiteSpace(ChkB_NRMK) = False Then
                            ErrMail("MAIN_B_NRMK與方案不符"， ChkB_NRMK)
                            Throw New Exception("方案與要費率不符，請重新登打此筆要保資料，或通知資訊部同仁協助。")
                        End If

                        '檢核要被保人關係為本人，資料有無差異
                        '因為可能發生保費試算後，user改了要保人資料再要保確認。故新增此卡控
                        Dim ChkNINSREL As String = checkNINSREL()
                        If String.IsNullOrWhiteSpace(ChkNINSREL) = False Then
                            Throw New Exception(ChkNINSREL)
                        End If

                        '要保日期要小於今日 不能修改
                        If Me.UcRocDobjbgn.GetDate() < Today() Then
                            Throw New Exception("要保日期小於輸入日期[無法修改]")
                        End If


                        'If rad_Disease_Y.Checked And checkTotal_it62it63("IT62") > 0 Then
                        '	Throw New Exception("當勾選本次旅遊海外突發疾病(含傳染病)-是，則不得承保旅遊海外突發疾病(不含傳染病)，請重新選取被保險人報價。")
                        'ElseIf rad_Disease_Y.Checked = False And checkTotal_it62it63("IT63") > 0 Then
                        '	Throw New Exception("當勾選本次旅遊海外突發疾病(含傳染病)-否，則不得承保旅遊海外突發疾病(含傳染病)，請重新選取被保險人報價。")
                        'End If

                        '20250508 提出移除中國大陸地區限制
                        'If rad_Disease_Y.Checked And TRAVEL_END.Text.Contains("中國") Then
                        '	ddlIT63.SelectedValue = ""
                        '	Throw New Exception("選擇中國大陸地區時，不建議保﹝海外突發疾病保險金(含傳染病)﹞。")
                        'ElseIf TRAVEL_END.Text.Contains("中國") And checkTotal_it62it63("IT63") > 0 Then
                        '	Throw New Exception("選擇中國大陸地區時，不建議保﹝海外突發疾病保險金(含傳染病)﹞。")
                        'End If

                        Dim dsC As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.checkstatus", Me.txtIQPMainOid.Text.Trim)

                        If dsC.Tables("IQPMAIN").Rows.Count > 0 Then
                            If dsC.Tables("IQPMAIN").Rows(0).Item("PESTATUS") = "C" Then
                                Throw New Exception("資料已轉檔[無法修改]")
                            End If
                        End If

                        '風險與制裁檢核===
                        ''PrintLog(Me.ScsUserID + "-bsHASCheckXFILE_Task開始執行")
                        'Dim dsNew As DataSet = New DataSet
                        'Dim dt As DataTable = New DataTable
                        'Dim str1, str2 As String
                        'dt.Columns.Add("ICUS")
                        'dt.Columns.Add("NCUS")

                        'For Each drIQPISSU As DataRow In StateMain.Tables("IQPISSU").Rows
                        '	If drIQPISSU.RowState <> DataRowState.Deleted Then
                        '		Dim drNew As DataRow = dt.NewRow
                        '		drNew("ICUS") = drIQPISSU("IISSU")
                        '		drNew("NCUS") = drIQPISSU("NISSU")
                        '		dt.Rows.Add(drNew)
                        '	End If
                        'Next
                        'dsNew.Tables.Add(dt)

                        'Dim rds As DataSet = Me.CallBussinessService("HASService.bsHASCheckXFILE_Task", dsNew)
                        'If rds.Tables(0).Rows.Count > 0 Then
                        '	For i As Integer = 1 To rds.Tables(0).Rows.Count
                        '		str1 = rds.Tables(0).Rows(i - 1).Item("PASS").ToString.Trim
                        '		str2 = rds.Tables(0).Rows(i - 1).Item("MSG").ToString.Trim
                        '		If str1 = "N" Then
                        '			InsertUP_BTN.Text = "X"
                        '			'sbErrAll += "被保險人" & rds.Tables(0).Rows(0)("ICUS") & "需經過本公司審核,請洽詢 EMAIL:shnkonah@skinsurance.com.tw </br>" & Environment.NewLine
                        '			Throw New Exception("被保險人" & rds.Tables(0).Rows(i - 1)("ICUS") & "_" & rds.Tables(0).Rows(i - 1)("NCUS") & "需經過本公司審核,請洽詢 EMAIL:shnkonah@skinsurance.com.tw”)
                        '		End If
                        '	Next
                        'End If
                        '===風險與制裁檢核

                        If InsertUP_BTN.Text = "X" Then
                            Throw New Exception(sbErrAll)
                        End If

                        If StateMain.Tables("IQPMAIN").Rows(0)("IPEOPLE") <= 30 Then

                            Dim ds As New beHASB2CTransQuery
                            Dim errorMsg As Boolean = False
                            Dim ErrorStr As String = ""
                            Dim sbErr1 As String = ""

                            For Each drE As DataRow In StateMain.Tables("IQPISSU").Rows
                                If drE.RowState <> DataRowState.Deleted Then
                                    Dim dr As beHASB2CTransQuery.HASB2CTransQueryRow = ds.HASB2CTransQuery.NewHASB2CTransQueryRow
                                    dr = ds.HASB2CTransQuery.NewHASB2CTransQueryRow
                                    dr("IISSU") = drE("IISSU")
                                    dr("DPLY") = UcRocDobjbgn.DateText
                                    dr("DPLY2") = UcRocDobjend.DateText
                                    dr("IINSKIND") = "ITP"
                                    dr("HH_BGN") = txtTobjbgn.Text
                                    dr("HH_END") = txtTobjend.Text
                                    ds.HASB2CTransQuery.AddHASB2CTransQueryRow(dr)
                                End If
                            Next
                            '30筆以下才跑線上檢核======
                            Dim dsTmp As DataSet = CallService(ds, Page)

                            Dim wsSwitch As Boolean = True
                            '先透過萬用代碼檢核，公會服務卡控&告警信是否開啟
                            Try
                                Dim dsGood2Go As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_B2CTRANS' AND CODE ='N'"})

                                If Not IsNothing(dsGood2Go) AndAlso dsGood2Go.Tables.Count > 0 AndAlso dsGood2Go.Tables(0).Rows.Count > 0 Then
                                    wsSwitch = False
                                End If
                            Catch ex As Exception
                            End Try

                            '檢核公會查詢Service是否失敗，檢核Dataset的Table ﹝ERROR_MSG﹞欄位﹝Msg﹞是否有字串值，有則表示未順利呼叫公會查詢Service，寄送告警Email給慶帆、振宇 
                            Dim FunctionObj As New B2BGPL_Function
                            Dim svTup As Tuple(Of Boolean, String) = FunctionObj.CheckHASB2CTrans(dsTmp, Me.txtIQPMainOid.Text.Trim, "B2B(BHAS101003)錯誤_", wsSwitch)


                            If svTup.Item1 = False Then
                                errorMsg = True
                                ErrorStr += svTup.Item2
                            ElseIf dsTmp.Tables("HASB2C") Is Nothing Then
                                If wsSwitch Then errorMsg = True

                                ErrorStr += " 公會通報系統連線失敗，請再重新傳送!"
                            End If

                            Dim TOTAMT As Integer = 0
                            If dsTmp.Tables("HASB2C").Rows.Count > 0 Then
                                For i As Integer = 1 To dsTmp.Tables("HASB2C").Rows.Count
                                    sbErr1 = String.Empty
                                    Dim IISSU As String = dsTmp.Tables("HASB2C").Rows(i - 1).Item("IISSU").ToString.Trim '被保人ID

                                    Dim tot1 As Integer = 0
                                    For Each drG As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU & "' and instyp in ('IT41','IQP01')")
                                        tot1 = drG("INSAMT")
                                    Next
                                    '未滿15歲
                                    Dim tot2 As Integer = 0
                                    For Each drG As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU & "' and instyp in ('IT4Y', 'IQP0Y')")
                                        tot2 = drG("INSAMT")
                                    Next

                                    TOTAMT = tot1 + tot2

                                    Dim DieTravelSK As String = dsTmp.Tables("HASB2C").Rows(i - 1).Item("DieTravelSK").ToString.Trim  '旅平身故
                                    Dim DieSK As String = dsTmp.Tables("HASB2C").Rows(i - 1).Item("DieSK").ToString.Trim  '新光身故
                                    Dim DieTravelSameTrade As String = dsTmp.Tables("HASB2C").Rows(i - 1).Item("DieTravelSameTrade").ToString.Trim  '同業旅平身故
                                    Dim DieSameTrade As String = dsTmp.Tables("HASB2C").Rows(i - 1).Item("DieSameTrade").ToString.Trim  '同業身故
                                    Dim SameTradeCount As String = dsTmp.Tables("HASB2C").Rows(i - 1).Item("SameTradeCount").ToString.Trim  '同業家數
                                    Dim Funeral As String = dsTmp.Tables("HASB2C").Rows(i - 1).Item("Funeral").ToString.Trim  '未成年喪葬費用
                                    '被保險人三個月內密集向二家(含)以上投保者
                                    If CInt(DieTravelSK) > 0 Then
                                        errorMsg = True
                                        ErrorStr += "被保險人 " & IISSU & "   同一時段已投保旅平險!   "
                                        sbErr1 = "被保險人 " & IISSU & "   同一時段已投保旅平險!"
                                    End If
                                    '20240115 -- 2000萬改2500萬
                                    If SameTradeCount > 2 Or (CInt(TOTAMT) + CInt(DieTravelSameTrade)) > 25000000 Then
                                        errorMsg = True
                                        ErrorStr += "被保險人 " & IISSU & " 同業累計保險金已超過評估標準，建議不投保此保險。!!" &
                                                    "當前身故保額：" & TOTAMT & "；同業身故保額：" & DieTravelSameTrade & "   "
                                        sbErr1 = "被保險人 " & IISSU & " 同業累計保險金已超過評估標準，建議不投保此保險。!!" &
                                                    "當前身故保額：" & TOTAMT & "；同業身故保額：" & DieTravelSameTrade
                                    End If

                                    If tot1 = 0 Then
                                        '未滿15歲 有累積金額 
                                        If CInt(Funeral) > 0 Then
                                            '能保的金額 
                                            Dim CY As Integer = 0
                                            CY = 690000 - CInt(Funeral)
                                            If CY < 100000 Then
                                                If tot2 = 0 Then '喪葬保0元 直接過
                                                Else
                                                    errorMsg = True
                                                    ErrorStr += "被保險人 " & IISSU & "本次無法提供喪葬/失能保險金投保，建議不投保此保險   "
                                                    sbErr1 = "被保險人 " & IISSU & "本次無法提供喪葬/失能保險金投保，建議不投保此保險"
                                                End If
                                            Else
                                                If CInt(tot2) > CY Then
                                                    Dim money As String = ""
                                                    money = CY.ToString.Substring(0, 1)
                                                    If money = "6" Then
                                                        '最高只能50萬
                                                        money = "5"
                                                    End If
                                                    errorMsg = True
                                                    ErrorStr += "被保險人 " & IISSU & "喪葬/失能保險金，建議保額最高" & money & "0萬   "
                                                    sbErr1 = "被保險人 " & IISSU & "喪葬/失能保險金，建議保額最高" & money & "0萬"
                                                End If
                                            End If
                                        End If
                                    End If

                                    If Not String.IsNullOrEmpty(sbErr1) Then
                                        'IISSU 去 IQPISSU 尋找對應的資料列
                                        Dim foundRows() As DataRow = StateMain.Tables("IQPISSU").Select($"IISSU = '{IISSU}'")
                                        ' 如果有找到對應的資料列，就更新 ERR 欄位
                                        For Each dr As DataRow In foundRows
                                            dr("ERR") = sbErr1
                                        Next
                                    End If
                                Next
                            Else
                                If wsSwitch Then errorMsg = True

                                ErrorStr += "資料回傳錯誤，回傳人數為0 -- 公會通報系統連線失敗，請再重新傳送!!" & Environment.NewLine
                            End If

                            Dim StatusForLog As String = "1"
                            Dim LogStr As String = Me.txtIQPMainOid.Text.Trim + "_bsHASB2CTrans_服務呼叫成功"

                            If errorMsg Then
                                StatusForLog = "3"
                                LogStr = Me.txtIQPMainOid.Text.Trim + "_bsHASB2CTrans_" + ErrorStr

                                StateMain.AcceptChanges()
                                BaseSetChildFormData(Me.dgdMain, Me.StateMain)
                            End If

                            ' 寫呼叫bsHASB2CTrans的Log
                            Me.InvokeCommonService("B2BService.bsBRKSaveLog", New String() {
                               StatusForLog,
                               LogStr,
                               Me.ScsUserID,
                               (Me.ScsUserID.Split("\"c)(0)),
                               ds.HASB2CTransQuery.DataSet.GetXml(),
                               dsTmp.GetXml()
                            })
                            '========================新公會通報_IQP & L方案不需檢核========================
                            If ddlTargetProject.SelectedValue.Trim = "ITP" AndAlso lblTargetPlan.Text.Trim.Substring(0, 1) <> "L" AndAlso lblTargetPlan.Text.Trim.Substring(0, 1) <> "A" Then
                                Dim dsNew As New beHAS_TI_SumAssuredCheck
                                For Each drE As DataRow In StateMain.Tables("IQPISSU").Rows
                                    If drE.RowState <> DataRowState.Deleted Then
                                        Dim dr As beHAS_TI_SumAssuredCheck.QueryRow = dsNew.Query.NewQueryRow
                                        dr = dsNew.Query.NewQueryRow
                                        dr("IISSU") = drE("IISSU")
                                        dr("DPLY_BGN") = Convert.ToDateTime((UcRocDobjbgn.GetDate).ToString("yyyy-MM-dd") + " " + txtTobjbgn.Text + ":00") 'UcRocDobjbgn.DateText
                                        dr("DPLY_END") = Convert.ToDateTime((UcRocDobjend.GetDate).ToString("yyyy-MM-dd") + " " + txtTobjend.Text + ":00") 'UcRocDobjend.DateText
                                        dsNew.Query.AddQueryRow(dr)
                                    End If
                                Next

                                Dim ErrorStr2 As String

                                ErrorStr2 = FunctionObj.GetHAS_TI_SumAssuredCheck(dsNew, Me.txtIQPMainOid.Text.Trim)
                                If ErrorStr2 <> String.Empty Then
                                    errorMsg = True
                                    ErrorStr += ErrorStr2 & Environment.NewLine
                                End If
                            End If
                            '========================新公會通報======================== 


                            If errorMsg = True Then
                                Throw New Exception(ErrorStr)
                            End If

                        End If
                        '======30筆以下才跑線上檢核
                        '暫時拿掉
                        Try
                            If Me.txtCompanyid.Text = "K077" Then
                                Try
                                    CallBussinessServiceSKL_GETIREC()
                                Catch ex As Exception
                                    Throw New Exception("新光金取收件號錯誤!!!" & ex.Message)
                                End Try

                                ''通路檢核 BrokerCheck()
                                Dim BCheck As String
                                BCheck = BrokerCheck()
                                If BCheck <> String.Empty Then
                                    Throw New Exception(BCheck)
                                End If

                            End If
                        Catch ex As Exception
                            Throw New Exception("通路單位錯誤!!!" & ex.Message)
                        End Try

                        MyBase.BaseDoUICommand(sender, e)

                        If Me.UIIsErrorOccurs = False Then
                            Me.txtIQPMainOid.Text = Q
                            Me.BaseDoQuery()
                            If txtPESTATUS.Text.Trim = "W" Or txtPESTATUS.Text = "E" Then
                                'Me.UIAlertMessage("查詢作業繁忙 結果稍後由EMAIL通知")
                            ElseIf txtPESTATUS.Text.Trim = "Y" Then
                                'Me.UIAlertMessage("資料完成，同業累保已回覆")
                            End If
                            'Me.UIAlertMessage("「本案僅完成投保資料輸入，尚未完成核保作業，請於○○○年○○月○○日(保期起日)前將要保書回傳至本公司，逾期未回傳者，視同未投保成功，如發生理賠事故，本公司不負理賠之責。」")

                        End If

                    Catch ex As Exception
                        Me.ProcessException(ex)
                        'Me.UIAppendError(ex.Message.ToString)
                    End Try

                Case UICommands.csDeleteItem
                    Dim view As DataView = Me.AppGetDataGridBoundDataView(Me.dgdMain, Me.StateMain)
                    Dim rowView As DataRowView = view(0)
                    Dim row As DataRow = rowView.Row
                    If row("ISEQ") = "1" Then
                        'Throw New Exception("刪除資料後請點選[要保檢核]")
                    Else
                        MyBase.BaseDoUICommand(sender, e)
                    End If
                Case UICommands.csDelete
                    '刪除 改成 要保取消
                    'C.要保取消權限()
                    '1.生效日<系統日 and PESTATUS <> "C"
                    '3.主檔 IPRINT_Cef 證明書統計欄位=0
                    Try
                        Dim IQPMainOid As String = ""
                        Dim FunctionObj As New B2BGPL_Function

                        '要保取消後將狀態清空，避免未清空導致不合理操作
                        '（要保完成接著按要保取消，字串依然保持"檢核資料正確" >>隱憂>> user改日期(不按enter\tab)狀態依然保持在"檢核資料正確"，接著還可以按"要保完成"；且IKYC也不會順利取號）
                        CheckSNNN.Text = ""

                        If Me.txtIQPMainOid.Text.Trim = "" Then
                            Throw New Exception("無查詢編號，無法執行[要保取消]功能")
                        Else
                            IQPMainOid = Me.txtIQPMainOid.Text.Trim
                        End If

                        Dim dsC As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.CheckStatus", Me.txtIQPMainOid.Text.Trim)

                        If dsC.Tables("IQPMAIN").Rows(0).Item("PESTATUS") = "C" Then
                            Throw New Exception("資料已轉檔無法執行[要保取消]")
                        End If

                        '已繳費無法取消 除了鴻裕保代
                        '<========鴻裕保代特殊流程========
                        '取得HAS_SPECIAL權限
                        Dim K294 As String = ""
                        Dim dsSing0 As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_SPECIAL' AND CODE ='" + GetUserIDFromTextBox() + "'", "CODE,NAME"})

                        If Not IsNothing(dsSing0) AndAlso dsSing0.Tables.Count > 0 AndAlso dsSing0.Tables(0).Rows.Count > 0 Then
                            '鴻裕保代特殊權限 
                            K294 = "Y"
                        End If
                        '========鴻裕保代特殊流程========>

                        If K294 = "Y" Then
                            '繳費預設可以壓
                            ''要保日期要小於今日 不能刪除
                            If Me.UcRocDobjbgn.GetDate() <= Today() Then
                                Throw New Exception("要保日期小於輸入日期無法執行[要保取消]")
                            End If
                        Else
                            'SIGN_PAY_STATUS
                            '取得SIGN_PAY_STATUS繳費沒
                            Dim SIGN_PAY_STATUS As String = ""
                            Dim PESTATUS As String = ""
                            Dim dsSing As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"IQPMAIN", "OID  ='" + IQPMainOid + "'", "SIGN_PAY_STATUS,PESTATUS"})

                            If Not IsNothing(dsSing) AndAlso dsSing.Tables.Count > 0 AndAlso dsSing.Tables(0).Rows.Count > 0 Then
                                PESTATUS = dsSing.Tables(0).Rows(0)("PESTATUS").ToString.Trim
                                SIGN_PAY_STATUS = dsSing.Tables(0).Rows(0)("SIGN_PAY_STATUS").ToString.Trim

                                If SIGN_PAY_STATUS = "Y" Then
                                    Throw New Exception("已繳費無法執行[要保取消]")
                                End If
                            End If
                        End If

                        '要保取消用
                        Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.cancel", Me.txtIQPMainOid.Text.Trim)

                        '新公會取消收件通報。如果失敗則function會寄出告警信
                        If FunctionObj.GetbsHAS_TI_inform_R("BHAS101003", Me.txtIQPMainOid.Text.Trim, Me.txtCompanyid.Text.Trim & "\" & Me.txtUserid.Text.Trim, "CANCEL") = False Then
                            Throw New Exception("新公會取消收件通報失敗")
                        End If

                        '還要取消收件 bsHASCircularInterface
                        '參數只有三個()
                        'TYPE:         類型傳C()
                        'ICIRCULAR:    新產批號傳欲取消之批號() 060823A00001
                        'ICIRSEQ:      新產序號可傳可不傳(, 不傳時即為整份相同批號的資料去做取消)
                        Dim ICIRCULAR As String = "" '去抓ICIRCULAR

                        Dim dsRtn As DataSet = Me.InvokeCommonService("PUBService.bsSingleSQL", New String() {"IQPMAIN", "OID ='" + Me.txtIQPMainOid.Text.Trim + "'", "ICIRCULAR"})
                        If dsRtn.Tables(0).DefaultView.Count > 0 Then
                            ICIRCULAR = dsRtn.Tables(0).Rows(0)("ICIRCULAR")

                            Dim dsHAS As New PUBEntity.beHASCircular
                            Dim drMis As PUBEntity.beHASCircular.QueryRow

                            drMis = dsHAS.Query.NewQueryRow
                            drMis.TYPE = "C" '類型傳C()
                            drMis.ICIRCULAR = ICIRCULAR '通路類別
                            dsHAS.Query.AddQueryRow(drMis)

                            Dim dsFromMIS As DataSet = Me.CallBussinessService("HASService.bsHASCircularInterface_A", dsHAS)

                            '檢核Dataset是否有名為Message的table，有則表示[要保取消]失敗，寄送告警Email給文彥、育如、慶帆、振宇
                            Dim valCheck As Boolean

                            valCheck = FunctionObj.CheckCircular(dsFromMIS, Me.txtIQPMainOid.Text.Trim, "B2B(BHAS101003)錯誤_")
                            If valCheck = False Then
                                Throw New Exception("取消收件通報失敗")
                            End If
                        Else
                            Throw New Exception("取消收件通報失敗")
                        End If

                        BaseDoClear()
                        NOLockPage()
                        '點選[要保取消]後資料需保留並不是像以下畫面清空(9/25已通知修改)
                        txtIQPMainOid.Text = IQPMainOid
                        Me.BaseDoQuery()
                    Catch ex As Exception
                        Me.ProcessException(ex)
                        'Me.UIAppendError(ex.Message.ToString)
                    End Try

                Case "PayMoney"

                    If txtTOTPREM_NEW.Text < 70 Then
                        Throw New Exception("使用[多元化繳費平台]繳款金額 需大於70元")
                    End If

                    '查詢出來要有Q號 0006IRQ0F00001
                    Try
                        Dim IQUO As String
                        IQUO = txtIQUO.Text

                        PageSession("commandname") = "BHAS101002"
                        Dim url As String = ""
                        url = "BHAS301000.aspx?txtIQUO=" + IQUO
                        CallPage("BHAS301000.aspx", "BHAS301000", "", "SELECT", New Object() {IQUO})
                    Catch ex As Exception
                        Me.ProcessException(ex)
                    End Try

                Case "SaveTemp"
                    '資料暫存
                    Try
                        Dim iindex As Integer = dgdMain.Items.Count

                        '暫存時將狀態清空，避免未清空導致不合理操作
                        CheckSNNN.Text = ""

                        If StateMain.Tables("IQPISSU").DefaultView.Count > 0 Then
                        Else
                            Throw New Exception("下方被保人資訊並無資料，請點選[新增被保人/修改]按鈕")
                        End If

                        If CHK5.Checked = True And iindex < 5 And Me.txtCompanyid.Text <> "K105" Then 'Not checklock() Then '
                            If CHK5_G.Text = "" Then
                                Throw New Exception("集體費必須要5人以上")
                            End If
                        End If
                        If sbErrAll <> "" Then
                            Throw New Exception(sbErrAll)
                        End If

                        '檢核帳號是否有誤
                        Dim dsAccount As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.account_check", Me.txtCompanyid.Text.Trim & "\" & Me.txtUserid.Text.Trim)
                        If dsAccount.Tables(0).DefaultView.Count > 0 Then
                            '有找到帳號
                        Else
                            Throw New Exception("[帳號]輸入錯誤，請確認")
                        End If

                        Dim ChkIns As String
                        ChkIns = CheckExlUP()
                        If ChkIns <> String.Empty Then
                            Throw New Exception(ChkIns)
                        End If

                        Dim ChkTime As String
                        ChkTime = CheckTime()
                        If ChkTime <> String.Empty Then
                            Throw New Exception(ChkTime)
                        End If

                        Dim ChkAreaLock As String
                        ChkAreaLock = CheckAreaLock()
                        If ChkAreaLock <> String.Empty Then
                            Throw New Exception(ChkAreaLock)
                        End If

                        If ddlTargetProject.SelectedValue.Trim = "ITP" Then
                            Dim ChkTyphoon As String
                            ChkTyphoon = CheckTyphoon()
                            If ChkTyphoon <> String.Empty Then
                                Throw New Exception(ChkTyphoon)
                            End If
                        End If

                        Dim dr2 As DataRow = BaseGetMainRow(StateMain)
                        AppCollectMainFormData(dr2)

                        Dim errorMsg As Boolean = False
                        Dim ErrorStr As String = ""

                        If Me.txtCompanyid.Text = "KKSB" Or Me.txtCompanyid.Text = "KA77" Or Me.txtCompanyid.Text = "KKEW" Or Me.txtCompanyid.Text = "KEJ0" Or
                           Me.txtCompanyid.Text = "KKFM" Or Me.txtCompanyid.Text = "KKAQ" Or Me.txtCompanyid.Text = "KEN0" Or Me.txtCompanyid.Text = "KKEX" Or
                           Me.txtCompanyid.Text = "KKFP" Then
                            'If Me.ScsOrgID = "K077" Or Me.ScsOrgID = "KA77" Or Me.ScsOrgID = "KKEW" Or Me.ScsOrgID = "KEJ0" Or Me.ScsOrgID = "KKFM" Then
                            '通路檢核 BrokerCheck()
                            Dim BCheck As String
                            BCheck = BrokerCheck()
                            If BCheck <> String.Empty Then
                                Throw New Exception(BCheck)
                            End If
                        ElseIf bolK077() Then
                            If txtIrecunit.Text = "" Then
                                Throw New Exception("請輸入通路單位代號!!!")
                            End If

                        End If

                        If errorMsg = False Then
                            '暫存
                            Dim dsA As DataSet = Me.InvokeBusinessService("B2BService.bsBHAS101003SaveTemp", StateMain)
                            If dsA.Tables("IQPMAIN").Rows.Count > 0 Then
                                txtIQPMainOid.Text = dsA.Tables("IQPMAIN").Rows(0)("OID").ToString
                                Me.UIAppendMessage("暫存單號為：" + txtIQPMainOid.Text.ToString.Trim)
                                Me.BaseDoQuery()
                                IISSU_Clear()

                                '將 方案資訊、保費小計、保費合計 整欄隱藏
                                divfuncinfoDisplay(False)
                            End If
                        End If

                    Catch ex As Exception
                        Me.ProcessException(ex)
                        'Me.UIAppendError(ex.Message.ToString)
                    End Try
                Case "SaveTemp_New"
                    '資校暫存_新 F115031101
                    Try
                        Dim iindex As Integer = dgdMain.Items.Count

                        '暫存時將狀態清空，避免未清空導致不合理操作
                        CheckSNNN.Text = ""

                        If CHK5.Checked = True And iindex < 5 And Me.txtCompanyid.Text <> "K105" Then 'Not checklock() Then '
                            If CHK5_G.Text = "" Then
                                Throw New Exception("集體費必須要5人以上")
                            End If
                        End If

                        '檢核帳號是否有誤
                        Dim dsAccount As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.account_check", Me.txtCompanyid.Text.Trim & "\" & Me.txtUserid.Text.Trim)
                        If dsAccount.Tables(0).DefaultView.Count > 0 Then
                            '有找到帳號
                        Else
                            Throw New Exception("[帳號]輸入錯誤，請確認")
                        End If

                        '旅遊地區整合
                        Dim COUNCODE_ALL As String = ""
                        COUNCODE_ALL = CHECK_COUNCODE1()
                        TRAVEL_A.SelectedValue = COUNCODE_ALL

                        Dim ChkIns As String
                        ChkIns = CheckExlUP()
                        If ChkIns <> String.Empty Then
                            Throw New Exception(ChkIns)
                        End If

                        Dim ChkTime As String
                        ChkTime = CheckTime()
                        If ChkTime <> String.Empty Then
                            Throw New Exception(ChkTime)
                        End If

                        Dim ChkAreaLock As String
                        ChkAreaLock = CheckAreaLock()
                        If ChkAreaLock <> String.Empty Then
                            Throw New Exception(ChkAreaLock)
                        End If

                        If ddlTargetProject.SelectedValue.Trim = "ITP" Then
                            Dim ChkTyphoon As String
                            ChkTyphoon = CheckTyphoon()
                            If ChkTyphoon <> String.Empty Then
                                Throw New Exception(ChkTyphoon)
                            End If
                        End If

                        Dim dr2 As DataRow = BaseGetMainRow(StateMain)
                        Custom_AppCollectMainFormData(dr2, e.CommandName)

                        Dim errorMsg As Boolean = False
                        Dim ErrorStr As String = ""
                        If Me.txtCompanyid.Text = "KKSB" Or Me.txtCompanyid.Text = "KA77" Or Me.txtCompanyid.Text = "KKEW" Or Me.txtCompanyid.Text = "KEJ0" Or
                           Me.txtCompanyid.Text = "KKFM" Or Me.txtCompanyid.Text = "KKAQ" Or Me.txtCompanyid.Text = "KEN0" Or Me.txtCompanyid.Text = "KKEX" Or
                           Me.txtCompanyid.Text = "KKFP" Then
                            'If Me.ScsOrgID = "K077" Or Me.ScsOrgID = "KA77" Or Me.ScsOrgID = "KKEW" Or Me.ScsOrgID = "KEJ0" Or Me.ScsOrgID = "KKFM" Then
                            '通路檢核 BrokerCheck()
                            Dim BCheck As String
                            BCheck = BrokerCheck()
                            If BCheck <> String.Empty Then
                                Throw New Exception(BCheck)
                            End If
                        ElseIf bolK077() Then
                            If txtIrecunit.Text = "" Then
                                Throw New Exception("請輸入通路單位代號!!!")
                            End If

                        End If

                        If errorMsg = False Then
                            '暫存
                            Dim dsA As DataSet = Me.InvokeBusinessService("B2BService.bsBHAS101003SaveTemp", StateMain)
                            If dsA IsNot Nothing AndAlso dsA.Tables("IQPMAIN") IsNot Nothing AndAlso dsA.Tables("IQPMAIN").Rows.Count > 0 Then
                                txtIQPMainOid.Text = dsA.Tables("IQPMAIN").Rows(0)("OID").ToString
                                Me.UIAppendMessage("暫存單號為：" + txtIQPMainOid.Text.ToString.Trim)
                                Me.BaseDoQuery()
                                IISSU_Clear()

                                '將 方案資訊、保費小計、保費合計 整欄隱藏
                                divfuncinfoDisplay(False)
                            End If
                        End If

                    Catch ex As Exception
                        Me.ProcessException(ex)
                        'Me.UIAppendError(ex.Message.ToString)
                    End Try
                Case "UpdateTemp"
                    Try
                        Dim iindex As Integer = dgdMain.Items.Count

                        '暫存時將狀態清空，避免未清空導致不合理操作
                        CheckSNNN.Text = ""

                        If StateMain.Tables("IQPISSU").DefaultView.Count > 0 Then
                        Else
                            Throw New Exception("下方被保人資訊並無資料，請點選[新增被保人/修改]按鈕")
                        End If

                        If CHK5.Checked = True And iindex < 5 And Me.txtCompanyid.Text <> "K105" Then 'Not checklock() Then '
                            If CHK5_G.Text = "" Then
                                Throw New Exception("集體費必須要5人以上")
                            End If
                        End If

                        If String.IsNullOrEmpty(Me.txtIQPMainOid.Text.Trim) Then
                            MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp_New", "").Visible = True
                            Throw New Exception("請先點選[資料暫存]取得要保號。")
                        End If

                        If sbErrAll <> "" Then
                            Throw New Exception(sbErrAll)
                        End If

                        '檢核帳號是否有誤
                        Dim dsAccount As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.account_check", Me.txtCompanyid.Text.Trim & "\" & Me.txtUserid.Text.Trim)
                        If dsAccount.Tables(0).DefaultView.Count > 0 Then
                            '有找到帳號
                        Else
                            Throw New Exception("[帳號]輸入錯誤，請確認")
                        End If

                        Dim ChkIns As String
                        ChkIns = CheckExlUP()
                        If ChkIns <> String.Empty Then
                            Throw New Exception(ChkIns)
                        End If

                        Dim ChkTime As String
                        ChkTime = CheckTime()
                        If ChkTime <> String.Empty Then
                            Throw New Exception(ChkTime)
                        End If

                        Dim ChkAreaLock As String
                        ChkAreaLock = CheckAreaLock()
                        If ChkAreaLock <> String.Empty Then
                            Throw New Exception(ChkAreaLock)
                        End If

                        If ddlTargetProject.SelectedValue.Trim = "ITP" Then
                            Dim ChkTyphoon As String
                            ChkTyphoon = CheckTyphoon()
                            If ChkTyphoon <> String.Empty Then
                                Throw New Exception(ChkTyphoon)
                            End If
                        End If

                        Dim dr2 As DataRow = BaseGetMainRow(StateMain)
                        AppCollectMainFormData(dr2)

                        Dim errorMsg As Boolean = False
                        Dim ErrorStr As String = ""

                        If Me.txtCompanyid.Text = "KKSB" Or Me.txtCompanyid.Text = "KA77" Or Me.txtCompanyid.Text = "KKEW" Or Me.txtCompanyid.Text = "KEJ0" Or
                           Me.txtCompanyid.Text = "KKFM" Or Me.txtCompanyid.Text = "KKAQ" Or Me.txtCompanyid.Text = "KEN0" Or Me.txtCompanyid.Text = "KKEX" Or
                           Me.txtCompanyid.Text = "KKFP" Then
                            'If Me.ScsOrgID = "K077" Or Me.ScsOrgID = "KA77" Or Me.ScsOrgID = "KKEW" Or Me.ScsOrgID = "KEJ0" Or Me.ScsOrgID = "KKFM" Then
                            '通路檢核 BrokerCheck()
                            Dim BCheck As String
                            BCheck = BrokerCheck()
                            If BCheck <> String.Empty Then
                                Throw New Exception(BCheck)
                            End If
                        ElseIf bolK077() Then
                            If txtIrecunit.Text = "" Then
                                Throw New Exception("請輸入通路單位代號!!!")
                            End If

                        End If

                        If errorMsg = False Then
                            '暫存
                            'F115031101需求單允許〔資料暫存〕&〔資料複製〕為 無被保資料&無方案狀態（僅旅遊資訊&要保人資訊），
                            '先使用ITP當作暫存OID。但後續使用者可能要保IQP，故判斷方案是否與當前OID相同；否則走SaveTemp取新要保號
                            Dim dsA As DataSet = New DataSet 'Me.InvokeBusinessService("B2BService.bsBHAS101003SaveEdit", StateMain)

                            If GetTargetProject().Substring(0, 2).ToUpper = Me.txtIQPMainOid.Text.Substring(0, 2).ToUpper Then
                                dsA = Me.InvokeBusinessService("B2BService.bsBHAS101003SaveEdit", StateMain)
                                Me.UIAppendMessage("更新成功")
                            Else
                                ' OID與 [新增被保人方案] 不同，IINSKIND不會更新，在此將IINSKIND歸正，SaveTemp取號才正確。
                                StateMain.Tables("IQPMAIN").Rows(0)("IINSKIND") = GetTargetProject().ToUpper
                                StateMain.AcceptChanges()

                                dsA = Me.InvokeBusinessService("B2BService.bsBHAS101003SaveTemp", StateMain)

                                If dsA IsNot Nothing AndAlso dsA.Tables("IQPMAIN") IsNot Nothing AndAlso dsA.Tables("IQPMAIN").Rows.Count > 0 Then
                                    txtIQPMainOid.Text = dsA.Tables("IQPMAIN").Rows(0)("OID").ToString
                                    Me.UIAppendMessage("暫存單號為：" + txtIQPMainOid.Text.ToString.Trim)
                                End If

                            End If

                            If dsA IsNot Nothing AndAlso dsA.Tables("IQPMAIN") IsNot Nothing AndAlso dsA.Tables("IQPMAIN").Rows.Count > 0 Then
                                StateMain = dsA
                                StateMain.AcceptChanges()
                                BaseSetChildFormData(Me.dgdMain, Me.StateMain)
                                BaseSetFormData(StateMain)

                                'Me.BaseDoQuery()
                                IISSU_Clear()

                                '將 方案資訊、保費小計、保費合計 整欄隱藏
                                divfuncinfoDisplay(False)
                            Else
                                Throw New Exception("資料更新發生異常，請重新操作")
                            End If
                        End If

                    Catch ex As Exception
                        Me.ProcessException(ex)
                        'Me.UIAppendError(ex.Message.ToString)
                    End Try
                Case "SaveEdit"
                    '修改 
                    Try
                        Dim iindex As Integer = dgdMain.Items.Count

                        If StateMain.Tables("IQPISSU").DefaultView.Count > 0 Then
                        Else
                            Throw New Exception("下方被保人資訊並無資料，請點選[新增被保人/修改]按鈕")
                        End If

                        If CHK5.Checked = True And iindex < 5 And Me.txtCompanyid.Text <> "K105" Then 'Not checklock() Then '
                            If CHK5_G.Text = "" Then
                                Throw New Exception("集體費必須要5人以上")
                            End If
                        End If
                        If sbErrAll <> "" Then
                            Throw New Exception(sbErrAll)
                        End If

                        '檢核帳號是否有誤
                        Dim dsAccount As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.account_check", Me.txtCompanyid.Text.Trim & "\" & Me.txtUserid.Text.Trim)
                        If dsAccount.Tables(0).DefaultView.Count > 0 Then
                            '有找到帳號
                        Else
                            Throw New Exception("[帳號]輸入錯誤，請確認")
                        End If

                        Dim ChkIns As String
                        ChkIns = CheckExlUP()
                        If ChkIns <> String.Empty Then
                            Throw New Exception(ChkIns)
                        End If

                        Dim ChkTime As String
                        ChkTime = CheckTime()
                        If ChkTime <> String.Empty Then
                            Throw New Exception(ChkTime)
                        End If

                        Dim ChkAreaLock As String
                        ChkAreaLock = CheckAreaLock()
                        If ChkAreaLock <> String.Empty Then
                            Throw New Exception(ChkAreaLock)
                        End If

                        If ddlTargetProject.SelectedValue.Trim = "ITP" Then
                            Dim ChkTyphoon As String
                            ChkTyphoon = CheckTyphoon()
                            If ChkTyphoon <> String.Empty Then
                                Throw New Exception(ChkTyphoon)
                            End If
                        End If

                        Dim dr2 As DataRow = BaseGetMainRow(StateMain)
                        AppCollectMainFormData(dr2)

                        Dim errorMsg As Boolean = False
                        Dim ErrorStr As String = ""

                        If Me.txtCompanyid.Text = "KKSB" Or Me.txtCompanyid.Text = "KA77" Or Me.txtCompanyid.Text = "KKEW" Or Me.txtCompanyid.Text = "KEJ0" Or
                           Me.txtCompanyid.Text = "KKFM" Or Me.txtCompanyid.Text = "KKAQ" Or Me.txtCompanyid.Text = "KEN0" Or Me.txtCompanyid.Text = "KKEX" Or
                           Me.txtCompanyid.Text = "KKFP" Then
                            '通路檢核 BrokerCheck()
                            Dim BCheck As String
                            BCheck = BrokerCheck()
                            If BCheck <> String.Empty Then
                                Throw New Exception(BCheck)
                            End If
                        ElseIf bolK077() Then
                            If txtIrecunit.Text = "" Then
                                Throw New Exception("請輸入通路單位代號!!!")
                            End If

                        End If

                        If errorMsg = False Then
                            '修改
                            Dim dsA As DataSet = Me.InvokeBusinessService("B2BService.bsBHAS101003SaveEdit", StateMain)
                            If dsA.Tables("IQPMAIN").Rows.Count > 0 Then
                                txtIQPMainOid.Text = dsA.Tables("IQPMAIN").Rows(0)("OID").ToString
                                Me.BaseDoQuery()
                                Me.UIAppendMessage("資料修改")
                            End If
                        End If
                    Catch ex As Exception
                        Me.ProcessException(ex)
                    End Try
                Case "SCOPY"
                    Try
                        'F115031101 重啟複製功能 僅能複製90天內的案件
                        Dim IQPMAINOID_Copy As String = Me.txtIQPMainOid.Text.Trim
                        If String.IsNullOrEmpty(IQPMAINOID_Copy) = False Then
                            Dim dsSing As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"IQPMAIN", "OID  ='" + IQPMAINOID_Copy + "'", "DCREATE"})
                            If Not IsNothing(dsSing) AndAlso dsSing.Tables.Count > 0 AndAlso dsSing.Tables(0).Rows.Count > 0 Then
                                Dim dcreate_Copy As DateTime = Convert.ToDateTime(dsSing.Tables(0).Rows(0)("DCREATE").ToString.Trim)
                                Dim dNow As DateTime = DateTime.Now
                                Dim diff As TimeSpan = dNow - dcreate_Copy
                                If diff.Days > 90 Then
                                    Throw New Exception("無法複製90天前的案件。")
                                End If
                            Else
                                Throw New Exception("要保號取資料失敗，請重試。")
                            End If
                        Else
                            Throw New Exception("無要保號可以複製!")
                        End If

                        NOLockPage()
                        Dim iindex As Integer = dgdMain.Items.Count


                        If IsDBNull(StateMain.Tables("IQPMAIN").Rows(0).Item("INSBD")) = True Then
                        Else
                            Birtdy_IINS.DateText = StateMain.Tables("IQPMAIN").Rows(0).Item("INSBD")
                        End If

                        '清除總保費\總保額
                        StateMain.Tables("IQPMAIN").Rows(0).Item("TOTAMT") = ""
                        StateMain.Tables("IQPMAIN").Rows(0).Item("TOTPREM") = ""
                        StateMain.Tables("IQPMAIN").Rows(0).Item("IINSKIND") = "ITP" '複製後預設取號為ITP，後續被保人改方案方便處理。(被保人IPLANNO要為空)

                        '清除被保人方案
                        For Each drE As DataRow In StateMain.Tables("IQPISSU").Rows
                            If drE.RowState <> DataRowState.Deleted Then
                                drE("ISSU_4") = ""
                                drE("IPLANNO") = ""
                                drE("IPROJECTNO") = ""
                            End If
                        Next


                        ''清除IQPISSU & IQPOBJ資料
                        'StateMain.Tables("IQPISSU").Clear()
                        StateMain.Tables("IQPOBJ").Clear()
                        StateMain.Tables("HASM_FLIGHT_DETAIL").Clear()
                        StateMain.AcceptChanges()
                        'Dim dr2 As DataRow = BaseGetMainRow(StateMain)
                        'AppCollectMainFormData(dr2)

                        Dim errorMsg As Boolean = False
                        Dim ErrorStr As String = ""

                        If errorMsg = False Then
                            '複製
                            Dim dsA As DataSet = Me.InvokeBusinessService("B2BService.bsBHAS101003SaveTemp", StateMain)
                            If dsA.Tables("IQPMAIN").Rows.Count > 0 Then
                                txtIQPMainOid.Text = dsA.Tables("IQPMAIN").Rows(0)("OID").ToString
                                UIFormIsDirty = False
                                Me.BaseDoQuery()
                                Me.UIAppendMessage("資料複製")

                                txtISEQ.Text = ""
                                txtIISSU_Ori.Text = ""
                                lblTargetPlan.Text = ""
                                txtCase.Text = ""
                                txtTOTPREM.Text = ""

                                MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveCheck", "").Visible = False
                                MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp_New", "").Visible = False
                            End If
                        End If
                    Catch ex As Exception
                        Me.ProcessException(ex)
                    End Try
                Case "SaveCheck"
                    '保費試算
                    Try
                        hdValidateFlag.Value = "0"
                        Dim TargetProject As String = ddlTargetProject.SelectedValue.Trim '20250320
                        Dim strErr As String = ""
                        Me.IsInputErr = False

                        '檢核指定保經代，應輸入﹝通路人員名稱(登錄證字號)﹞，指定的保經代紀錄在萬用代碼：
                        Dim IRECPESNMsg As String
                        IRECPESNMsg = checkIRECPESN()
                        If IRECPESNMsg <> String.Empty Then
                            Throw New Exception(IRECPESNMsg)
                        End If

                        '檢核要、被保人是否使用護照投保國內旅遊
                        Dim ChkInsRe As String
                        ChkInsRe = CheckBasicRe()
                        If ChkInsRe <> String.Empty Then
                            Throw New Exception(ChkInsRe)
                        End If

                        '旅遊地區整合
                        Dim COUNCODE_ALL As String = ""
                        COUNCODE_ALL = CHECK_COUNCODE1()
                        TRAVEL_A.SelectedValue = COUNCODE_ALL

                        If TRAVEL_A.SelectedValue = "400" AndAlso StateMain.Tables("IQPISSU").Rows.Count < 2 Then
                            '20240702 國內只要人數不過就直接踢掉 不要再往下走
                            Throw New Exception("「國內地區」僅受理投保人數2人(含)以上要保作業！")
                        End If

                        ''檢查所有被保人，海突含傳染&不含傳染的投保選擇是否一致
                        'If rad_Disease_Y.Checked = True And checkTotal_it62it63("IT62") > 0 Then
                        '	Throw New Exception("當勾選本次旅遊海外突發疾病(含傳染病)-是，則不得承保旅遊海外突發疾病(不含傳染病)，請重新選取被保險人報價。")
                        'ElseIf rad_Disease_N.Checked = True And checkTotal_it62it63("IT63") > 0 Then
                        '	Throw New Exception("當勾選本次旅遊海外突發疾病(含傳染病)-否，則不得承保旅遊海外突發疾病(含傳染病)，請重新選取被保險人報價。")
                        'End If

                        '20250508 提出移除中國大陸地區限制
                        'If rad_Disease_Y.Checked And TRAVEL_END.Text.Contains("中國") Then
                        '	ddlIT63.SelectedValue = ""
                        '	strErr &= "選擇中國大陸地區時，不建議保﹝海外突發疾病保險金(含傳染病)﹞。" & Environment.NewLine
                        'ElseIf TRAVEL_END.Text.Contains("中國") And checkTotal_it62it63("IT63") > 0 Then
                        '	strErr &= "選擇中國大陸地區時，不建議保﹝海外突發疾病保險金(含傳染病)﹞。" & Environment.NewLine
                        'End If

                        Dim ChkTime As String = ""
                        ChkTime = CheckTime()
                        If ChkTime <> String.Empty Then
                            Throw New Exception(ChkTime)
                        End If
                        'End If

                        Dim ChkAreaLock As String
                        ChkAreaLock = CheckAreaLock()
                        If ChkAreaLock <> String.Empty Then
                            Throw New Exception(ChkAreaLock)
                        End If

                        Dim ChkFlghit As String
                        ChkFlghit = CheckAllowFlightOrNot()
                        If ChkFlghit <> String.Empty Then
                            Throw New Exception(ChkFlghit)
                        End If

                        If ddlTargetProject.SelectedValue.Trim = "ITP" Then
                            Dim ChkTyphoon As String
                            ChkTyphoon = CheckTyphoon()
                            If ChkTyphoon <> String.Empty Then
                                Throw New Exception(ChkTyphoon)
                            End If
                        End If

                        ''檢查日本JP720出發日只能在投保日後2~30天內投保---------
                        'Dim Checkdate As Boolean = CheckAreaDply_bgn()
                        'If Checkdate Then
                        '    Throw New Exception("旅綜險若投保日本地區，出發日須於登打日+30個日曆日內。")
                        'End If

                        '---------

                        GetDay() '計算天數


                        Dim ds As New beHASB2CTransQuery
                        Dim errorMsg As Boolean = False
                        Dim ErrorStr As String = ""


                        For Each drE As DataRow In StateMain.Tables("IQPISSU").Rows
                            If drE.RowState <> DataRowState.Deleted Then
                                Dim dr As beHASB2CTransQuery.HASB2CTransQueryRow = ds.HASB2CTransQuery.NewHASB2CTransQueryRow
                                dr = ds.HASB2CTransQuery.NewHASB2CTransQueryRow
                                dr("IISSU") = drE("IISSU")
                                dr("DPLY") = UcRocDobjbgn.DateText
                                dr("DPLY2") = UcRocDobjend.DateText
                                dr("IINSKIND") = TargetProject '20250320
                                dr("HH_BGN") = txtTobjbgn.Text
                                dr("HH_END") = txtTobjend.Text
                                ds.HASB2CTransQuery.AddHASB2CTransQueryRow(dr)
                            End If
                        Next

                        Dim dsReturn As DataSet = New DataSet()
                        Try
                            ''重複投保-----------------------------------
                            dsReturn = Me.InvokeBusinessService("B2BService.bsBHAS101002CheckTime", ds)
                        Catch ex As Exception
                            Throw New Exception("系統服務忙碌中，請再試一次。")
                        End Try

                        If dsReturn.Tables("CheckB2B").Rows.Count > 0 Then
                            If dsReturn.Tables("CheckB2B").Rows(0)("MSG") <> "" Then
                                Throw New Exception(dsReturn.Tables("CheckB2B").Rows(0)("MSG"))
                            End If
                        End If

                        '20231017 為了避免接續投保（IQPMAIN下pestatus= Y 跟C 的要鎖） 提示訊息：該保單保期與前張保單接續，請洽核保人員
                        Dim dtx As DataTable = New DataTable("IQPISSU")
                        dtx.Columns.AddRange(New DataColumn() {New DataColumn("IISSU", Type.GetType("System.String")), New DataColumn("DPLY_BGN", Type.GetType("System.DateTime")), New DataColumn("NRMK5", Type.GetType("System.String"))})
                        For Each rowissu As DataRow In StateMain.Tables("IQPISSU").Rows
                            If rowissu.RowState <> DataRowState.Deleted Then
                                dtx.Rows.Add(rowissu("IISSU"), UcRocDobjbgn.DateText, COUNCODE_ALL)
                            End If
                        Next
                        dtx.AcceptChanges()
                        Dim dsx As New DataSet
                        dsx.Tables.Add(dtx)
                        Dim dsResult As beNameValues = Me.InvokeBusinessService("B2BService.bsBHASContinueInsuranceCheck", dsx)
                        If dsResult.NameValues.Rows.Count > 0 Then
                            If Not String.IsNullOrEmpty(dsResult.Tables(0).Rows(0).Item(1)) Then Throw New Exception(dsResult.Tables(0).Rows(0).Item(1))
                        End If
                        '-------------------------------------------------------------------------------------------------------

                        '組資料 主要檢核
                        CheckMain()

                        '當CheckMain()檢核有異常，則IsInputErr = true，不做後續
                        If Me.IsInputErr Then
                            Exit Select
                        End If

                        If GetTargetProject() = "ITP" Then
                            If Me.txtCase.Text.Trim.Length > 0 AndAlso Me.txtCase.Text.Trim.Substring(0, 1) = "A" Then
                                '不鎖控航班
                            Else
                                '鎖控航班
                                'CheckFlight()  03.25系統移除”航班編號”欄位
                                CheckFlight_NEW()  '新航班編號”欄位
                            End If
                        End If

                        If InsertUP_BTN.Text = "X" Then
                            Throw New Exception(sbErrAll)
                        End If

                        'StateMain.AcceptChanges()
                        'BaseSetChildFormData(Me.dgdMain, Me.StateMain)

                        Dim dr2 As DataRow = BaseGetMainRow(StateMain)
                        AppCollectMainFormData(dr2)

                        '統一投保種類、轉換代碼，並寫入或更新（BHAS101003SaveTemp、BHAS101003SaveEdit）
                        procChangedTargetProject()
                        '快速理賠_有航班資料，則IFLIGHT設為Y
                        checkFlightOrNot()
                        StateMain.AcceptChanges()
                        BaseSetChildFormData(Me.dgdMain, Me.StateMain)

                        Me.BaseRestoreRowIndex = Me.BaseMainRowIndex
                        BaseSetFormData(StateMain)
                        'BaseMsgCopyOK()
                        UIFormIsDirty = True

                        If txtTOTPREM_NEW.Text < 70 Then
                            Throw New Exception("合計總保費需高於70元，請確認!")
                        End If

                        '試算後顯示保費欄位
                        ShowPremInfoOrNot(True)

                        '檢核要保號是否與被保人投保方案相符（因被保人方案任意切換，導致OID與投保方案不符的問題）
                        Dim dsCheckOIDnInsTyp As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b_bhas101003CheckInstyp", Me.txtIQPMainOid.Text.Trim)
                        If IsDBNull(dsCheckOIDnInsTyp) Or dsCheckOIDnInsTyp.Tables.Count = 0 Then
                            Throw New Exception("要保號與方案不符，請放棄此張要保號，重新輸入全新的要保資料!")
                        End If

                        CheckS.Text = "0"
                        CK1.Checked = True
                        'ReCaculate.Enabled = False

                        CK2.Checked = True
                        CK2.Enabled = False
                        hdValidateFlag.Value = "1"

                        '重新取值給與要保人關係，因為BaseSetFormData(StateMain)會將該欄清空
                        For Each drD1 As DataRow In StateMain.Tables("IQPISSU").Select("IISSU = '" & txtIins_ISSU2.Text & "'")
                            Relationship.SelectedValue = drD1("NINSREL")
                        Next

                        '將 方案資訊、保費小計、保費合計 整欄隱藏
                        divfuncinfoDisplay(False)

                        '是否使用EKYC
                        If CommonFunc.EKYCOrNot(Me, "BHAS101003", Me.txtCompanyid.Text) Then
                            '導傳至EKYC頁面
                            If IsInputErr = False AndAlso CheckSNNN.Text.Trim = "檢核資料正確" Then
                                CheckSNNN.Text = ""
                                DirectToEKYC()
                            End If
                        End If

                    Catch ex As Exception
                        'Me.ProcessException(ex)
                        ' Me.UIAppendError(ex.Message.ToString)

                        InsertUP_BTN.Text = "X"
                        CheckSNNN.Text = "資料有誤, 請修正後重新檢核"
                        Me.UIAppendError(ex.Message.ToString)
                        Throw New Exception(ex.Message.ToString)
                    End Try
                Case "PrintManager"
                    If Me.txtIQPMainOid.Text <> "" Then
                        If Me.txtCompanyid.Text = "KEN0" Then '和泰飽經
                            Dim IQPMainOID As String
                            IQPMainOID = txtIQPMainOid.Text
                            '取得SIGN_PAY_STATUS繳費沒
                            Dim SIGN_PAY_STATUS As String = ""
                            Dim dsSing As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"IQPMAIN", "OID  ='" + IQPMainOID + "'", "SIGN_PAY_STATUS"})
                            If Not IsNothing(dsSing) AndAlso dsSing.Tables.Count > 0 AndAlso dsSing.Tables(0).Rows.Count > 0 Then
                                SIGN_PAY_STATUS = dsSing.Tables(0).Rows(0)("SIGN_PAY_STATUS").ToString.Trim
                                If SIGN_PAY_STATUS = "Y" Then
                                Else
                                    Throw New Exception("尚未繳費(或已繳費但是資料尚未過檔至財務系統)，無法列要保文件.")
                                End If
                            End If
                        End If
                    End If
                    If Me.txtIQPMainOid.Text <> "" Then
                        PrintManager()
                    Else
                        Throw New Exception("無查詢編號，無法執行列印功能")
                    End If
                Case "PrintManager2"
                    Try
                        If Me.txtIQPMainOid.Text <> "" And txtIQUO.Text <> "" Then

                            '=======繳費狀態更新=======
                            '內建 證明書列印 申根簽證列印
                            '繳費後確認是否已繳費
                            'Dim ds As New DataSet
                            'Me.InvokeBusinessService("B2BService.bsBHASCheckSignPay", ds)
                            '=======繳費狀態更新=======

                            Dim IQPMainOID As String
                            IQPMainOID = txtIQPMainOid.Text
                            'SIGN_PAY_STATUS
                            '取得SIGN_PAY_STATUS繳費沒
                            Dim SIGN_PAY_STATUS As String = ""
                            Dim dsSing As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"IQPMAIN", "OID  ='" + IQPMainOID + "'", "SIGN_PAY_STATUS"})

                            If Not IsNothing(dsSing) AndAlso dsSing.Tables.Count > 0 AndAlso dsSing.Tables(0).Rows.Count > 0 Then
                                SIGN_PAY_STATUS = dsSing.Tables(0).Rows(0)("SIGN_PAY_STATUS").ToString.Trim
                                If SIGN_PAY_STATUS = "Y" Then
                                    '有繳費
                                    PrintManager2()
                                Else
                                    ' Throw New Exception("尚未繳費(或已繳費但是資料尚未過檔至財務系統)，無法列印證明書")
                                End If
                            End If
                        Else
                            Throw New Exception("無查詢編號，無法執行列印功能")
                        End If
                    Catch ex As Exception
                        Throw ex
                    End Try

                Case "PrintManager4"
                    Try
                        If Me.txtIQPMainOid.Text <> "" And txtIQUO.Text <> "" Then
                            Dim IQPMainOID As String
                            IQPMainOID = txtIQPMainOid.Text
                            'SIGN_PAY_STATUS
                            '取得SIGN_PAY_STATUS繳費沒
                            Dim SIGN_PAY_STATUS As String = ""
                            Dim dsSing As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"IQPMAIN", "OID  ='" + IQPMainOID + "'", "SIGN_PAY_STATUS"})

                            If Not IsNothing(dsSing) AndAlso dsSing.Tables.Count > 0 AndAlso dsSing.Tables(0).Rows.Count > 0 Then
                                SIGN_PAY_STATUS = dsSing.Tables(0).Rows(0)("SIGN_PAY_STATUS").ToString.Trim
                                If SIGN_PAY_STATUS = "Y" Then
                                    '有繳費
                                    PrintManager4()
                                Else
                                    'Throw New Exception("尚未繳費(或已繳費但是資料尚未過檔至財務系統)，無法列印證明書")
                                End If
                            End If
                        Else
                            Throw New Exception("無查詢編號，無法執行列印功能")
                        End If
                    Catch ex As Exception
                        Throw ex
                    End Try

                Case "PrintManager5"
                    '線上付款
                    Dim PAY_NUM As String = Me.txtPAY_NUM.Text
                    If PAY_NUM <> "" Then
                        If ConfigurationManager.AppSettings("bolWsITPSKL") Then
                            Dim sUrl As String = "https://www.sk858.com.tw/payrwd/PayMentQuery.aspx?pkey=QRCODE&txtorder_recode=" & PAY_NUM
                            Response.Write("<script>window.open('" + sUrl + "','');</script>")
                        Else
                            Dim sUrl As String = "https://ectest.sk-ins.com.tw/payrwd/PayMentQuery.aspx?pkey=QRCODE&txtorder_recode=" & PAY_NUM
                            Response.Write("<script>window.open('" + sUrl + "','');</script>")
                        End If
                    End If
                Case "ModifyWhenY"
                    '基本資料修改(調整要被保人)_僅PEStatus = "Y"有此功能
                    Try
                        Dim dsPEStatus As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.checkstatus", Me.txtIQPMainOid.Text.Trim)

                        If Not IsNothing(dsPEStatus) AndAlso dsPEStatus.Tables.Count > 0 AndAlso dsPEStatus.Tables(0).Rows.Count > 0 Then
                            Dim STR_PEStatus As String = dsPEStatus.Tables(0).Rows(0)("PESTATUS")
                            If STR_PEStatus.ToUpper.Trim = "Y" Then
                                PageSession("commandname") = "ModifyWhenY"
                                Me.CallOpenModel("BHAS_ModifyWhenY.aspx", "BHAS_ModifyWhenY", "", "ModifyWhenY", New Object() {txtIQPMainOid.Text})
                            Else
                                Throw New Exception("此要保號已經轉檔，無法再調整要、被保人資料。您可重新查詢此要保號，以確認要保狀態。")
                            End If
                        End If
                    Catch ex As Exception
                        Throw ex
                    End Try
                Case "MainClear"
                    BaseDoClear()
                    NOLockPage()
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("PayMoney", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveCheck", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("UpdateTemp", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument(UICommands.csDelete, "").Visible = False
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager", "").Visible = False
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager2", "").Visible = False
                    'MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager4", "").Visible = False
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("PrintManager5", "").Visible = False
                    'F115031101'MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp", "").Visible = True
                    MyBase.StateCommandBar.CommandState.FindByNameArgument("SaveTemp_New", "").Visible = True

                    '初始設定 不是查詢進來的
                    If ddlCountry_A1.SelectedValue = "" And txtIQPMainOid.Text = "" Then
                        Me.TRAVEL_hidden.Visible = False
                        Me.ShowS1.Visible = True
                    End If

                    '取得PASSPORT權限
                    Dim PassportAccess As String = ""
                    Dim dsSing As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_PASSPORT' AND CODE ='" + Me.ScsUserID + "'", "CODE"})

                    If Not IsNothing(dsSing) AndAlso dsSing.Tables.Count > 0 AndAlso dsSing.Tables(0).Rows.Count > 0 Then
                        PassportAccess = dsSing.Tables(0).Rows(0)("CODE").ToString.Trim
                    End If

                    '1140616 開啟供使用者勾選
                    PassportShow.Text = "1"
                    Me.CHK_P1.Visible = False
                    '初始設定 有權限才能看到
                    If PassportAccess <> "" Then
                        'Me.CHK_P1.Visible = True
                        'Me.CHK_P2.Visible = True
                        'PassportShow.Text = "1"
                    Else
                        'Me.CHK_P1.Visible = False
                        'Me.CHK_P2.Visible = False
                        'PassportShow.Text = "0"
                    End If

                    If txtTobjbgn.Text = "" Then
                        txtTobjbgn.Text = "12"
                    End If
                    If txtTobjend.Text = "" Then
                        txtTobjend.Text = "12"
                    End If
                    If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                        Benefi.Text = "法定繼承人"
                    End If

                    txtISEQ.Text = ""
                    txtIISSU_Ori.Text = ""
                    lblTargetPlan.Text = ""
                    txtCase.Text = ""

                    '將 方案資訊、保費小計、保費合計 整欄顯示
                    divfuncinfoDisplay(True)
                Case Else

                    '若點擊"查詢"按鈕，先檢查要保號是否存在
                    If e.CommandName = UICommands.csQuery Then
                        Dim dsOID As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"IQPMAIN", "OID  ='" + txtIQPMainOid.Text.Trim + "'", " OID "})

                        If Not IsNothing(dsOID) AndAlso dsOID.Tables.Count > 0 AndAlso dsOID.Tables(0).Rows.Count > 0 Then
                            MyBase.BaseDoUICommand(sender, e)

                            '20250512 user提IQP & ITP方案選擇不使用下拉，改透過方案選擇按鈕設定，故將可選國家最大化，待﹝保費試算﹞時檢核該方案能否選擇該國家。
                            Dim tr As HtmlTable = Form1.FindControl("divShow")
                            If (txtIQPMainOid.Text.Trim.ToUpper).Substring(0, 2) = "IT" Then
                                ddlTargetProject.SelectedValue = "ITP"
                                'tr.Attributes.Add("style", "")
                            ElseIf (txtIQPMainOid.Text.Trim.ToUpper).Substring(0, 2) = "IQ" Then
                                ddlTargetProject.SelectedValue = "IQP"
                                'tr.Attributes.Add("style", "display:none")
                            End If

                            btnQueryClick = True

                            '將 方案資訊、保費小計、保費合計 整欄顯示
                            divfuncinfoDisplay(True)

                        Else
                            Throw New Exception("無查詢要保號，請確認後再查詢")
                        End If
                    Else
                        MyBase.BaseDoUICommand(sender, e)
                    End If


            End Select
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Function checkTotal_it62it63(type As String) As Integer
        Dim TotAmt As Integer = 0
        Dim IQPType = "IQP11"

        If type = "IT63" Then
            IQPType = "IQP12"
        End If

        For Each dr_IT41 As DataRow In StateMain.Tables("IQPOBJ").Select(" instyp in ('" + type + "', '" + IQPType + "')")
            TotAmt += CInt(dr_IT41("INSAMT"))
        Next

        Return TotAmt
    End Function

    Private Sub dgdMain_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dgdMain.ItemDataBound
        '設定編輯和刪除的位置
        MyBase.BaseDataGridItemDataBound(CType(sender, DataGrid), e.Item, 0, 1)

        Select Case e.Item.ItemType

            Case ListItemType.EditItem
                '取得欄位值
                Dim dr As DataRow = CType(e.Item.DataItem, DataRowView).Row

                Dim ISEQ As Label = CType(e.Item.FindControl("labISEQ"), Label)
                Dim IISSU As Label = CType(e.Item.FindControl("labIISSU"), Label)
                Dim NISSU As Label = CType(e.Item.FindControl("labNISSU"), Label)
                Dim NINSREL As Label = CType(e.Item.FindControl("labNINSREL"), Label)
                Dim NISSU_TEL As Label = CType(e.Item.FindControl("labNISSU_TEL"), Label)
                Dim NAMEREL As Label = CType(e.Item.FindControl("labNAMEREL"), Label)

                txtISEQ.Text = ISEQ.Text
                txtIins_ISSU2.Text = IISSU.Text
                NAME_IISSU.Text = NISSU.Text
                IISSU_TEL.Text = NISSU_TEL.Text
                Benefi.Text = NAMEREL.Text
                'F114010802 旅遊險b2b被保人與要保人關係#9078
                If Relationship.Items.FindByValue(NINSREL.Text.Substring(0, 1)) Is Nothing Then
                    Relationship.SelectedValue = ""
                    UIAppendMessage("[與要保人關係]必需符合下拉清單內項目。")
                Else
                    Relationship.SelectedValue = NINSREL.Text.Substring(0, 1)
                End If
            Case ListItemType.AlternatingItem


        End Select
    End Sub
    Private Sub dgdMain_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dgdMain.ItemCommand
        Try

            Dim view As DataView = Me.AppGetDataGridBoundDataView(Me.dgdMain, Me.StateMain)
            If (e.Item.DataSetIndex) = -1 Then
                Exit Sub
            End If
            Dim rowView As DataRowView = view(e.Item.DataSetIndex)

            Dim row As DataRow = rowView.Row

            '保費狀態
            If e.CommandName = "Select" Then

                '先清空被保人相關資訊，避免上一動調整的被保人資料殘留
                IISSU_Clear()

                txtISEQ.Text = row("ISEQ")
                txtIISSU_Ori.Text = row("IISSU")
                CheckSNNN.Text = ""
                '主被保人 不能改ID 先拿掉
                '只能改方案

                ddl_IISSU.SelectedValue = row("IISSU_TP")
                txtIins_ISSU2.Text = row("IISSU")
                If PassportShow.Text = "1" Then
                    If ddl_IISSU.SelectedValue = "5" Then
                        CHK_P2.Checked = True
                    End If
                End If
                NAME_IISSU.Text = row("NISSU") 'NAME_IISSU
                'F114010802 旅遊險b2b被保人與要保人關係#9078
                Try
                    If Relationship.Items.FindByValue(row("NINSREL").ToString) Is Nothing Then
                        Relationship.SelectedValue = ""
                        txtRelationshipOther.Text = ""
                        UIAppendMessage("[與要保人關係]必需符合下拉清單內項目。 ")
                    Else
                        Relationship.SelectedValue = row("NINSREL")
                        If row("NINSREL") = "0" Then
                            If row("JOBOFFICE").ToString.Length > 5 Then
                                txtRelationshipOther.Text = row("JOBOFFICE").ToString.Substring(0, 5)
                                UIAppendMessage("[與要保人關係]選其他說明欄限五個字。 ")
                            Else
                                txtRelationshipOther.Text = row("JOBOFFICE")
                            End If
                        Else
                            txtRelationshipOther.Text = ""
                        End If
                    End If
                Catch ex As Exception
                    '不影響後面程序
                End Try
                IISSU_AREA.Text = row("IISSU_ZIP")
                IISSU_ADDR.Text = row("AISSU")
                IISSU_TEL.Text = row("NISSU_TEL")
                If IsDBNull(row("NISSUMAIL")) = False Then
                    EMAIL.Text = row("NISSUMAIL")
                End If

                UcRocDbirth.DateText = row("NISSUBD")
                Benefi.Text = row("NAMEREL")

                If IsDBNull(row("IDREL")) = False Then
                    Benefi_ID.Text = row("IDREL")
                    ShowS22.Visible = True
                    ShowS23.Visible = True
                End If
                If IsDBNull(row("IISSU_COUNTRY")) = False Then
                    ddlIISSU_COUNTRY.SelectedValue = row("IISSU_COUNTRY")
                End If
                If IsDBNull(row("TYPEREL")) = False Then
                    Relationship2.SelectedValue = row("TYPEREL")
                End If

                If IsDBNull(row("NAMEREL2")) = False Then
                    Benefi2.Text = row("NAMEREL2")
                End If
                If IsDBNull(row("IDREL2")) = False Then
                    Benefi_ID2.Text = row("IDREL2")
                End If
                If IsDBNull(row("TYPEREL2")) = False Then
                    Relationship22.SelectedValue = row("TYPEREL2")
                End If

                If IsDBNull(row("NAMEREL3")) = False Then
                    Benefi3.Text = row("NAMEREL3")
                End If
                If IsDBNull(row("IDREL3")) = False Then
                    Benefi_ID3.Text = row("IDREL3")
                End If
                If IsDBNull(row("TYPEREL3")) = False Then
                    Relationship23.SelectedValue = row("TYPEREL3")
                End If

                ddlSex.SelectedValue = row("IISSU_SEX")


                If IsDBNull(row("ZIPREL")) = False Then
                    txtZIPREL.Text = row("ZIPREL")
                End If
                If IsDBNull(row("TELREL")) = False Then
                    txtTELREL.Text = row("TELREL")
                End If
                If IsDBNull(row("AREL")) = False Then
                    txtAREL.Text = row("AREL")
                End If
                If IsDBNull(row("PASSPORT_NAME")) = False And IsDBNull(row("PASSPORT_NUMBER")) = False Then
                    PASSPORT_NAME.Text = HttpUtility.HtmlDecode(row("PASSPORT_NAME"))
                    PASSPORT_NUMBER.Text = HttpUtility.HtmlDecode(row("PASSPORT_NUMBER"))
                End If
                '是否懷孕
                If row("EMP") = "Y" Then
                    isEMP_Y.Checked = True
                ElseIf row("EMP") = "N" Then
                    isEMP_N.Checked = True
                End If

                ddlIT41.SelectedValue = ""
                ddlIT42.SelectedValue = ""
                ddlIT4Y.SelectedValue = ""
                ddlIT44.SelectedValue = ""
                ddlIT43.SelectedValue = ""
                ddlIT61.SelectedValue = ""
                ddlIT63.SelectedValue = ""
                ddlIT64.SelectedValue = ""
                ddlIT6E.SelectedValue = ""
                Dim IISSU_D As String = row("IISSU")
                For Each drD As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "' and instyp in ('IT41','IQP01')")
                    'drD.Delete()
                    If drD("insamt") = 0 Then
                        ddlIT41.SelectedValue = ""
                    Else
                        ddlIT41.SelectedValue = drD("insamt") / 10000
                    End If

                Next
                For Each drD As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "' and instyp in ('IT42', 'IQP02')")
                    'drD.Delete()
                    If drD("insamt") = 0 Then
                        ddlIT42.SelectedValue = ""
                    Else
                        ddlIT42.SelectedValue = drD("insamt") / 10000
                    End If
                Next
                For Each drD As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "' and instyp in ('IT44', 'IQP04')")
                    'drD.Delete()
                    If drD("insamt") = 0 Then
                        ddlIT44.SelectedValue = ""
                    Else
                        Dim selectedVal44 = (drD("insamt") / 10000).ToString()
                        ' 若 ddlIT44 沒有該值則新增，因為J、K、L方案重大燒燙傷保額可選高於250萬。
                        If ddlIT44.Items.FindByValue(selectedVal44) Is Nothing Then
                            ddlIT44.Items.Add(New ListItem(selectedVal44, selectedVal44))
                        End If

                        ddlIT44.SelectedValue = selectedVal44
                    End If
                Next
                For Each drD As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "' and instyp in ('IT43', 'IQP03')")
                    'drD.Delete()
                    If drD("insamt") = 0 Then
                        ddlIT43.SelectedValue = ""
                    Else
                        ddlIT43.SelectedValue = drD("insamt") / 10000
                    End If
                Next
                For Each drD As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "' and instyp in ('IT4Y', 'IQP0Y')")
                    'drD.Delete()
                    If drD("insamt") = 0 Then
                        ddlIT4Y.SelectedValue = ""
                    Else
                        ddlIT4Y.SelectedValue = drD("insamt") / 10000
                    End If
                Next

                For Each drD As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "' and instyp in ('IT62', 'IQP11')")
                    'drD.Delete()
                    If drD("insamt") = 0 Then
                        ddlIT61.SelectedValue = ""
                    Else
                        ddlIT61.SelectedValue = drD("insamt") / 10000
                        'rad_Disease_N.Checked = True
                        'ddlIT61.Enabled = True
                        'ddlIT63.Enabled = False
                    End If
                Next

                For Each drD As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "' and instyp in ('IT63', 'IQP12')")
                    'drD.Delete()
                    If drD("insamt") = 0 Then
                        ddlIT63.SelectedValue = ""
                    Else
                        ddlIT63.SelectedValue = drD("insamt") / 10000

                        'rad_Disease_Y.Checked = True
                        'ddlIT63.Enabled = True
                        'ddlIT61.Enabled = False
                    End If
                Next

                For Each drD As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "' and instyp in ('IT64', 'IQP13')")
                    'drD.Delete()
                    If drD("insamt") = 0 Then
                        ddlIT64.SelectedValue = ""
                    Else
                        ddlIT64.SelectedValue = drD("insamt") / 10000
                    End If
                Next

                For Each drD As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "' and instyp in ('IT6E')")
                    'drD.Delete()
                    If drD("insamt") = 0 Then
                        ddlIT6E.SelectedValue = ""
                    Else
                        ddlIT6E.SelectedValue = drD("insamt") / 10000
                    End If
                Next

                Dim Project As String = ""

#Region "1140415 因為當user已新增被保人，而又於ITP & IQP之間切換時會因方案差異造成錯誤，故取消方案檢查卡控，因最後依然要點選保費試算"
                '				If ddlTargetProject.SelectedValue = "ITP" Then
                '					'iPlanNo 方案
                '					Project = row("iPlanNo").trim
                '					Project = Project.Substring(0, 1)

                '#Region "1140415 嘗試調改StateMain內容，但後決定不做方案檢查卡控，因最後依然要點選保費試算"
                '					If Not String.IsNullOrEmpty(row("iPlanNo")) Then
                '						Project = row("iPlanNo").trim
                '						Project = Project.Substring(0, 1)
                '					ElseIf Not String.IsNullOrEmpty(txtCase.Text) Then
                '						'當user用IQP新增被保人，改成ITP再點選已新增的被保人時，改變SateMain中的值滿足後續需求
                '						Project = (txtCase.Text).Substring(0, 1)
                '						For Each drD1 As DataRow In StateMain.Tables("IQPISSU").Select("IISSU = '" & IISSU_D & "'")
                '							drD1("iPlanNo") = Project
                '						Next
                '#End Region
                '					Else
                '						Throw New Exception("請先選擇[投保方案]")
                '					End If
                '				End If
#End Region

                If ddlTargetProject.SelectedValue = "ITP" Then
                    If (row("iPlanNo").ToString.Trim).Length > 0 Then
                        Project = (row("iPlanNo").ToString.Trim).Substring(0, 1)
                    Else
                        UIAppendMessage("請留意當前要保號與要保方案是否相符。")
                    End If
                End If

                Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
                'txtJOB
                If IsDBNull(row("JOBCONTENT")) = False Then
                    txtJOB.Text = row("JOBCONTENT")
                Else
                    txtJOB.Text = ""
                End If

                Dim IQPMAINOID As String = ""
                Dim instyp As String = ""
                If Project = "A" Then
                    BT1_BUTTON(age)
                ElseIf Project = "B" Then
                    BT2_BUTTON(age)
                ElseIf Project = "C" Then
                    BT3_BUTTON(age)
                    'ElseIf Project = "D" Then
                    '    BT4_BUTTON(age, instyp)
                ElseIf Project = "D" Then
                    BT5_BUTTON(age)
                ElseIf Project = "J" Then
                    BT6_BUTTON(age)
                ElseIf Project = "K" Then
                    BT7_BUTTON(age)
                ElseIf Project = "L" Then
                    BT8_BUTTON(age)
                ElseIf Project = "M" Then
                    BT9_BUTTON(age)
                ElseIf Project = "N" Then
                    BT10_BUTTON(age)
                ElseIf Project = "O" Then
                    BT11_BUTTON(age)
                End If

                '將 方案資訊、保費小計、保費合計 整欄顯示
                divfuncinfoDisplay(True)

            ElseIf e.CommandName = "Delete" Then
                CheckSNNN.Text = ""

#Region "F115031101 不論是否主被保人，直接刪除。所有ISEQ往前遞補重新給號"
                '至少留一個被保人
                If StateMain.Tables("IQPISSU").Rows.Count = 1 Then
                    Throw New Exception("至少要有一位[被保險人]。")
                    Exit Sub
                End If

                '刪除dgdMain1  
                Dim IISSU_D As String = row("IISSU")
                Dim ISWQ_D As String = row("ISEQ")
                For Each drD As DataRow In StateMain.Tables("IQPISSU").Select("IISSU = '" & IISSU_D & "'")
                    drD.Delete()
                Next


                '刪除dgdMain2  
                For Each drD2 As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "'")
                    drD2.Delete()
                Next

                '立即針對Table做刪除，避免後續查找StateMain時發生錯誤
                StateMain.Tables("IQPISSU").AcceptChanges()
                StateMain.Tables("IQPOBJ").AcceptChanges()
                'F115031101 ISEQ重新給號
                ResetISEQAfterDEL(ISWQ_D, "IQPISSU")

                BaseSetChildFormData(dgdMain, StateMain)
                BaseSetChildFormData(dgdMain2, StateMain)

                txtISEQ.Text = ""
                txtIISSU_Ori.Text = ""
#End Region

                'If row("ISEQ") = "1" Then
                '    Throw New Exception("請勿刪除[主被保險人]")
                '    Exit Sub
                'Else
                '    '刪除dgdMain1  
                '    Dim IISSU_D As String = row("IISSU")
                '    For Each drD As DataRow In StateMain.Tables("IQPISSU").Select("IISSU = '" & IISSU_D & "'")
                '        drD.Delete()
                '    Next

                '    ''F115031101 預習ISEQ重新給號
                '    'ResetISEQAfterDEL(row("ISEQ"))

                '    BaseSetChildFormData(dgdMain, StateMain)
                '    '刪除dgdMain2  
                '    For Each drD2 As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_D & "'")
                '        drD2.Delete()
                '    Next
                '    BaseSetChildFormData(dgdMain2, StateMain)

                '    '立即針對Table做刪除，避免後續查找StateMain時發生錯誤
                '    StateMain.Tables("IQPISSU").AcceptChanges()
                '    StateMain.Tables("IQPOBJ").AcceptChanges()

                '    txtISEQ.Text = ""
                '    txtIISSU_Ori.Text = ""
                'End If
            End If
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Private Sub SetDropDownList(ByVal Obj As String, ByVal ddlCur As DropDownList)
        Dim item As ListItem
        Dim Flag As Boolean = False
        For Each item In ddlCur.Items
            If item.Value = Obj.ToString Then
                item.Selected = True
                Flag = True
                Return
            End If
        Next
        If Flag = False Then
            ddlCur.Items.Add(New ListItem(Obj.ToString()))
            ddlCur.SelectedValue = Obj.ToString()
        End If
    End Sub
    ''列印1
    Private Sub PrintManager()

        If txtIQPMainOid.Text.Trim = "" Or Me.txtRtype.Text = "V" Then
            Throw New Exception("檢視模式無法列印!!!")
        End If

        '新版直接印
        Dim ds As beFSReports
        Dim dr As beFSReports.ReportItemsRow
        Dim dsParam As New beNameValues
        dsParam.NameValues.AddNameValuesRow("OID", txtIQPMainOid.Text.Trim)

        If (txtIQPMainOid.Text.Trim).Substring(0, 2) = "IT" Then
            dsParam.NameValues.AddNameValuesRow("B2BReport", "/B2BReport/BHAS101001") 'NEW
            dsParam.NameValues.AddNameValuesRow("R", "N")
        ElseIf (txtIQPMainOid.Text.Trim).Substring(0, 2) = "IQ" Then
            dsParam.NameValues.AddNameValuesRow("B2BReport", "/B2BReport/BHAS102001")
            dsParam.NameValues.AddNameValuesRow("R", "N")
        End If

        ds = CType(InvokeBusinessService("B2BService.bsBHAS201000PrintM", dsParam), beFSReports)
        dr = ds.ReportItems(0)

        If Request.Browser.Browser = "IE" Then 'ie瀏灠器
            '跳出視窗開啟檔案 
            FSReportViewer.ViewReport(Me, ds, FSReportViewer.ViewType.TempUrl, "status=no,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes")
        Else
            'hidUrl.Value = "http://" & System.Environment.MachineName() & createHttpTempFile(dr)
            Dim vu As VirtualUser = VirtualUser.CreateVirtualUserFromConfig(VirtualUser.ConfigKeyFileServiceUser)
            Try
                vu.Impersonate()
                Me.JSAppendFormTailScripts("window.open('" & FSCreateTempContent(HttpUtility.UrlEncode(System.IO.Path.GetFileName(dr.FileName)), FSReadContent(Server.MapPath(createHttpTempFile(dr)))) & "');")
            Catch ex As Exception
                Throw
            Finally
                vu.Dispose()
            End Try
        End If

    End Sub
    ''列印2 證明書列印 要繳費後才能印
    Private Sub PrintManager2()

        If txtIQPMainOid.Text.Trim = "" Or Me.txtRtype.Text = "V" Then
            Throw New Exception("檢視模式無法列印!!!")
        End If
        Dim ds As beFSReports
        Dim dr As beFSReports.ReportItemsRow
        Dim dsParam As New beNameValues
        dsParam.NameValues.AddNameValuesRow("OID", txtIQPMainOid.Text.Trim)
        dsParam.NameValues.AddNameValuesRow("B2BReport", "/B2BReport/ITP101000_1")
        'bsBBTP002000PrintM
        dsParam.NameValues.AddNameValuesRow("R", "N")
        ds = CType(InvokeBusinessService("B2BService.bsBHAS201000PrintM", dsParam), beFSReports)
        dr = ds.ReportItems(0)

        'ds = CType(InvokeBusinessService("B2BService.bsITP101000Print", dsParam), beFSReports)
        'dr = ds.ReportItems(0)

        If Request.Browser.Browser = "IE" Then 'ie瀏灠器
            '跳出視窗開啟檔案 
            'FSReportViewer.ViewReport(Me, ds, FSReportViewer.ViewType.TempUrl, "status=no,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes")
            'hidUrl.Value = "http://" & System.Environment.MachineName() & createHttpTempFile(dr)
            FSReportViewer.ViewReport(Me, ds, FSReportViewer.ViewType.TempUrl, "status=no,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes")
        Else
            Dim vu As VirtualUser = VirtualUser.CreateVirtualUserFromConfig(VirtualUser.ConfigKeyFileServiceUser)
            vu.Impersonate()
            Me.JSAppendFormTailScripts("window.open('" & FSCreateTempContent(HttpUtility.UrlEncode(System.IO.Path.GetFileName(dr.FileName)), FSReadContent(Server.MapPath(createHttpTempFile(dr)))) & "');")
            vu.Dispose()

        End If

        '列印2 證明書  要UPDATE
        Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.updateprint", Me.txtIQPMainOid.Text.Trim)

    End Sub
    ''列印3
    Private Sub PrintManager3()

        If txtIQPMainOid.Text.Trim = "" Or Me.txtRtype.Text = "V" Then
            Throw New Exception("檢視模式無法列印!!!")
        End If

        Dim ireport As String = Guid.NewGuid().ToString("N").ToUpper
        '判斷是否有未滿15
        Dim dsB2B As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.itp101000", Me.txtIQPMainOid.Text.Trim)

        If dsB2B.Tables("IQPOBJ").Rows.Count > 0 Then
            Dim ds As beFSReports
            Dim dr As beFSReports.ReportItemsRow
            Dim dsParam As New beNameValues
            dsParam.NameValues.AddNameValuesRow("OID", txtIQPMainOid.Text.Trim)
            dsParam.NameValues.AddNameValuesRow("B2BReport", "/B2BReport/ITP101000_31")

            'bsBBTP002000PrintM
            dsParam.NameValues.AddNameValuesRow("R", "N")
            ds = CType(InvokeBusinessService("B2BService.bsBHAS201000PrintM", dsParam), beFSReports)
            dr = ds.ReportItems(0)

            If Request.Browser.Browser = "IE" Then 'ie瀏灠器
                '跳出視窗開啟檔案 
                FSReportViewer.ViewReport(Me, ds, FSReportViewer.ViewType.TempUrl, "status=no,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes")
            Else
                Dim vu As VirtualUser = VirtualUser.CreateVirtualUserFromConfig(VirtualUser.ConfigKeyFileServiceUser)
                vu.Impersonate()
                Me.JSAppendFormTailScripts("window.open('" & FSCreateTempContent(HttpUtility.UrlEncode(System.IO.Path.GetFileName(dr.FileName)), FSReadContent(Server.MapPath(createHttpTempFile(dr)))) & "');")
                vu.Dispose()
            End If
        Else
            Dim ds As beFSReports
            Dim dr As beFSReports.ReportItemsRow
            Dim dsParam As New beNameValues
            dsParam.NameValues.AddNameValuesRow("OID", txtIQPMainOid.Text.Trim)
            dsParam.NameValues.AddNameValuesRow("B2BReport", "/B2BReport/ITP101000_3")

            'bsBBTP002000PrintM
            dsParam.NameValues.AddNameValuesRow("R", "N")
            ds = CType(InvokeBusinessService("B2BService.bsBHAS201000PrintM", dsParam), beFSReports)
            dr = ds.ReportItems(0)

            If Request.Browser.Browser = "IE" Then 'ie瀏灠器
                '跳出視窗開啟檔案 
                FSReportViewer.ViewReport(Me, ds, FSReportViewer.ViewType.TempUrl, "status=no,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes")
            Else
                Dim vu As VirtualUser = VirtualUser.CreateVirtualUserFromConfig(VirtualUser.ConfigKeyFileServiceUser)
                vu.Impersonate()
                Me.JSAppendFormTailScripts("window.open('" & FSCreateTempContent(HttpUtility.UrlEncode(System.IO.Path.GetFileName(dr.FileName)), FSReadContent(Server.MapPath(createHttpTempFile(dr)))) & "');")
                vu.Dispose()

            End If

        End If

    End Sub
    ''列印4 申根
    Private Sub PrintManager4()

        If txtIQPMainOid.Text.Trim = "" Or Me.txtRtype.Text = "V" Then
            Throw New Exception("檢視模式無法列印!!!")
        End If

        If Checkbox2.Checked = True And Me.TRAVEL_A.SelectedValue = "750" Then
            Check750.Text = TRAVEL_A.SelectedValue.Trim
            Dim ds As beFSReports
            Dim dr As beFSReports.ReportItemsRow
            Dim dsParam As New beNameValues
            dsParam.NameValues.AddNameValuesRow("OID", txtIQPMainOid.Text.Trim)
            dsParam.NameValues.AddNameValuesRow("B2BReport", "/B2BReport/BHAS_Schengen01")

            'bsBBTP002000PrintM
            dsParam.NameValues.AddNameValuesRow("R", "N")
            ds = CType(InvokeBusinessService("B2BService.bsBHAS201000PrintM", dsParam), beFSReports)
            dr = ds.ReportItems(0)

            If Request.Browser.Browser = "IE" Then 'ie瀏灠器
                '跳出視窗開啟檔案 
                FSReportViewer.ViewReport(Me, ds, FSReportViewer.ViewType.TempUrl, "status=no,toolbar=no,menubar=no,location=no,resizable=yes,scrollbars=yes")
            Else

                Dim vu As VirtualUser = VirtualUser.CreateVirtualUserFromConfig(VirtualUser.ConfigKeyFileServiceUser)
                vu.Impersonate()
                Me.JSAppendFormTailScripts("window.open('" & FSCreateTempContent(HttpUtility.UrlEncode(System.IO.Path.GetFileName(dr.FileName)), FSReadContent(Server.MapPath(createHttpTempFile(dr)))) & "');")
                vu.Dispose()

            End If

            '列印2 證明書  要UPDATE
            Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.updateprint", Me.txtIQPMainOid.Text.Trim)
        Else
            Throw New Exception("非申根證明文件無法列印!!!")
        End If


    End Sub
    '給按鈕權限
    Protected Overrides Function ScsGetAssociatedAction(ByVal commandName As String) As Integer
        Select Case commandName
            Case "SaveTemp"
                Return ScsActionType.Run
            Case "SaveEdit"
                Return ScsActionType.Run
            Case "SaveCheck"
                Return ScsActionType.Run
            Case "PayMoney"
                Return ScsActionType.Run
            Case "PrintManager"
                Return ScsActionType.Run
            Case "PrintManager2"
                Return ScsActionType.Run
            Case "PrintManager4"
                Return ScsActionType.Run
            Case "PrintManager5"
                Return ScsActionType.Run
            Case "MainClear"
                Return ScsActionType.Run
            Case "Print_Combine"
                Return ScsActionType.Run
            Case "SCOPY"
                Return ScsActionType.Run
            Case "UpdateTemp"
                Return ScsActionType.Run
            Case "SaveTemp_New"
                Return ScsActionType.Run
            Case "ModifyWhenY"
                Return ScsActionType.Run
            Case Else
                Return MyBase.ScsGetAssociatedAction(commandName)
        End Select

    End Function
#Region "投保方案Button"
    Private Sub BT1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT1.Click
        '國內陸上型
        Try
            lblTargetPlan.Text = "A 國內海陸型"
            ddlTargetProject.SelectedValue = "ITP"
            Dim ChkIns As String
            ChkIns = CheckBasic()
            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            '旅遊地區整合
            Dim COUNCODE_ALL As String = ""
            COUNCODE_ALL = CHECK_COUNCODE1()
            TRAVEL_A.SelectedValue = COUNCODE_ALL

            If TRAVEL_A.SelectedValue <> "400" Then
                Throw New Exception("旅遊地區為國內，才能選擇此方案")
            End If

            GetDay() '計算天數

            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            Dim dobj As DiseaseFields = New DiseaseFields()
            dobj.IT61 = False
            dobj.IT63 = False
            dobj.IT64 = False
            dobj.IT6E = False
            dobj.IT65 = False
            dobj.IT66 = False
            dobj.IT6F = False
            dobj.IT6G = False
            dobj.IT6D = False
            FieldsSwtich_Disease(dobj)

            ddlIT43.Enabled = False
            ddlIT43.SelectedValue = ""
            txtIT43.Text = ""

            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String

            Me.txtCase.Text = "A 國內海陸型"
            IISSU_N.Enabled = True

            '將各方案金額清空
            ResetFields(include65_66:=True, include6D:=True, include6F_6G:=True)

            ddlIT61.SelectedValue = ""
            'ITP 不便險取保費的方式
            '1:InsType(險種)
            '2:InsDay(天數)
            '3.InsCoverage 保額 元
            '4.InsMdeduct 自負額 元
            '5.CountryType 1=國內 2=國外

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "50000"
            InsMdeduct = "2500"

            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '==========IT28==========
            InsType = "IT28"
            InsCoverage = "3000"
            InsMdeduct = "0"
            Dim ds28 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds28.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds28.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If

            '==========IT2D==========
            'InsType = "IT2D"
            'InsCoverage = "3000"
            'InsMdeduct = "0"
            'Dim ds2D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds2D.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT2D.Text = ds2D.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT2D.Text = "0"
            'End If
            '==========IT2G==========
            InsType = "IT2G"
            InsCoverage = "3000"
            InsMdeduct = "0"
            Dim ds2G As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2G.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2G.Text = ds2G.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2G.Text = "0"
            End If

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)
            'LockPage()

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try

    End Sub
    Private Sub BT2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT2.Click

        '國內翱翔型
        '判斷被保人年齡
        Try
            ddlTargetProject.SelectedValue = "ITP"
            'CheckFlight() '航班鎖控 03.25系統移除”航班編號”欄位
            CheckFlight_NEW()  '新航班鎖控
            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If

            '旅遊地區整合
            Dim COUNCODE_ALL As String = ""
            COUNCODE_ALL = CHECK_COUNCODE1()
            TRAVEL_A.SelectedValue = COUNCODE_ALL

            If TRAVEL_A.SelectedValue <> "400" Then
                Throw New Exception("旅遊地區為國內，才能選擇此方案")
            End If

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數

            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If
            ddlIT61.Enabled = False
            ddlIT61.SelectedValue = ""
            txtIT61.Text = ""
            ddlIT63.Enabled = False
            ddlIT63.SelectedValue = ""
            txtIT63.Text = ""
            ddlIT64.SelectedValue = ""
            lblIT65.Text = ""
            lblIT66.Text = ""
            txtIT64.Text = ""
            txtIT65.Text = ""
            txtIT66.Text = ""
            ddlIT43.Enabled = False
            ddlIT43.SelectedValue = ""
            txtIT43.Text = ""

            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "B 國內飛翔型"
            IISSU_N.Enabled = True

            txtIT11.Text = "0"
            txtIT27.Text = "0"
            txtIT28.Text = "0"
            txtIT2D.Text = "0"
            txtIT2E.Text = "0"
            txtIT2G.Text = "0"
            txtIT2I.Text = "0"
            txtIT2M.Text = "0"
            txtIT2N.Text = "0"
            txtIT2P.Text = "0"
            txtIT2R.Text = "0"
            txtIT2S.Text = "0"
            txtIT2T.Text = "0"
            txtIT2U.Text = "0"
            txtIT31.Text = "0"
            txtIT32.Text = "0"
            txtIT33.Text = "0"
            txtIT34.Text = "0"
            txtIT35.Text = "0"
            txtIT36.Text = "0"
            txtIT37.Text = "0"
            txtIT38.Text = "0"
            txtIT39.Text = "0"
            txtIT3A.Text = "0"
            txtIT3B.Text = "0"
            txtIT3C.Text = "0"
            txtIT3D.Text = "0"
            txtIT3E.Text = "0"

            ddlIT61.SelectedValue = ""

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "50000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "1500"
            InsMdeduct = "0"

            Dim ds24 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds24.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds24.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "1000"
            InsMdeduct = "0"

            Dim ds2A As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2A.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = ds2A.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If

            '====================================
            InsType = "IT27"
            InsCoverage = "2000"
            InsMdeduct = "0"

            Dim ds27 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds27.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT27.Text = ds27.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT27.Text = "0"
            End If
            '====================================
            InsType = "IT28"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds28 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds28.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds28.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If
            '====================================
            'InsType = "IT2D"
            'InsCoverage = "3000"
            'InsMdeduct = "0"

            'Dim ds2D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds2D.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT2D.Text = ds2D.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT2D.Text = "0"
            'End If
            '====================================
            InsType = "IT2G"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2G As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2G.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2G.Text = ds2G.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2G.Text = "0"
            End If

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)
            'LockPage()
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Private Sub CheckFlight()
        Try
            If String.IsNullOrEmpty(txtFlightNum.Text) Then
                Throw New Exception("選取旅遊方案時，需填寫 欄位[航班編號]")
            End If
            For Each item As Char In txtFlightNum.Text
                If IsNumeric(item) Or (item >= "a" And item <= "z") Or (item >= "A" And item <= "Z") Or item = "," Then
                Else
                    Throw New Exception("限制為[英文]或[數字]或[,]")
                End If
            Next
            Dim dsAccount As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_IFLIGHT'", "DATA1"})

            If dsAccount.Tables(0).Rows.Count > 0 Then
                Dim arrStr() = dsAccount.Tables(0).Rows(0).Item(0).ToString.Split(",")
                For Each s As String In arrStr
                    If txtFlightNum.Text.ToUpper().Contains(s) Then
                        Throw New Exception("該航班不受理投保旅綜險，若有保障需求可改投保旅平險，造成不便敬請見諒。")
                    End If
                Next
            End If
        Catch ex As Exception
            CheckSNNN.Text = ""
            ProcessException(ex)
        End Try
    End Sub
    Private Sub CheckFlight_NEW()
        Try
            If rad_flight_Y.Checked = True Then

                Dim dsAccount As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_IFLIGHT'  AND STATUS='1'", "DATA1,DATA2,DATA3,NREMARK"})

                If dsAccount.Tables(0).Rows.Count > 0 Then
                    Dim DATA2, DATA3 As Date
                    DATA2 = CDate(dsAccount.Tables(0).Rows(0).Item("DATA2").ToString.Trim)
                    DATA3 = CDate(dsAccount.Tables(0).Rows(0).Item("DATA3").ToString.Trim)
                    If (UcRocDobjbgn.GetDate >= DATA2) And (UcRocDobjbgn.GetDate <= DATA3) Then
                        Throw New Exception(dsAccount.Tables(0).Rows(0).Item("NREMARK").ToString.Trim)
                    End If
                    If (UcRocDobjend.GetDate >= DATA2) And (UcRocDobjend.GetDate <= DATA3) Then
                        Throw New Exception(dsAccount.Tables(0).Rows(0).Item("NREMARK").ToString.Trim)
                    End If
                End If
            End If
        Catch ex As Exception
            CheckSNNN.Text = ""
            ProcessException(ex)
        End Try
    End Sub
    Private Sub BT3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT3.Click
        '海外基本型
        '判斷被保人年齡

        Try
            '於後續BT3_BUTTON(age)內一併調整
            'lblTargetPlan.Text = "C 海外基本型"
            ddlTargetProject.SelectedValue = "ITP"

            'CheckFlight() '航班鎖控
            CheckFlight_NEW() '新航班鎖控
            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If
            '旅遊地區整合
            Dim COUNCODE_ALL As String = ""
            COUNCODE_ALL = CHECK_COUNCODE1()
            TRAVEL_A.SelectedValue = COUNCODE_ALL

            If TRAVEL_A.SelectedValue = "400" Then
                Throw New Exception("旅遊地區為國內，無法能選擇此方案")
            End If

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數
            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            '於後續BT3_BUTTON(age)內一併調整
            'Me.txtCase.Text = "C 海外輕旅型"

            IISSU_N.Enabled = True

            '將各方案金額清空
            ResetFields(include65_66:=True, include6D:=False, include6F_6G:=False)

            'ddlIT61.Enabled = True

            BT3_BUTTON(age)

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    Private Sub BT4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT4.Click
        '海外自駕型
        Try

            ddlTargetProject.SelectedValue = "ITP"

            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If
            '旅遊國家 組合判斷2
            If Checkbox2.Checked = True Then
                Throw New Exception("勾選[申根公約國證明文件]請選擇[申根輕旅型]。")
            End If

            If ddlCountry_B2.SelectedValue = "" And ddlCountry_C2.SelectedValue = "" And ddlCountry_D2.SelectedValue = "" And
                ddlCountry_E2.SelectedValue = "" And ddlCountry_F2.SelectedValue = "" And ddlCountry_G2.SelectedValue = "" And ddlCountry_H2.SelectedValue = "" Then
                '只有一組 用第一組判斷就好
                COUNCODE1 = HttpUtility.HtmlEncode(ddlCountry_A2.SelectedValue.Substring(2, 3))
                If COUNCODE1 = "750" Then
                    '申根地區判斷
                    If Checkbox2.Checked = True Then
                        If PASSPORT_NAME.Text = "" Or PASSPORT_NUMBER.Text = "" Then
                            Throw New Exception("勾選申根公約國證明文件 請輸入英文姓名 與 護照號碼。")
                        Else
                            '申根地區
                            COUNCODE1 = "750"
                        End If
                    Else
                        '改成歐洲地區
                        COUNCODE1 = "740"
                    End If

                End If
            Else
                Dim ddlA, ddlB, ddlC, ddlD, ddlE, ddlF, ddlG, ddlH As String

                If ddlCountry_A2.SelectedValue <> "" Then
                    ddlA = HttpUtility.HtmlEncode(ddlCountry_A2.SelectedValue.Substring(2, 3))
                Else
                    ddlA = ""
                End If
                If ddlCountry_B2.SelectedValue <> "" Then
                    ddlB = HttpUtility.HtmlEncode(ddlCountry_B2.SelectedValue.Substring(2, 3))
                Else
                    ddlB = ""
                End If
                If ddlCountry_C2.SelectedValue <> "" Then
                    ddlC = HttpUtility.HtmlEncode(ddlCountry_C2.SelectedValue.Substring(2, 3))
                Else
                    ddlC = ""
                End If
                If ddlCountry_D2.SelectedValue <> "" Then
                    ddlD = HttpUtility.HtmlEncode(ddlCountry_D2.SelectedValue.Substring(2, 3))
                Else
                    ddlD = ""
                End If
                If ddlCountry_E2.SelectedValue <> "" Then
                    ddlE = HttpUtility.HtmlEncode(ddlCountry_E2.SelectedValue.Substring(2, 3))
                Else
                    ddlE = ""
                End If
                If ddlCountry_F2.SelectedValue <> "" Then
                    ddlF = HttpUtility.HtmlEncode(ddlCountry_F2.SelectedValue.Substring(2, 3))
                Else
                    ddlF = ""
                End If
                If ddlCountry_G2.SelectedValue <> "" Then
                    ddlG = HttpUtility.HtmlEncode(ddlCountry_G2.SelectedValue.Substring(2, 3))
                Else
                    ddlG = ""
                End If
                If ddlCountry_H2.SelectedValue <> "" Then
                    ddlH = HttpUtility.HtmlEncode(ddlCountry_H2.SelectedValue.Substring(2, 3))
                Else
                    ddlH = ""
                End If

                If ddlA = "400" Or ddlB = "400" Or ddlC = "400" Or ddlD = "400" Or ddlE = "400" Or ddlF = "400" Or ddlG = "400" Or ddlH = "400" Then
                    ddlCountry_A1.SelectedValue = "1"
                    ddlCountry_A2.SelectedValue = "TW400"
                    ddlCountry_B1.SelectedValue = ""
                    ddlCountry_B2.SelectedValue = ""
                    ddlCountry_C1.SelectedValue = ""
                    ddlCountry_C2.SelectedValue = ""
                    ddlCountry_D1.SelectedValue = ""
                    ddlCountry_D2.SelectedValue = ""
                    ddlCountry_E1.SelectedValue = ""
                    ddlCountry_E2.SelectedValue = ""
                    ddlCountry_F1.SelectedValue = ""
                    ddlCountry_F2.SelectedValue = ""
                    ddlCountry_G1.SelectedValue = ""
                    ddlCountry_G2.SelectedValue = ""
                    ddlCountry_H1.SelectedValue = ""
                    ddlCountry_H2.SelectedValue = ""
                    Throw New Exception("選擇台灣地區時無法再選取其他區域。")
                End If

                '全球各地
                If ddlA = "750" Or ddlB = "750" Or ddlC = "750" Or ddlD = "750" Or ddlE = "750" Or ddlF = "750" Or ddlG = "750" Or ddlH = "750" Then
                    '申根地區判斷
                    If Checkbox2.Checked = True Then
                        If PASSPORT_NAME.Text = "" Or PASSPORT_NUMBER.Text = "" Then
                            Throw New Exception("勾選申根公約國證明文件 請輸入英文姓名 與 護照號碼。")
                        Else
                            '申根地區
                            COUNCODE1 = "750"
                        End If
                    Else
                        '改成歐洲地區
                        COUNCODE1 = "740"
                    End If
                Else
                    '改成全球地區
                    COUNCODE1 = "800"
                End If

            End If
            TRAVEL_A.SelectedValue = COUNCODE1

            If TRAVEL_A.SelectedValue = "400" Then
                Throw New Exception("旅遊地區為國內， 無法能選擇此方案")
            End If

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            'If txtPESTATUS.Text.Trim = "S" Then
            '    ChkTime = CheckTimeK417()
            '    If ChkTime <> String.Empty Then
            '        Throw New Exception(ChkTime)
            '    End If
            'Else
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數

            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            'ddlIT61.Enabled = True
            'If rad_Disease_N.Checked = True Then
            '	ddlIT61.Enabled = True
            '	ddlIT63.Enabled = False
            'ElseIf rad_Disease_Y.Checked = True Then
            '	ddlIT61.Enabled = False
            '	ddlIT63.Enabled = True
            'End If
            ddlIT61.Enabled = True
            ddlIT63.Enabled = True

            '預設被保人投保上限
            'GetDefault(age)
            'ddlIT61.SelectedValue = HttpUtility.HtmlEncode(ddlIT42.SelectedValue)

            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String

            Me.txtCase.Text = "D.海外自駕型"

            txtIT11.Text = "0"
            txtIT3A.Text = "0"
            txtIT38.Text = "0"
            txtIT2I.Text = "0"
            txtIT27.Text = "0"
            txtIT28.Text = "0"
            txtIT2E.Text = "0"
            txtIT2G.Text = "0"

            txtIT11.Text = "0"
            txtIT3B.Text = "0"
            txtIT39.Text = "0"
            txtIT3A.Text = "0"
            txtIT37.Text = "0"
            txtIT36.Text = "0"
            txtIT38.Text = "0"
            txtIT2M.Text = "0"
            txtIT27.Text = "0"
            txtIT2N.Text = "0"

            IISSU_N.Enabled = True

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "500000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '====================================
            InsType = "IT3B"
            InsCoverage = "1000"
            InsMdeduct = "0"

            Dim ds2L As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2L.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3B.Text = ds2L.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3B.Text = "0"
            End If

            '====================================
            InsType = "IT39"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds29 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds29.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT39.Text = ds29.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT39.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2A As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2A.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = ds2A.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If
            '====================================
            InsType = "IT37"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim ds2K As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2K.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT37.Text = ds2K.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT37.Text = "0"
            End If
            '====================================
            InsType = "IT36"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim ds2J As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2J.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT36.Text = ds2J.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT36.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds24 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds24.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds24.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If
            '====================================
            InsType = "IT2M"
            InsCoverage = "2000"
            InsMdeduct = "0"

            Dim ds5 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds5.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2M.Text = ds5.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2M.Text = "0"
            End If
            '====================================
            InsType = "IT27"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds4 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds4.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT27.Text = ds4.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT27.Text = "0"
            End If
            '====================================
            InsType = "IT2N"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim dsN As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsN.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2N.Text = dsN.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2N.Text = "0"
            End If

            '====================================

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    Private Sub BT5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT5.Click
        Try
            '於後續BT5_BUTTON(age)內一併調整
            'lblTargetPlan.Text = "D 海外超值型"

            ddlTargetProject.SelectedValue = "ITP"

            'CheckFlight() '航班鎖控
            CheckFlight_NEW() '新航班鎖控
            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If

            '旅遊地區整合
            Dim COUNCODE_ALL As String = ""
            COUNCODE_ALL = CHECK_COUNCODE1()
            TRAVEL_A.SelectedValue = COUNCODE_ALL

            If TRAVEL_A.SelectedValue = "400" Then
                Throw New Exception("旅遊地區為國內，無法能選擇此方案")
            End If

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數

            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            '於後續BT5_BUTTON(age)內一併調整
            'Me.txtCase.Text = "D 海外超值型"

            IISSU_N.Enabled = True

            '將各方案金額清空
            ResetFields(include65_66:=True, include6D:=False, include6F_6G:=False)

            'ddlIT61.Enabled = True

            BT5_BUTTON(age)

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    Private Sub BT6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT6.Click
        Try
            ddlTargetProject.SelectedValue = "ITP"

            'CheckFlight() '航班鎖控
            CheckFlight_NEW() '新航班鎖控
            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If

            '旅遊地區整合
            Dim COUNCODE_ALL As String = ""
            COUNCODE_ALL = CHECK_COUNCODE1()
            TRAVEL_A.SelectedValue = COUNCODE_ALL

            If TRAVEL_A.SelectedValue = "400" Then
                Throw New Exception("旅遊地區為國內，無法能選擇此方案")
            End If

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數

            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            '於後續BT6_BUTTON(age)內一併調整
            'Me.txtCase.Text = "J 海外基本型"

            IISSU_N.Enabled = True

            '將各方案金額清空
            ResetFields(include65_66:=False, include6D:=True, include6F_6G:=True)

            BT6_BUTTON(age)

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    Private Sub BT7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT7.Click
        Try
            ddlTargetProject.SelectedValue = "ITP"

            'CheckFlight() '航班鎖控
            CheckFlight_NEW() '新航班鎖控
            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If

            '旅遊地區整合
            Dim COUNCODE_ALL As String = ""
            COUNCODE_ALL = CHECK_COUNCODE1()
            TRAVEL_A.SelectedValue = COUNCODE_ALL

            If TRAVEL_A.SelectedValue = "400" Then
                Throw New Exception("旅遊地區為國內，無法能選擇此方案")
            End If

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數

            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            '於後續BT6_BUTTON(age)內一併調整
            'Me.txtCase.Text = "K 海外豪華型"

            IISSU_N.Enabled = True

            '將各方案金額清空
            ResetFields(include65_66:=False, include6D:=True, include6F_6G:=True)

            BT7_BUTTON(age)

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    Private Sub BT8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT8.Click
        Try
            ddlTargetProject.SelectedValue = "ITP"

            'CheckFlight() '航班鎖控
            CheckFlight_NEW() '新航班鎖控
            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If

            '旅遊地區整合
            Dim COUNCODE_ALL As String = ""
            COUNCODE_ALL = CHECK_COUNCODE1()
            TRAVEL_A.SelectedValue = COUNCODE_ALL

            If TRAVEL_A.SelectedValue = "400" Then
                Throw New Exception("旅遊地區為國內，無法能選擇此方案")
            End If

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數

            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            '於後續BT6_BUTTON(age)內一併調整
            'Me.txtCase.Text = "K 海外豪華型"

            IISSU_N.Enabled = True

            '將各方案金額清空
            ResetFields(include65_66:=False, include6D:=True, include6F_6G:=True)

            BT8_BUTTON(age)

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Private Sub BT9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT9.Click
        'K077 海外基本型
        '判斷被保人年齡

        Try
            '於後續BT9_BUTTON(age)內一併調整
            'lblTargetPlan.Text = "M 海外基本型"
            ddlTargetProject.SelectedValue = "ITP"

            'CheckFlight() '航班鎖控
            CheckFlight_NEW() '新航班鎖控
            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If
            '旅遊地區整合
            Dim COUNCODE_ALL As String = ""
            COUNCODE_ALL = CHECK_COUNCODE1()
            TRAVEL_A.SelectedValue = COUNCODE_ALL

            If TRAVEL_A.SelectedValue = "400" Then
                Throw New Exception("旅遊地區為國內，無法能選擇此方案")
            End If

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數
            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            IISSU_N.Enabled = True

            '將各方案金額清空
            ResetFields(include65_66:=False, include6D:=False, include6F_6G:=True)

            'ddlIT61.Enabled = True

            BT9_BUTTON(age)

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    Private Sub BT10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT10.Click
        'K077 海外超值型
        Try
            '於後續BT10_BUTTON(age)內一併調整
            'lblTargetPlan.Text = "N 海外超值型"

            ddlTargetProject.SelectedValue = "ITP"

            'CheckFlight() '航班鎖控
            CheckFlight_NEW() '新航班鎖控
            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If

            '旅遊地區整合
            Dim COUNCODE_ALL As String = ""
            COUNCODE_ALL = CHECK_COUNCODE1()
            TRAVEL_A.SelectedValue = COUNCODE_ALL

            If TRAVEL_A.SelectedValue = "400" Then
                Throw New Exception("旅遊地區為國內，無法能選擇此方案")
            End If

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數

            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            IISSU_N.Enabled = True

            '將各方案金額清空
            ResetFields(include65_66:=False, include6D:=False, include6F_6G:=True)

            'ddlIT61.Enabled = True

            BT10_BUTTON(age)

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    Private Sub BT11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT11.Click
        '旅行社_大陸保代 海外小資型
        Try
            '於後續BT11_BUTTON(age)內一併調整
            'lblTargetPlan.Text = "O 海外小資型"

            ddlTargetProject.SelectedValue = "ITP"

            'CheckFlight() '航班鎖控
            CheckFlight_NEW() '新航班鎖控
            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If

            '旅遊地區整合
            Dim COUNCODE_ALL As String = ""
            COUNCODE_ALL = CHECK_COUNCODE1()
            TRAVEL_A.SelectedValue = COUNCODE_ALL

            If TRAVEL_A.SelectedValue = "400" Then
                Throw New Exception("旅遊地區為國內，無法能選擇此方案")
            End If

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數

            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            IISSU_N.Enabled = True

            '將各方案金額清空
            ResetFields(include65_66:=False, include6D:=False, include6F_6G:=True)

            'ddlIT61.Enabled = True

            BT11_BUTTON(age)

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Private Sub BTIQP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTIQP.Click
        Try
            '於後續BT5_BUTTON(age)內一併調整
            'lblTargetPlan.Text = "D 海外超值型"

            ddlTargetProject.SelectedValue = "IQP"
            Me.txtCase.Text = "僅人身保險"
            Me.lblTargetPlan.Text = ""


            If ddlIT41.SelectedValue <> "" Then
                ddlIT4Y.SelectedValue = ""
                GetIT44FromIT41()
                GetIT43FromIT41()
            ElseIf ddlIT4Y.SelectedValue <> "" Then
                ddlIT41.SelectedValue = ""
                GetIT44FromIT4Y()
                GetIT43FromIT4Y()
            End If

            'CheckFlight() '航班鎖控
            CheckFlight_NEW() '新航班鎖控

            Dim ChkIns As String
            ChkIns = CheckBasic()

            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            GetDay() '計算天數

            '判斷被保人年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If

            '設定海突相關控制項顯示
            Dim dobj As DiseaseFields = New DiseaseFields()
            dobj.IT61 = True
            dobj.IT63 = True
            dobj.IT6E = False
            dobj.IT64 = False
            dobj.IT65 = True
            dobj.IT66 = True
            dobj.IT6D = False
            dobj.IT6F = False
            dobj.IT6G = False
            FieldsSwtich_Disease(dobj)

            ddlIT43.Enabled = False
            ddlIT43.SelectedValue = ""
            txtIT43.Text = ""

            GetDay() '計算天數

            '於後續BT5_BUTTON(age)內一併調整
            'Me.txtCase.Text = "D 海外超值型"

            IISSU_N.Enabled = True

            '將各方案金額清空
            ResetFields(include65_66:=True, include6D:=True, include6F_6G:=True)
            'ddlIT61.Enabled = True

            '將 方案資訊、保費小計、保費合計 整欄顯示
            divfuncinfoDisplay(True)

            '每個方案按鈕都加
            Me.UIAppendMessage("選擇不便險方案後，因「海外突發疾病保險金」會清空重置，請務必留意資料是否填妥。")
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

#End Region

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '	DirectToEKYC()
    'End Sub

    Private Sub chkAssignACC_CheckedChanged(sender As Object, e As EventArgs) Handles chkAssignACC.CheckedChanged
        Try
            If checkAssignACC().Item1 = False Then
                If chkAssignACC.Checked = False Then
                    chkAssignACC.Checked = True
                End If

                Throw New Exception("您無操作權限")
                Return
            End If

            '清空經辦代號的名稱
            labNIOfficer.Text = ""
            '清空方案資料
            lblTargetPlan.Text = ""
            txtCase.Text = ""
            '清空判斷[可保特殊方案(豪華型、標準型)]的變數，清空後會抓通路代碼重新判斷
            txtIsSPEChannel.Text = ""
            '已勾選：顯示方案按鈕&新增修改被保人按鈕；取消勾選：隱藏方案按鈕&新增修改被保人按鈕
            If chkAssignACC.Checked Then
                labChkMsg.InnerText = "確認通路代碼與帳號(若要調整請取消勾選)"
                divfunc.Attributes.Add("style", "display:inhert")
                divIISSU_NnE.Attributes.Add("style", "display:inhert")
            Else
                labChkMsg.InnerText = "確認通路代碼與帳號(勾選確認後，才可接續登打被保人)"
                divfunc.Attributes.Add("style", "display:none")
                divIISSU_NnE.Attributes.Add("style", "display:none")
            End If

            If chkAssignACC.Checked Then
                If String.IsNullOrEmpty(Me.txtCompanyid.Text.Trim) Then
                    chkAssignACC.Checked = False
                    Throw New Exception("[通路代碼]不得為空")
                    Return
                End If
                If String.IsNullOrEmpty(Me.txtUserid.Text.Trim) Then
                    chkAssignACC.Checked = False
                    Throw New Exception("[帳號]不得為空")
                    Return
                End If
                If String.IsNullOrEmpty(Me.txtIRECPESN_TxtIn.Text.Trim) Then
                    chkAssignACC.Checked = False
                    Throw New Exception("[登錄證字號]不得為空")
                    Return
                End If

                Me.txtCompanyid.Text = Me.txtCompanyid.Text.ToUpper().Trim
                Me.txtUserid.Text = Me.txtUserid.Text.ToUpper().Trim
                Dim KeyinUserID As String = Me.txtCompanyid.Text.Trim & "\" & Me.txtUserid.Text.Trim

                '檢核帳號是否有誤
                Dim dsAccount As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.account_check", KeyinUserID)
                If dsAccount.Tables(0).DefaultView.Count > 0 Then
                    '有找到帳號
                Else
                    chkAssignACC.Checked = False
                    Throw New Exception("[通路代碼]或[帳號]輸入錯誤，或是[帳號]不存在，請確認")
                End If

                '檢核帳號是否有旅遊險的投保權限
                Dim dsHAS As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.get_roleuser", KeyinUserID, "HAS_USER")
                If dsHAS.Tables(0).DefaultView.Count > 0 Then
                    '該帳號有HAS_USER權限
                Else
                    chkAssignACC.Checked = False
                    Throw New Exception("[帳號]無權限投保旅遊險，請確認")
                End If

                '取經辦代號
                Dim dsOfficer As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.getiofficer", KeyinUserID)
                If dsOfficer.Tables("Organization").Rows.Count > 0 Then
                    labIOFFICER.Text = dsOfficer.Tables("Organization").Rows(0).Item("Iofficer")
                Else
                    chkAssignACC.Checked = False
                    Throw New Exception("找不到本組織的[單位代號]或[經辦代號]")
                End If

                '判斷是使用者本人 或是 指定者用者，執行各別判斷
                If Me.ScsUserID = KeyinUserID Then
                    '如果是使用者本人
                    '若有填寫登錄證號，則將證號貼至內控資料區塊的登錄證號輸入欄
                    If String.IsNullOrEmpty(txtIRECPESN_TxtIn.Text.Trim) Then
                        txtIRECPESN.Text = txtIRECPESN_TxtIn.Text.Trim.ToUpper()
                    End If

                    'txtCompanyid & txtUserid & txtIRECPESN_TxtIn enable設為false
                    Me.txtCompanyid.Enabled = False
                    Me.txtUserid.Enabled = False
                    Me.txtIRECPESN_TxtIn.Enabled = False

                    '取得SKI 經辦名稱
                    If Me.ScsUserID.Substring(0, 4) = "SKI\" Then
                        Dim PSNiofficer As String = Me.ScsUserID.Substring(4, 6)
                        Dim dsRtn_Temp As DataSet = Me.CallCommonService("PUBService.bsSingleSQL", New String() {"PSNV_OFFICER_ALL", "IOFFICER ='" + PSNiofficer + "'", "IEMAIL,NNAME"})
                        If dsRtn_Temp.Tables(0).DefaultView.Count > 0 Then
                            labNIOfficer.Text = dsRtn_Temp.Tables(0).Rows(0)("NNAME")
                        End If
                    End If

                Else
                    '如果是指定帳號
                    '檢核該使用者有無權限管理登打的經辦代號；檢核登錄證字號
                    Dim sqlWhere As String = $"IMANAGER ='{(Me.ScsUserID.Split("\"))(1)}' AND iofficer = '{labIOFFICER.Text}' "
                    Dim dsRtn As DataSet = CallCommonService("PUBService.bsSingleSQL", New String() {"CMNP_OFF_SOURCE", sqlWhere, "*"})
                    If dsRtn Is Nothing OrElse dsRtn.Tables.Count = 0 OrElse dsRtn.Tables(0).Rows.Count = 0 Then
                        chkAssignACC.Checked = False
                        '    Throw New Exception($"管理人({Me.ScsUserID.Replace("SKI\", "")})沒有權限管理經辦代號({labNIOfficer.Text})")
                        Throw New Exception("沒有管理權限，僅可輸入使用者轄下經辦代號的帳號")
                    End If

                    '非SKI 取經辦名稱
                    Dim dsRtn_Temp2 As DataSet = Me.CallCommonService("PUBService.bsSingleSQL", New String() {"CMNP_OFF_SOURCE", "IOFFICER ='" + labIOFFICER.Text + "'", "NOFFICER"})
                    If dsRtn_Temp2.Tables(0).DefaultView.Count > 0 Then
                        labNIOfficer.Text = dsRtn_Temp2.Tables(0).Rows(0)("NOFFICER")
                    End If

                    'txtCompanyid & txtUserid & txtIRECPESN_TxtIn enable設為false
                    Me.txtCompanyid.Enabled = False
                    Me.txtUserid.Enabled = False
                    Me.txtIRECPESN_TxtIn.Enabled = False

                    '若有填寫登錄證號，則將證號貼至內控資料區塊的登錄證號輸入欄
                    If String.IsNullOrEmpty(txtIRECPESN_TxtIn.Text.Trim) = False Then
                        txtIRECPESN.Text = txtIRECPESN_TxtIn.Text.Trim.ToUpper()
                    End If

                End If

                '經辦對應的設定、卡控（此function設定原本在Page_Load()的卡控）
                LogicAfterAssignACC()

            Else
                '將被保人資料清空
                If Me.StateMain.Tables("IQPISSU").Rows.Count > 0 Then
                    Me.StateMain.Tables("IQPISSU").Clear()
                End If
                If Me.StateMain.Tables("IQPOBJ").Rows.Count > 0 Then
                    Me.StateMain.Tables("IQPOBJ").Clear()
                End If
                If Me.StateMain.Tables.Contains("HASM_FLIGHT_DETAIL") AndAlso Me.StateMain.Tables("HASM_FLIGHT_DETAIL").Rows.Count > 0 Then
                    Me.StateMain.Tables("HASM_FLIGHT_DETAIL").Clear()
                End If

                labNIOfficer.Text = ""
                CheckSNNN.Text = ""
                StateMain.AcceptChanges()
                BaseSetChildFormData(dgdMain, StateMain)
                BaseSetChildFormData(dgdFlight, StateMain)
                IISSU_Clear()

                'txtCompanyid & txtUserid & txtIRECPESN_TxtIn enable設為false
                Me.txtCompanyid.Enabled = True
                Me.txtUserid.Enabled = True
                Me.txtIRECPESN_TxtIn.Enabled = True

                '復原特定通路別的初始設定
                '農經KEJ0初始設定
                dllGetIrec.Visible = False
                'KKEU初始設定
                txtIrecunit.Text = ""

            End If

        Catch ex As Exception
            ProcessException(ex)
        End Try
    End Sub

    Private Sub LogicAfterAssignACC()
        'Page_Load內的控制在此function重新判斷、控制

#Region "指定的經辦代號可否顯示特殊方案"

        If checkSpecialChannel().Item1 = True Then
            BT3.Attributes.Add("style", "display:inline")
            BT5.Attributes.Add("style", "display:inline")
            BT6.Attributes.Add("style", "display:inline")
            BT7.Attributes.Add("style", "display:inline")
            BT9.Attributes.Add("style", "display:none")
            BT10.Attributes.Add("style", "display:none")
            BT11.Attributes.Add("style", "display:none")

            If TRAVEL_A.SelectedValue = "400" Then
                BT8.Attributes.Add("style", "display:none")
                BTIQP.Attributes.Add("style", "display:inline")
            Else
                BT8.Attributes.Add("style", "display:inline")
                BTIQP.Attributes.Add("style", "display:none")
            End If
        ElseIf checkChannelO().Item1 Then
            BT3.Attributes.Add("style", "display:inline")
            BT5.Attributes.Add("style", "display:inline")
            BT6.Attributes.Add("style", "display:inline")
            BT7.Attributes.Add("style", "display:inline")
            BT9.Attributes.Add("style", "display:none")
            BT10.Attributes.Add("style", "display:none")
            BT11.Attributes.Add("style", "display:inline")

            If TRAVEL_A.SelectedValue = "400" Then
                BT8.Attributes.Add("style", "display:none")
                BTIQP.Attributes.Add("style", "display:inline")
            Else
                BT8.Attributes.Add("style", "display:inline")
                BTIQP.Attributes.Add("style", "display:none")
            End If
        ElseIf checkChannelMN().Item1 Then
            BT3.Attributes.Add("style", "display:none")
            BT5.Attributes.Add("style", "display:none")
            BT6.Attributes.Add("style", "display:none")
            BT7.Attributes.Add("style", "display:none")
            BT8.Attributes.Add("style", "display:none")
            BT9.Attributes.Add("style", "display:inline")
            BT10.Attributes.Add("style", "display:inline")
            BT11.Attributes.Add("style", "display:none")
            BTIQP.Attributes.Add("style", "display:inline")
        Else
            BT3.Attributes.Add("style", "display:inline")
            BT5.Attributes.Add("style", "display:inline")
            BT6.Attributes.Add("style", "display:none")
            BT7.Attributes.Add("style", "display:none")
            BT8.Attributes.Add("style", "display:none")
            BT9.Attributes.Add("style", "display:none")
            BT10.Attributes.Add("style", "display:none")
            BT11.Attributes.Add("style", "display:none")
            BTIQP.Attributes.Add("style", "display:inline")
        End If

#End Region

#Region "農經資料查詢"
        If Me.txtCompanyid.Text = "KEJ0" Then
            GetIrecItems(labIOFFICER.Text)
        Else
            dllGetIrec.Visible = False
        End If
#End Region


#Region "B2B作業程式，當通路別=KKSB增加功能 "
        If Me.txtCompanyid.Text = "KKSB" Then
            If (Me.txtUserid.Text.Trim).Substring(0, 4) = "KKSB" Then
                '帳號前4碼非[KKSB]者,將B2B帳號傳入內控資料[通路人員代號]欄位。
                txtIRECPESN.Text = ""
            Else
                txtIRECPESN.Text = Me.txtUserid.Text
            End If
        End If
#End Region

#Region "KKEU，將070705帶入登錄證號欄位(txtIRECPESN)"
        '當通路別=KKEU，通路單位預設070705
        If Me.txtCompanyid.Text = "KKEU" Then
            '帳號前4碼[KKEU]通路單位預設070705。
            txtIrecunit.Text = "070705"
        Else
            txtIrecunit.Text = String.Empty
        End If
#End Region

#Region "取得鎖定集體費率權限"
        Dim HAS_GROUP As String = ""
        Dim dsSingG As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_GROUP' AND CODE ='" + GetUserIDFromTextBox() + "'", "CODE"})

        If Not IsNothing(dsSingG) AndAlso dsSingG.Tables.Count > 0 AndAlso dsSingG.Tables(0).Rows.Count > 0 Then
            HAS_GROUP = dsSingG.Tables(0).Rows(0)("CODE").ToString.Trim
        End If
        If HAS_GROUP <> "" Then
            CHK5.Checked = True
            CHK5_G.Text = "直費限定"
            CHK5.Enabled = False
        Else
            CHK5.Checked = False
            CHK5_G.Text = ""
            CHK5.Enabled = False
        End If
#End Region
    End Sub

    Private Sub GetIrecItems(IOfficerStr As String)
        Dim sArr As String() = New String() {"B2B_GETIREC", IOfficerStr}
        Dim dsGETIREC As DataSet = CallCommonService("CarService.bsCarCommonService", sArr)
        If dsGETIREC.Tables(0).Rows.Count > 0 Then
            dllGetIrec.Visible = True
            For Each dr As DataRow In dsGETIREC.Tables(0).Rows
                Dim b As New ListItem
                b.Text = dr.Item("DEPNAME").ToString.Trim + " " + dr.Item("NAME").ToString.Trim + "(" + dr.Item("ID").ToString.Trim + ")"
                b.Value = dr.Item("ID").ToString.Trim + "|" + dr.Item("NAME").ToString.Trim + "|" + dr.Item("DEPNO").ToString.Trim
                dllGetIrec.Items.Add(b)
            Next
        Else
            dllGetIrec.Visible = False
        End If
    End Sub

    ''' <summary> 新增被保人/修改 </summary>
    Private Sub IISSU_N_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IISSU_N.Click
        Try
            '因應需求單F115030200 【特定通路可以指定要保的通路代碼與帳號】，先檢查chkAssignACC是否勾選，因為勾選or取消勾選會執行對應檢核及動作
            If chkAssignACC.Checked = False Then
                Throw New Exception("請先勾選【確認通路代碼與帳號】。（保險期間上方）")
            End If

            Dim TargetProject As String = ddlTargetProject.SelectedValue.Trim '20250320
            Dim CaseCode As String = ""
            If String.IsNullOrEmpty(txtCase.Text.Trim) Then Throw New Exception("必需選擇方案。若無指定方案，請點選按鈕〔僅人身保險〕。")

            'If ddlTargetProject.SelectedValue.Trim = "ITP" Then
            '	If lblTargetPlan.Text.Trim = "" Then Throw New Exception("ITP必需選擇方案。方案不可空白。")

#Region "1140415 因為當user已新增被保人，而又於ITP & IQP之間切換時會因方案差異造成錯誤，故取消方案檢查卡控，因最後依然要點選保費試算"
            'If StateMain.Tables("IQPISSU") IsNot Nothing AndAlso StateMain.Tables("IQPISSU").Rows.Count > 0 Then
            '	If Not StateMain.Tables("IQPISSU").AsEnumerable().FirstOrDefault().Field(Of String)("IPLANNO").
            '	ToString.Trim() = If(TargetProject = "ITP", lblTargetPlan.Text.Trim.Substring(0, 1), "") Then Throw New Exception("[投保方案]需相同！！！！")
            'End If
#End Region

            'End If

            Dim ChkIns As String
            ChkIns = CheckBasic()
            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            CommonFunc.Relationship_addpeople_btn(CType(Me.StateMain, beITP101000).Tables("IQPISSU"), Relationship, txtIins_IINS2, txtIins_ISSU2, txtRelationshipOther)

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            If TargetProject = "ITP" Then
                Dim ChkAreaLock As String
                ChkAreaLock = CheckAreaLock()
                If ChkAreaLock <> String.Empty Then
                    Throw New Exception(ChkAreaLock)
                End If
            End If

            Dim ChkIns2 As String
            ChkIns2 = CheckInsert()
            If ChkIns2 <> String.Empty Then
                Throw New Exception(ChkIns2)
            End If

            'A114123106)全險「通訊地址、電話」管控需求－PUB地址卡控程式呼叫說明------------------------------------------------------------------------
            '叫用  PUBService.bsPubAddrCHK  範例
            '(00) 0檢查用/1送審用(必)        (06)要保人郵遞區號     (12)被保人地址   
            '(01)險別(C,M,F,A,I,E,H) (必)    (07)要保人地址      　 (13)被保人電話區碼  
            '(02)資料來源(B2B/B2C/MIS) (必)  (08)要保人電話區碼     (14)被保人電話         　
            '(03)報價/出單/續保/批單(必)     (09)要保人電話         　
            '(04)PROGRANID(必)               (10)被保人統編       　
            '(05)要保人統編                  (11)被保人郵遞區號   
            'IISSU_N_Click
            Dim ds99 = CallCommonService("PUBService.bsPubAddrCHK", New String() {"0", "H", "B2B", "1", "BHAS101003",
            txtIins_IINS2.Text.Trim, IINS_AREA.Text.Trim, IINS_ADDR.Text.Trim, "", IINS_TEL.Text.Trim,
            txtIins_ISSU2.Text.Trim, IISSU_AREA.Text.Trim, IISSU_ADDR.Text.Trim, "", IISSU_TEL.Text.Trim})
            If ds99.Tables("RESULT").Rows(0)("PASS") = "N" AndAlso ds99.Tables("ERROR").Rows.Count > 0 Then
                UIAppendMessage(ds99.Tables("ERROR").Rows(0)("NAME"))
                Throw New Exception(ds99.Tables("ERROR").Rows(0)("NAME"))
            End If
            'A114123106)全險「通訊地址、電話」管控需求－PUB地址卡控程式呼叫說明------------------------------------------------------------------------


            CheckSNNN.Text = ""

            ''判斷被保人年齡
            '判斷年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If
            '檢查時間  生效、到期日同一天，要控管生效時需小於到期時。

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            If bolK077() Then
                If txtIrecunit.Text.Trim = "" Then
                    Throw New Exception("通路單位不可為空!!!")
                End If
            End If

            If TargetProject = "IQP" Then
                '海途要卡
                If TRAVEL_A.SelectedValue = "400" And ddlIT61.SelectedValue <> "" Then
                    Throw New Exception("旅遊地區為[400台灣地區及離島地區時]不得承接海突。")
                ElseIf TRAVEL_A.SelectedValue = "400" And (ddlIT63.SelectedValue <> "" Or ddlIT6E.SelectedValue <> "") Then
                    Throw New Exception("旅遊地區為[400台灣地區及離島地區時]不得承接海突。")
                End If
            End If

            '被保險人與要保人關係沒控管
            '(下例為要/被保險人是相同，但卻可以選擇本人以外的選項；反之亦同)
            If txtIins_IINS2.Text = txtIins_ISSU2.Text Then
                If Relationship.SelectedValue <> "1" Then
                    Throw New Exception("要保人與被保人相同時，與要保人關係請選擇 1.本人")
                End If
            End If
            '為本人時候的檢查
            Dim strErr As String = ""
            Me.IsInputErr = False

            If Relationship.SelectedValue = "1" Then

                If txtIins_IINS2.Text <> txtIins_ISSU2.Text Then
                    Me.IsInputErr = True
                    strErr &= "與要保人關係選擇[本人]時，要保人與被保人[身分證字號]要相同. " & Environment.NewLine

                End If
                If NAME_IINS.Text <> NAME_IISSU.Text Then
                    Me.IsInputErr = True
                    strErr &= "與要保人關係選擇[本人]時，要保人與被保人[姓名]要相同. " & Environment.NewLine
                End If
                'Birtdy_IINS
                If Birtdy_IINS.DateText <> UcRocDbirth.DateText Then
                    Me.IsInputErr = True
                    strErr &= "與要保人關係選擇[本人]時，要保人與被保人[生日]要相同. " & Environment.NewLine
                End If

            End If
            If Me.IsInputErr = True Then
                Throw New Exception(strErr)
            End If

            If TargetProject = "ITP" AndAlso Me.txtCase.Text = "" Then
                Throw New Exception("請先選擇[投保方案]")
            ElseIf TargetProject = "ITP" AndAlso Me.txtCase.Text <> "" Then
                CaseCode = (Me.txtCase.Text).Substring(0, 1)
            End If

            Dim ChkStr As String = String.Empty

            If txtISEQ.Text = "" Then
                '新增一筆 新的
                '先檢查 是不是有重複IISSU

                Dim IISSU_D As String = txtIins_ISSU2.Text.Trim
                For Each drD1 As DataRow In StateMain.Tables("IQPISSU").Select("IISSU = '" & IISSU_D & "'")
                    Throw New Exception("被保人ID重複")
                Next

                '20260116 再次加入檢核〔 [投保方案]需要與主被保人相同〕 '20250320 [投保方案]需要與主被保人相同 此段已經無意義 加入期間又更改．．．　在保費試算在重新更正
                If StateMain.Tables("IQPISSU").Rows.Count > 0 Then
                    '先檢查 方案要一樣
                    Dim ICASE As String = Me.txtCase.Text.Substring(0, 1)
                    Dim ICASE2 As String = ""
                    If txtCase.Text = "僅人身保險" Then ICASE = ""

                    For Each drCASE2 As DataRow In StateMain.Tables("IQPISSU").Select("ISEQ = '1' ")
                        If drCASE2("iPlanNO").ToString.Length >= 1 Then
                            ICASE2 = drCASE2("iPlanNO").ToString.Substring(0, 1)
                        End If
                    Next

                    If ICASE <> ICASE2 Then
                        Throw New Exception("[投保方案]需要與主被保人相同")
                    End If
                End If

                '1 本人
                '再檢查 是不是有重複 本人3.	被保險人與要保人關係沒控管(下例為要/被保險人是相同，但卻可以選擇本人以外的選項；反之亦同)
                Dim Relation As String = HttpUtility.HtmlEncode(Relationship.SelectedValue)
                If Relationship.SelectedValue = "1" Then
                    For Each drD2 As DataRow In StateMain.Tables("IQPISSU").Select("NINSREL = '1'")
                        Throw New Exception(" 與要保人關係 本人 不能重複")
                    Next
                End If

                '檢核生日 要跟下拉選單的IT41 IT4Y相同
                If age < 15 Then
                    If ddlIT41.SelectedValue <> "" Then
                        ddlIT41.SelectedValue = ""
                        txtIT41.Text = "0"
                        Throw New Exception(" 未滿15歲 請選擇 意外喪葬費用或意外失能保險金")
                    ElseIf ddlIT41.SelectedValue <> "" Then
                        ddlIT41.SelectedValue = "0"
                        txtIT41.Text = "0"
                    End If
                Else
                    If ddlIT4Y.SelectedValue <> "" Then
                        ddlIT4Y.SelectedValue = ""
                        txtIT4Y.Text = "0"
                        Throw New Exception(" 滿15足歲 請選擇意外身故失能保險金")
                    ElseIf ddlIT4Y.SelectedValue <> "" Then
                        ddlIT4Y.SelectedValue = "0"
                        txtIT4Y.Text = "0"
                    End If
                End If

                'ITP 年齡檢核
                Dim sSolution, IADD1, IADD2, IADD3, IADD4, IADD5, IADD6 As String

                If ddlIT41.SelectedValue <> "" Then
                    sSolution = ddlIT41.SelectedValue.Trim + "0000"
                ElseIf ddlIT4Y.SelectedValue <> "" And ddlIT4Y.SelectedValue <> "0" Then
                    sSolution = CDec(ddlIT4Y.SelectedValue.Trim) * 10000
                Else
                    If age < 15 Then
                        If ddlIT4Y.SelectedValue = "" Or ddlIT4Y.SelectedValue = "0" Then
                            If ddl_IISSU.SelectedValue = "5" And TRAVEL_A.SelectedValue = "400" Then
                                Throw New Exception("國內地區喪葬費用保額不可為0.")
                            Else
                                If ((checkSpecialChannel().Item1 = True Or checkChannelO().Item1 = True) AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L")) Then
                                    If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or ddlIT64.SelectedValue = "") Then
                                        Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                    End If

                                ElseIf checkChannelO().Item1 = True AndAlso CaseCode = "O" Then
                                    If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or (ddlIT61.SelectedValue = "" And ddlIT64.SelectedValue = "")) Then
                                        Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                    End If

                                ElseIf (CaseCode <> "J" And CaseCode <> "K" And CaseCode <> "L") AndAlso checkChannelMN.Item1 Then
                                    If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or (ddlIT63.SelectedValue = "" And ddlIT61.SelectedValue = "" And ddlIT64.SelectedValue = "")) Then
                                        Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                    End If

                                ElseIf (CaseCode <> "J" And CaseCode <> "K" And CaseCode <> "L") Then
                                    If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or (ddlIT63.SelectedValue = "" And ddlIT61.SelectedValue = "" And ddlIT6E.SelectedValue = "")) Then
                                        Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                    End If

                                End If
                            End If
                        End If
                    Else
                        Throw New Exception("主保險契約:[旅行平安保險]請選擇[保險金額]")
                    End If
                End If

                If ddlIT42.SelectedValue <> "" Then
                    IADD1 = ddlIT42.SelectedValue.Trim + "0000"
                Else
                    IADD1 = 0
                End If
                If ddlIT61.SelectedValue <> "" Then 'And rad_Disease_N.Checked = True Then
                    IADD2 = ddlIT61.SelectedValue.Trim + "0000"
                Else
                    IADD2 = 0
                End If
                If ddlIT44.SelectedValue <> "" Then
                    IADD3 = ddlIT44.SelectedValue.Trim + "0000"
                Else
                    IADD3 = 0
                End If
                If ddlIT63.SelectedValue <> "" Then 'And rad_Disease_Y.Checked = True Then
                    IADD4 = ddlIT63.SelectedValue.Trim + "0000"
                Else
                    IADD4 = 0
                End If
                If ddlIT6E.SelectedValue <> "" Then
                    IADD4 = ddlIT6E.SelectedValue.Trim + "0000"
                ElseIf ddlIT63.SelectedValue = "" Then
                    IADD4 = 0
                End If
                If ddlIT43.SelectedValue <> "" Then
                    IADD5 = ddlIT43.SelectedValue.Trim + "0000"
                Else
                    IADD5 = 0
                End If
                If ddlIT64.SelectedValue <> "" Then
                    IADD6 = ddlIT64.SelectedValue.Trim + "0000"
                Else
                    IADD6 = 0
                End If

                If (checkSpecialChannel().Item1 = True AndAlso TargetProject = "ITP" AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L")) Or
                   (checkChannelO().Item1 = True AndAlso TargetProject = "ITP" AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L" Or CaseCode = "O")) Then
                    ChkStr = CheckAge_ITP_SPEChannel(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate(), sSolution, IADD1, IADD2, "", ddl_IISSU.SelectedValue.ToString, txtCase.Text)
                Else
                    ChkStr = CheckAge_ITP(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate(), sSolution, IADD1, IADD2, CaseCode, ddl_IISSU.SelectedValue.ToString) 'Me.txtCase.Text.Substring(0, 1)
                End If

                If ChkStr <> String.Empty Then
                    Throw New Exception(ChkStr)
                End If

                Dim ChkStr2 As String = String.Empty
                If (checkSpecialChannel().Item1 = True AndAlso TargetProject = "ITP" AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L")) Or
                   (checkChannelO().Item1 = True AndAlso TargetProject = "ITP" AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L" Or CaseCode = "O")) Then
                    ChkStr2 = CheckiPlanNo_SPEChannel(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate(), txtCase.Text, sSolution, IADD1, IADD2, IADD3, IADD6, ddl_IISSU.SelectedValue.ToString, IADD5)
                Else
                    'ChkStr2 = CheckiPlanNo(Me.txtCase.Text.Substring(0, 1), sSolution, IADD1, IADD2, IADD3)
                    ChkStr2 = CheckiPlanNo(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate(), "", sSolution, IADD1, IADD2, IADD3, If(checkChannelMN().Item1 AndAlso TargetProject = "ITP", IADD6, IADD4), ddl_IISSU.SelectedValue.ToString, IADD5) 'Me.txtCase.Text.Substring(0, 1)
                End If

                If ChkStr2 <> String.Empty Then
                    Throw New Exception(ChkStr2)
                End If

                'F115060500 因投保項目而調整保險期間的天數
                ChkStr2 = String.Empty
                ChkStr2 = CheckPlanNTime(sSolution, IADD1, IADD2, IADD4, IADD6, IADD4)
                If ChkStr2 <> String.Empty Then
                    Throw New Exception(ChkStr2)
                End If

                Dim ds As beITP101000 = CType(Me.StateMain, beITP101000)
                Dim dr As beITP101000.IQPISSURow
                dr = ds.IQPISSU.NewIQPISSURow
                If Relationship.SelectedValue = "0" Then dr("JOBOFFICE") = txtRelationshipOther.Text.Trim 'F114010802 旅遊險b2b被保人與要保人關係#9078
                Dim iindex As Integer = ds.IQPISSU.DefaultView.Count
                If iindex = 0 Then
                    dr("ISEQ") = "1"
                    txtMIISSU.Text = txtIins_ISSU2.Text.Trim
                    txtMNIISSU.Text = NAME_IISSU.Text.Trim

                    If ddl_IISSU.SelectedValue = "5" Then
                        Throw New Exception("主被保險人必須為本國人或擁有居留證之外國人")
                    End If

                Else
                    For Each drISEQ As DataRow In StateMain.Tables("IQPISSU").Select("'1' = '1'", "ISEQ")
                        dr("ISEQ") = drISEQ("ISEQ")
                    Next
                    'If dr("ISEQ") <> "" Then
                    If IsDBNull(dr("ISEQ")) = False Then
                        dr("ISEQ") = dr("ISEQ") + 1
                    Else
                        dr("ISEQ") = iindex + 1
                    End If
                End If

                dr("NISSU") = NAME_IISSU.Text.Trim
                dr("IISSU") = txtIins_ISSU2.Text.Trim
                dr("NISSUBD") = UcRocDbirth.DateText.Trim

                dr("IISSU_TP") = ddl_IISSU.SelectedValue.Trim
                dr("IISSU_ZIP") = IISSU_AREA.Text.Trim
                dr("AISSU") = IISSU_ADDR.Text.Trim
                dr("NISSU_TEL") = IISSU_TEL.Text.Trim
                dr("NISSUMAIL") = EMAIL.Text.Trim
                dr("NINSREL") = Relationship.SelectedValue.Trim

                dr("NAMEREL") = Benefi.Text.Trim
                dr("IDREL") = Benefi_ID.Text.Trim
                dr("TYPEREL") = Relationship2.SelectedValue.Trim

                dr("NAMEREL2") = Benefi2.Text.Trim
                dr("IDREL2") = Benefi_ID2.Text.Trim
                dr("TYPEREL2") = Relationship22.SelectedValue.Trim

                dr("NAMEREL3") = Benefi3.Text.Trim
                dr("IDREL3") = Benefi_ID3.Text.Trim
                dr("TYPEREL3") = Relationship23.SelectedValue.Trim

                dr("RELMOTHED") = Me.ddlRELMOTHED.SelectedValue.Trim

                dr("iPlanNo") = If(TargetProject = "ITP", Me.txtCase.Text.Substring(0, 1), "") '20250320 決定方案項目
                dr("TELREL") = txtTELREL.Text.Trim

                dr("ZIPREL") = txtZIPREL.Text.Trim
                dr("AREL") = txtAREL.Text.Trim
                dr("IISSU_SEX") = ddlSex.SelectedValue.Trim
                If ddl_IISSU.SelectedValue = "1" Then
                    dr("IISSU_COUNTRY") = "TW"
                ElseIf ddl_IISSU.SelectedValue = "5" Then
                    dr("IISSU_COUNTRY") = ddlIISSU_COUNTRY.SelectedValue.ToString
                Else
                    dr("IISSU_COUNTRY") = "XG"
                End If
                dr("JOBCONTENT") = txtJOB.Text

                'F114081901 要求海外地區 英文姓名 & 護照號碼不得擇一填寫。原本只要求勾選申請﹝申根文件﹞需填寫，但81歲亦有限制
                If TRAVEL_A.SelectedValue <> "400" And age <= 81 Then
                    dr("PASSPORT_NAME") = PASSPORT_NAME.Text.Trim
                    dr("PASSPORT_NUMBER") = PASSPORT_NUMBER.Text.Trim
                End If

                '是否懷孕 20221227
                If isEMP_Y.Checked = True Then
                    dr("EMP") = "Y"
                ElseIf isEMP_N.Checked = True Then
                    dr("EMP") = "N"
                End If
                If dr("IISSU_SEX") = "1" And dr("EMP") = "Y" Then
                    Throw New Exception("被保險人是否為孕婦? 選擇性別錯誤")
                End If
                If dr("EMP") = "Y" Then
                    '懷孕無法投保海突
                    If ddlIT61.SelectedValue <> "" Or ddlIT63.SelectedValue <> "" Or ddlIT64.SelectedValue <> "" Then
                        Throw New Exception("被保險人為孕婦，建議不投保海外突發疾病保險金")
                    End If
                End If

                Dim a, b, c, d As Integer

                If ddlIT41.SelectedValue <> "" Then
                    dr("ISSU_1") = ddlIT41.SelectedValue.Trim + "0000"
                    a = CInt(txtIT41.Text)
                ElseIf ddlIT4Y.SelectedValue <> "" Then
                    dr("ISSU_1") = CInt(CDec(ddlIT4Y.SelectedValue.Trim) * 10000)
                    a = CInt(txtIT4Y.Text)
                End If
                If ddlIT42.SelectedValue <> "" Then
                    dr("ISSU_2") = ddlIT42.SelectedValue.Trim + "0000"
                End If
                If ddlIT61.SelectedValue <> "" Then
                    dr("ISSU_3") = ddlIT61.SelectedValue.Trim + "0000"
                End If
                If ddlIT63.SelectedValue <> "" Then
                    dr("ISSU_3D") = ddlIT63.SelectedValue.Trim + "0000"
                End If
                If ddlIT6E.SelectedValue <> "" Then
                    dr("ISSU_3D") = ddlIT6E.SelectedValue.Trim + "0000"
                End If
                If ddlIT64.SelectedValue <> "" Then
                    dr("ISSU_IT64") = ddlIT64.SelectedValue.Trim + "0000"
                End If
                If ddlIT44.SelectedValue <> "" Then
                    dr("ISSU_IT44") = ddlIT44.SelectedValue.Trim + "0000"
                End If
                If ddlIT43.SelectedValue <> "" Then
                    dr("ISSU_IT43") = ddlIT43.SelectedValue.Trim + "0000"
                End If

                If txtIT42.Text <> "" Then
                    b = CInt(txtIT42.Text)
                End If
                If txtIT44.Text <> "" Then
                    d = CInt(txtIT44.Text)
                End If
                If txtIT61.Text <> "" Then
                    c = CInt(txtIT61.Text)
                End If
                '===================================================================

                '新增InsertIQPOBJ
                InsertIQPOBJ(dr("iPlanNo"), "")

                '重算總保費
                Dim TOTAMT As Integer = 0
                Dim TOTPREM As Integer = 0
                For Each drT As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & dr("IISSU") & "'")
                    'drD.Delete()
                    TOTPREM += CInt(drT("INSPREM"))
                Next

                dr("ISSU_4") = TOTPREM

                ds.IQPISSU.AddIQPISSURow(dr)
                BaseSetChildFormData(Me.dgdMain, Me.StateMain)

                '清空投保方案 
                '20160202 改成不清空投保方案
                'Me.txtCase.Text = ""
                txtISEQ.Text = ""
                txtIISSU_Ori.Text = ""
                '新增修改完之後 清除資料
                IISSU_Clear()

                '將 方案資訊、保費小計、保費合計 整欄顯示
                divfuncinfoDisplay(True)
                '隱藏保費資訊
                ShowPremInfoOrNot(False)
            Else
                Throw New Exception("請確認是否應執行[修改被保人]。否則請重新登打[被保人名稱]，或檢核被保人資料，是否與已新增的被保人重複。")
            End If

        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    Protected Sub IISSU_E_Click(sender As Object, e As EventArgs) Handles IISSU_E.Click
        'IQPMAIN先AcceptChanges，避免當exception後的RejectChanges，IQPMAIN被Reject成空的導致底層funtion BasePreRender()出錯
        'IQPISSU & IQPOBJ先AcceptChanges，避免當exception後的RejectChanges，被保人資料被清空
        StateMain.Tables("IQPMAIN").AcceptChanges()
        StateMain.Tables("IQPISSU").AcceptChanges()
        StateMain.Tables("IQPOBJ").AcceptChanges()

        Try
            '因應需求單F115030200 【特定通路可以指定要保的通路代碼與帳號】，先檢查chkAssignACC是否勾選，因為勾選or取消勾選會執行對應檢核及動作
            If chkAssignACC.Checked = False Then
                Throw New Exception("請先勾選【確認通路代碼與帳號】。（保險期間上方）")
            End If

            Dim TargetProject As String = ddlTargetProject.SelectedValue.Trim '20250320
            Dim CaseCode As String = ""
            If String.IsNullOrEmpty(txtCase.Text.Trim) Then Throw New Exception("必需選擇方案。" & Environment.NewLine + "若無指定方案，請點選按鈕〔僅人身保險〕。")
            'If ddlTargetProject.SelectedValue.Trim = "ITP" Then
            '	If lblTargetPlan.Text.Trim = "" Then Throw New Exception("ITP必需選擇方案。方案不可空白。")

#Region "1140415 因為當user已新增被保人，而又於ITP & IQP之間切換時會因方案差異造成錯誤，故取消方案檢查卡控，因最後依然要點選保費試算"
            'If StateMain.Tables("IQPISSU") IsNot Nothing AndAlso StateMain.Tables("IQPISSU").Rows.Count > 0 Then
            '	If Not StateMain.Tables("IQPISSU").AsEnumerable().FirstOrDefault().Field(Of String)("IPLANNO").
            '	ToString.Trim() = If(TargetProject = "ITP", lblTargetPlan.Text.Trim.Substring(0, 1), "") Then Throw New Exception("[投保方案]需相同！！！！")
            'End If
#End Region

            'End If

            Dim ChkIns As String
            ChkIns = CheckBasic()
            If ChkIns <> String.Empty Then
                Throw New Exception(ChkIns)
            End If

            CommonFunc.Relationship_addpeople_btn(CType(Me.StateMain, beITP101000).Tables("IQPISSU"), Relationship, txtIins_IINS2, txtIins_ISSU2, txtRelationshipOther)

            Dim ChkTime As String
            ChkTime = CheckTime()
            If ChkTime <> String.Empty Then
                Throw New Exception(ChkTime)
            End If
            'End If

            If TargetProject = "ITP" Then
                Dim ChkAreaLock As String
                ChkAreaLock = CheckAreaLock()
                If ChkAreaLock <> String.Empty Then
                    Throw New Exception(ChkAreaLock)
                End If
            End If

            Dim ChkIns2 As String
            ChkIns2 = CheckInsert()
            If ChkIns2 <> String.Empty Then
                Throw New Exception(ChkIns2)
            End If

            'A114123106)全險「通訊地址、電話」管控需求－PUB地址卡控程式呼叫說明------------------------------------------------------------------------
            '叫用  PUBService.bsPubAddrCHK  範例
            '(00) 0檢查用/1送審用(必)        (06)要保人郵遞區號     (12)被保人地址   
            '(01)險別(C,M,F,A,I,E,H) (必)    (07)要保人地址      　 (13)被保人電話區碼  
            '(02)資料來源(B2B/B2C/MIS) (必)  (08)要保人電話區碼     (14)被保人電話         　
            '(03)報價/出單/續保/批單(必)     (09)要保人電話         　
            '(04)PROGRANID(必)               (10)被保人統編       　
            '(05)要保人統編                  (11)被保人郵遞區號   
            'IISSU_N_Click
            Dim ds99 = CallCommonService("PUBService.bsPubAddrCHK", New String() {"0", "H", "B2B", "1", "BHAS101003",
            txtIins_IINS2.Text.Trim, IINS_AREA.Text.Trim, IINS_ADDR.Text.Trim, "", IINS_TEL.Text.Trim,
            txtIins_ISSU2.Text.Trim, IISSU_AREA.Text.Trim, IISSU_ADDR.Text.Trim, "", IISSU_TEL.Text.Trim})
            If ds99.Tables("RESULT").Rows(0)("PASS") = "N" AndAlso ds99.Tables("ERROR").Rows.Count > 0 Then
                UIAppendMessage(ds99.Tables("ERROR").Rows(0)("NAME"))
                Throw New Exception(ds99.Tables("ERROR").Rows(0)("NAME"))
            End If
            'A114123106)全險「通訊地址、電話」管控需求－PUB地址卡控程式呼叫說明------------------------------------------------------------------------


            CheckSNNN.Text = ""

            ''判斷被保人年齡
            '判斷年齡
            Dim age As String = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            txtDBIRTH.Text = age
            If Benefi.Text = "" Or Benefi.Text = "無身故保險金" Then
                Benefi.Text = "法定繼承人"
            End If
            '檢查時間  生效、到期日同一天，要控管生效時需小於到期時。

            '出生日期要小於輸入日期
            If Me.UcRocDbirth.GetDate() > Today() Then
                Throw New Exception("出生日期 不得小於 輸入日期")
            End If

            If bolK077() Then
                If txtIrecunit.Text.Trim = "" Then
                    Throw New Exception("通路單位不可為空!!!")
                End If
            End If

            '判斷孕婦相關
            If ddlSex.SelectedValue.Trim = "1" And isEMP_Y.Checked Then
                Throw New Exception("被保險人是否為孕婦? 選擇性別錯誤")
            End If
            If isEMP_Y.Checked Then
                '懷孕無法投保海突
                If ddlIT61.SelectedValue <> "" Then
                    Throw New Exception("被保險人為孕婦，建議不投保海外突發疾病保險金")
                ElseIf (ddlIT63.SelectedValue <> "" Or ddlIT64.SelectedValue <> "" Or ddlIT6E.SelectedValue <> "") Then
                    Throw New Exception("被保險人為孕婦，建議不投保海外突發疾病保險金")
                End If
            End If

            If TargetProject = "IQP" Then
                '海途要卡
                If TRAVEL_A.SelectedValue = "400" And ddlIT61.SelectedValue <> "" Then
                    Throw New Exception("旅遊地區為[400台灣地區及離島地區時]不得承接海突。")
                ElseIf TRAVEL_A.SelectedValue = "400" And (ddlIT63.SelectedValue <> "" Or ddlIT6E.SelectedValue <> "" Or ddlIT64.SelectedValue <> "") Then
                    Throw New Exception("旅遊地區為[400台灣地區及離島地區時]不得承接海突。")
                End If
            End If

            '被保險人與要保人關係沒控管
            '(下例為要/被保險人是相同，但卻可以選擇本人以外的選項；反之亦同)
            If txtIins_IINS2.Text = txtIins_ISSU2.Text Then
                If Relationship.SelectedValue <> "1" Then
                    Throw New Exception("要保人與被保人相同時，與要保人關係請選擇 1.本人")
                End If
            End If
            '為本人時候的檢查
            Dim strErr As String = ""
            Me.IsInputErr = False

            If Relationship.SelectedValue = "1" Then

                If txtIins_IINS2.Text <> txtIins_ISSU2.Text Then
                    Me.IsInputErr = True
                    strErr &= "與要保人關係選擇[本人]時，要保人與被保人[身分證字號]要相同. " & Environment.NewLine

                End If
                If NAME_IINS.Text <> NAME_IISSU.Text Then
                    Me.IsInputErr = True
                    strErr &= "與要保人關係選擇[本人]時，要保人與被保人[姓名]要相同. " & Environment.NewLine
                End If
                'Birtdy_IINS
                If Birtdy_IINS.DateText <> UcRocDbirth.DateText Then
                    Me.IsInputErr = True
                    strErr &= "與要保人關係選擇[本人]時，要保人與被保人[生日]要相同. " & Environment.NewLine
                End If

            End If
            If Me.IsInputErr = True Then
                Throw New Exception(strErr)
            End If

            If TargetProject = "ITP" AndAlso Me.txtCase.Text = "" Then
                Throw New Exception("請先選擇[投保方案]")
            End If

            Dim ChkStr As String = String.Empty

            If txtISEQ.Text = "" Then
                Throw New Exception("[修改被保人]找不到修改資料")
            Else
                '--------------UPDATE 也有可能是新增-------------------------

                '出生日期要小於輸入日期
                If Me.UcRocDbirth.GetDate() > Today() Then
                    Throw New Exception("出生日期 不得小於 輸入日期")
                End If

                '新增一筆 新的
                '先檢查 是不是有重複IISSU
                Dim UpdatCode As String = "0"
                Dim IISSU_0 As String = txtIins_ISSU2.Text.Trim
                'For Each drD0 As DataRow In StateMain.Tables("IQPISSU").Select("IISSU = '" & IISSU_0 & "'")
                '    '有重複IISSU 是修改
                '    UpdatCode = "1"
                'Next
#Region "收納一下"
                '20260116 再次加入檢核〔 [投保方案]需要與主被保人相同〕 '20250320 [投保方案]需要與主被保人相同 此段已經無意義 加入期間又更改．．．　在保費試算在重新更正
                If StateMain.Tables("IQPISSU").Rows.Count > 0 Then
                    '先檢查更改對象是否為主被保人
                    '※※※ 若「資料複製」後修改被保人。「資料複製」時被保人iPlanNo設空值、且要保號預設ITP(IqpMain.IINSKIND = "ITP")
                    Dim ICASE As String = Me.txtCase.Text.Substring(0, 1)
                    Dim ICASE2 As String = ""
                    Dim MainInsured As String = ""
                    'Dim dsChk As beITP101000 = CType(Me.StateMain, beITP101000)

                    If txtCase.Text = "僅人身保險" Then ICASE = ""

                    For Each drCASE2 As DataRow In StateMain.Tables("IQPISSU").Select("ISEQ = '1' ")
                        MainInsured = drCASE2("IISSU").ToString
                        If drCASE2("iPlanNO").ToString.Length >= 1 Then
                            ICASE2 = drCASE2("iPlanNO").ToString.Substring(0, 1)
                        End If
                    Next

                    If ICASE <> ICASE2 Then
                        If MainInsured = IISSU_0 Then
                            '更改對象為主被保人，其他被保人方案一同變更

                            For Each row As DataRow In StateMain.Tables("IQPISSU").Rows
                                row("IPlanNo") = ICASE
                            Next

                            'F115031101 更改IQPMAIN的IINSKIND，以符合選擇變更ITP、IQP方案後，試算能順利執行
                            Dim ModProject As String = ""
                            If ICASE = "" Then
                                ModProject = "IQP"
                            Else
                                ModProject = "ITP"
                            End If

#Region "因為變更ITP、IQP方案，故再次檢核"
                            '修改時主被保人方案時，不允許和要保號方案不同
                            If String.IsNullOrEmpty(Me.txtIQPMainOid.Text) = False Then
                                If ModProject.Substring(0, 2) <> (Me.txtIQPMainOid.Text).Substring(0, 2) Then
                                    Throw New Exception("[要保號]與[投保方案]需相符！")
                                End If
                            End If

                            If ModProject = "ITP" Then
                                Dim ChkAreaLock As String
                                ChkAreaLock = CheckAreaLock()
                                If ChkAreaLock <> String.Empty Then
                                    Throw New Exception(ChkAreaLock)
                                End If
                            End If

                            If ModProject = "IQP" Then
                                '海途要卡
                                If TRAVEL_A.SelectedValue = "400" And ddlIT61.SelectedValue <> "" Then
                                    Throw New Exception("旅遊地區為[400台灣地區及離島地區時]不得承接海突。")
                                ElseIf TRAVEL_A.SelectedValue = "400" And (ddlIT63.SelectedValue <> "" Or ddlIT64.SelectedValue <> "" Or ddlIT6E.SelectedValue <> "") Then
                                    Throw New Exception("旅遊地區為[400台灣地區及離島地區時]不得承接海突。")
                                End If
                            End If
#End Region

                            StateMain.Tables("IQPMAIN").Rows(0)("IINSKIND") = ModProject
                            TargetProject = ModProject
                            ddlTargetProject.SelectedValue = ModProject

                            'dsChk.Tables("IQPISSU").AcceptChanges()
                            'dsChk.Tables("IQPMAIN").AcceptChanges()
                        Else
                            '檢查被保人方案要與主被保人一樣
                            Throw New Exception("[投保方案]需要與主被保人相同")
                        End If
                    End If
                End If



                'ITP 年齡檢核
                If TargetProject = "ITP" AndAlso Me.txtCase.Text = "" Then
                    Throw New Exception("請先選擇[投保方案]")
                ElseIf TargetProject = "ITP" AndAlso Me.txtCase.Text <> "" Then
                    CaseCode = (Me.txtCase.Text).Substring(0, 1)
                End If

                '檢核生日 要跟下拉選單的IT41 IT4Y相同
                If age < 15 Then
                    If ddlIT41.SelectedValue <> "" Then
                        ddlIT41.SelectedValue = ""
                        txtIT41.Text = "0"
                        Throw New Exception(" 未滿15足歲，請選擇失能保險金")

                    End If
                Else
                    If ddlIT4Y.SelectedValue <> "" Then
                        ddlIT4Y.SelectedValue = ""
                        txtIT4Y.Text = "0"
                        Throw New Exception(" 滿15足歲，請選擇意外身故失能保險金")
                    End If
                End If

                Dim sSolution, IADD1, IADD2, IADD3, IADD4, IADD5, IADD6 As String

                If ddlIT41.SelectedValue <> "" Then
                    sSolution = ddlIT41.SelectedValue.Trim + "0000"
                ElseIf ddlIT4Y.SelectedValue <> "" And ddlIT4Y.SelectedValue <> "0" Then
                    sSolution = CDec(ddlIT4Y.SelectedValue.Trim) * 10000
                Else
                    If age < 15 Then
                        If ddlIT4Y.SelectedValue = "" Or ddlIT4Y.SelectedValue = "0" Then
                            If ddl_IISSU.SelectedValue = "5" And TRAVEL_A.SelectedValue = "400" Then
                                Throw New Exception("國內地區喪葬費用保額不可為0.")
                            Else
                                If (checkSpecialChannel().Item1 = True Or checkChannelO().Item1 = True) AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L") Then

                                    If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or ddlIT64.SelectedValue = "") Then
                                        Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                    End If

                                ElseIf checkChannelO.Item1 = True AndAlso (CaseCode = "O") Then

                                    If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or ((ddlIT63.SelectedValue = "" And ddlIT61.SelectedValue = "" And ddlIT64.SelectedValue = ""))) Then
                                        Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                    End If

                                ElseIf (CaseCode <> "J" And CaseCode <> "K" And CaseCode <> "L") AndAlso checkChannelMN.Item1 Then

                                    If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or ((ddlIT63.SelectedValue = "" And ddlIT61.SelectedValue = "" And ddlIT64.SelectedValue = ""))) Then
                                        Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                    End If

                                ElseIf (CaseCode <> "J" And CaseCode <> "K" And CaseCode <> "L") Then

                                    If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or (ddlIT61.SelectedValue = "" And (ddlIT63.SelectedValue = "" And ddlIT6E.SelectedValue = ""))) Then
                                        Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                    End If

                                End If
                            End If
                        End If
                    Else
                        Throw New Exception("主保險契約:[旅行平安保險]請選擇[保險金額]")
                    End If
                End If

                If ddlIT42.SelectedValue <> "" Then
                    IADD1 = ddlIT42.SelectedValue.Trim + "0000"
                Else
                    IADD1 = 0
                End If
                If ddlIT61.SelectedValue <> "" Then ' And rad_Disease_N.Checked = True Then
                    IADD2 = ddlIT61.SelectedValue.Trim + "0000"
                Else
                    IADD2 = 0
                End If
                If ddlIT63.SelectedValue <> "" Then ' And rad_Disease_Y.Checked = True Then
                    IADD4 = ddlIT63.SelectedValue.Trim + "0000"
                Else
                    IADD4 = 0
                End If
                If ddlIT6E.SelectedValue <> "" Then
                    IADD4 = ddlIT6E.SelectedValue.Trim + "0000"
                ElseIf ddlIT63.SelectedValue = "" Then
                    IADD4 = 0
                End If
                If ddlIT44.SelectedValue <> "" Then
                    IADD3 = ddlIT44.SelectedValue.Trim + "0000"
                Else
                    IADD3 = 0
                End If
                If ddlIT43.SelectedValue <> "" Then
                    IADD5 = ddlIT43.SelectedValue.Trim + "0000"
                Else
                    IADD5 = 0
                End If
                If ddlIT64.SelectedValue <> "" Then
                    IADD6 = ddlIT64.SelectedValue.Trim + "0000"
                Else
                    IADD6 = 0
                End If

                If (checkSpecialChannel().Item1 = True AndAlso TargetProject = "ITP" AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L")) Or
                         (checkChannelO().Item1 = True AndAlso TargetProject = "ITP" AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L" Or CaseCode = "O")) Then
                    ChkStr = CheckAge_ITP_SPEChannel(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate(), sSolution, IADD1, IADD2, "", ddl_IISSU.SelectedValue.ToString, Me.txtCase.Text)
                Else
                    ChkStr = CheckAge_ITP(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate(), sSolution, IADD1, IADD2, CaseCode, ddl_IISSU.SelectedValue.ToString) 'Me.txtCase.Text.Substring(0, 1)
                End If

                If ChkStr <> String.Empty Then
                    Throw New Exception(ChkStr)
                End If

                Dim ChkStr2 As String = String.Empty
                If (checkSpecialChannel().Item1 = True AndAlso TargetProject = "ITP" AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L")) Or
                        (checkChannelO().Item1 = True AndAlso TargetProject = "ITP" AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L" Or CaseCode = "O")) Then
                    ChkStr2 = CheckiPlanNo_SPEChannel(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate(), txtCase.Text, sSolution, IADD1, IADD2, IADD3, IADD6, ddl_IISSU.SelectedValue.ToString, IADD5) 'Me.txtCase.Text.Substring(0, 1)
                Else
                    'ChkStr2 = CheckiPlanNo(Me.txtCase.Text.Substring(0, 1), sSolution, IADD1, IADD2, IADD3)
                    ChkStr2 = CheckiPlanNo(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate(), "", sSolution, IADD1, IADD2, IADD3, If(checkChannelMN().Item1 AndAlso TargetProject = "ITP", IADD6, IADD4), ddl_IISSU.SelectedValue.ToString, IADD5) 'Me.txtCase.Text.Substring(0, 1)
                End If

                If ChkStr2 <> String.Empty Then
                    Throw New Exception(ChkStr2)
                End If


                'F115060500 因投保項目而調整保險期間的天數
                ChkStr2 = String.Empty
                ChkStr2 = CheckPlanNTime(sSolution, IADD1, IADD2, IADD4, IADD6, IADD4)
                If ChkStr2 <> String.Empty Then
                    Throw New Exception(ChkStr2)
                End If

                '==公會檢核
#End Region

#Region "F115031101 修改被保人時可改身分證，而不像原先會新增一個被保人，IISSU一律先刪除再新增"

                If ddl_IISSU.SelectedValue = "5" Then
                    Throw New Exception("主被保險人必須為本國人或擁有居留證之外國人")
                End If

                'F115031101 修改被保人時可改身分證，而不像原先會新增一個被保人，IISSU一律先刪除再新增

                '檢核被保人ID是否重複，ISEQ不同、ID相同
                For Each drD1 As DataRow In StateMain.Tables("IQPISSU").Select("IISSU = '" & IISSU_0 & "' and ISEQ <> '" & txtISEQ.Text & "'")
                    Throw New Exception("被保人ID重複")
                Next

                '先刪除dgdMain
                Dim IISSU_Ori As String = txtIISSU_Ori.Text.Trim
                For Each drD1 As DataRow In StateMain.Tables("IQPISSU").Select("IISSU = '" & IISSU_Ori & "'")
                    drD1.Delete()
                Next

                '再刪除dgdMain2  
                For Each drD2 As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & IISSU_Ori & "'")
                    drD2.Delete()
                Next

                '立即針對Table做刪除，避免後續查找StateMain時發生錯誤
                StateMain.Tables("IQPISSU").AcceptChanges()
                StateMain.Tables("IQPOBJ").AcceptChanges()

                BaseSetChildFormData(dgdMain2, StateMain)

                If txtISEQ.Text = "1" Then
                    txtMIISSU.Text = txtIins_ISSU2.Text.Trim
                    txtMNIISSU.Text = NAME_IISSU.Text.Trim
                End If

                '再新增
                Dim ds As beITP101000 = CType(Me.StateMain, beITP101000)
                Dim dr As beITP101000.IQPISSURow
                dr = ds.IQPISSU.NewIQPISSURow
                If Relationship.SelectedValue = "0" Then dr("JOBOFFICE") = txtRelationshipOther.Text.Trim 'F114010802 旅遊險b2b被保人與要保人關係#9078
                'F115031101 修改被保人時，可改身分證而不是新增一個被保人，IISSU一律先刪除再新增
                dr("ISEQ") = txtISEQ.Text.Trim

                'If UpdatCode = "1" Then
                '    'SEQ不變
                '    dr("ISEQ") = txtISEQ.Text.Trim
                'Else
                '    Dim iindex As Integer = ds.IQPISSU.DefaultView.Count
                '    dr("ISEQ") = iindex + 1
                'End If

#End Region

                dr("NISSU") = NAME_IISSU.Text.Trim
                dr("IISSU") = txtIins_ISSU2.Text.Trim
                dr("NISSUBD") = UcRocDbirth.DateText.Trim

                dr("IISSU_TP") = ddl_IISSU.SelectedValue.Trim
                dr("IISSU_ZIP") = IISSU_AREA.Text.Trim
                dr("AISSU") = IISSU_ADDR.Text.Trim
                dr("NISSU_TEL") = IISSU_TEL.Text.Trim
                dr("NISSUMAIL") = EMAIL.Text.Trim

                dr("NAMEREL") = Benefi.Text.Trim
                dr("NINSREL") = Relationship.SelectedValue.Trim
                dr("IDREL") = Benefi_ID.Text.Trim
                'Relationship2
                dr("TYPEREL") = Relationship2.SelectedValue.Trim

                dr("NAMEREL2") = Benefi2.Text.Trim
                dr("IDREL2") = Benefi_ID2.Text.Trim
                dr("TYPEREL2") = Relationship22.SelectedValue.Trim

                dr("NAMEREL3") = Benefi3.Text.Trim
                dr("IDREL3") = Benefi_ID3.Text.Trim
                dr("TYPEREL3") = Relationship23.SelectedValue.Trim

                dr("RELMOTHED") = Me.ddlRELMOTHED.SelectedValue.Trim

                dr("iPlanNo") = If(TargetProject = "ITP", Me.txtCase.Text.Substring(0, 1), "") '20250320 決定方案項目

                dr("TELREL") = txtTELREL.Text.Trim
                dr("ZIPREL") = txtZIPREL.Text.Trim
                dr("AREL") = txtAREL.Text.Trim
                dr("IISSU_SEX") = ddlSex.SelectedValue.Trim

                dr("JOBCONTENT") = txtJOB.Text

                'F114081901 要求海外需填寫英文姓名 & 護照號碼。原本只要求勾選申請﹝申根文件﹞需填寫，但81歲亦有限制
                If TRAVEL_A.SelectedValue <> "400" And age <= 81 Then
                    dr("PASSPORT_NAME") = PASSPORT_NAME.Text.Trim
                    dr("PASSPORT_NUMBER") = PASSPORT_NUMBER.Text.Trim
                End If

                '是否懷孕 20221227
                If isEMP_Y.Checked = True Then
                    dr("EMP") = "Y"
                ElseIf isEMP_N.Checked = True Then
                    dr("EMP") = "N"
                End If

                If ddl_IISSU.SelectedValue = "1" Then
                    dr("IISSU_COUNTRY") = "TW"
                ElseIf ddl_IISSU.SelectedValue = "5" Then
                    dr("IISSU_COUNTRY") = ddlIISSU_COUNTRY.SelectedValue.ToString
                Else
                    dr("IISSU_COUNTRY") = "XG"
                End If

                'dr("IISSU_COUNTRY") = IISSU_NATION.SelectedValue

                Dim a, b, c, d As Integer

                If ddlIT41.SelectedValue <> "" Then
                    dr("ISSU_1") = ddlIT41.SelectedValue.Trim + "0000"
                    a = CInt(txtIT41.Text)
                ElseIf ddlIT4Y.SelectedValue <> "" Then
                    dr("ISSU_1") = CInt(CDec(ddlIT4Y.SelectedValue.Trim) * 10000)
                    a = CInt(txtIT4Y.Text)
                End If

                If ddlIT42.SelectedValue <> "" Then
                    dr("ISSU_2") = ddlIT42.SelectedValue.Trim + "0000"
                End If
                If ddlIT61.SelectedValue <> "" Then
                    dr("ISSU_3") = ddlIT61.SelectedValue.Trim + "0000"
                End If
                If ddlIT63.SelectedValue <> "" Then
                    dr("ISSU_3D") = ddlIT63.SelectedValue.Trim + "0000"
                End If
                If ddlIT6E.SelectedValue <> "" Then
                    dr("ISSU_3D") = ddlIT6E.SelectedValue.Trim + "0000"
                End If
                If ddlIT64.SelectedValue <> "" Then
                    dr("ISSU_IT64") = ddlIT64.SelectedValue.Trim + "0000"
                End If
                If ddlIT44.SelectedValue <> "" Then
                    dr("ISSU_IT44") = ddlIT44.SelectedValue.Trim + "0000"
                End If
                If ddlIT43.SelectedValue <> "" Then
                    dr("ISSU_IT43") = ddlIT43.SelectedValue.Trim + "0000"
                End If

                If txtIT42.Text <> "" Then
                    b = CInt(txtIT42.Text)
                End If
                If txtIT44.Text <> "" Then
                    d = CInt(txtIT44.Text)
                End If
                If txtIT61.Text <> "" Then
                    c = CInt(txtIT61.Text)
                End If

                '===================================================================
                '新增InsertIQPOBJ
                InsertIQPOBJ(dr("iPlanNo"), "")

                '重算總保費
                Dim TOTAMT As Integer = 0
                Dim TOTPREM As Integer = 0
                For Each drT As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & dr("IISSU") & "'")
                    'drD.Delete()
                    TOTPREM += CInt(drT("INSPREM"))
                Next
                dr("ISSU_4") = TOTPREM
                'StateMain.AcceptChanges()

                ds.IQPISSU.AddIQPISSURow(dr)
                BaseSetChildFormData(Me.dgdMain, Me.StateMain)
                '隱藏保費資訊
                ShowPremInfoOrNot(False)

                '清空投保方案
                'Me.txtCase.Text = ""
                txtISEQ.Text = ""
                txtIISSU_Ori.Text = ""
                '新增修改完之後 清除資料
                IISSU_Clear()
                '新增修改完鎖住 重要要保資料
                'LockPage()
                '加上鎖住方案
                'CK2.Checked = True
                'CK2.Enabled = False
            End If

        Catch ex As Exception
            StateMain.Tables("IQPMAIN").RejectChanges()
            StateMain.Tables("IQPISSU").RejectChanges()
            StateMain.Tables("IQPOBJ").RejectChanges()
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Public Function CheckAge_ITP(ByVal DplyBgn As Date, ByVal sDBIRTH As Date, ByVal sISOLUTION As String, ByVal IADDSOL1 As String, ByVal IADDSOL2 As String, ByVal iPlanNo As String, ByVal IISSU_TP As String) As String

        Dim err As String = ""

        Dim AgeNow As Decimal = B2BEntity.HasFunction.GetAgeInLaw(DplyBgn, sDBIRTH)
        Dim AMT11 As Decimal = GetDecimal(sISOLUTION) '死亡及失能
        Dim AMT21 As Decimal = GetDecimal(IADDSOL1) '實支實付型 
        Dim AMT31 As Decimal = GetDecimal(IADDSOL2) '海外

        If AgeNow < 15 Then
            If HasFunction.IsUnder15YearsOldInLaw(sDBIRTH, DplyBgn) Then
                If AMT11 > 690000 Then
                    err &= "未滿15足歲，身故失能保額建議最高69萬;" & Environment.NewLine
                End If
            End If
        ElseIf AgeNow >= 15 And AgeNow < 71 Then
            If TRAVEL_A.SelectedValue = "400" Then
                '國內      
                If IISSU_TP = "5" Then

                Else
                    If AMT11 > 5000000 Then
                        err &= "15~70歲，國內 身故失能保額建議最高500萬;" & Environment.NewLine
                    End If
                End If
            Else
                If AMT11 > 15000000 Then
                    err &= "15~70歲，國外 身故失能保額建議最高1500萬;" & Environment.NewLine
                End If

                If iPlanNo.ToUpper = "C" AndAlso Me.txtCompanyid.Text.Trim <> "K077" Then ' 海外基本型 
                    'F115031200 新增卡控至少100萬
                    If AMT11 < 1000000 Then
                        err &= "15~70歲，海外基本型 身故失能保額建議至少100萬;" & Environment.NewLine
                    End If
                End If

                If iPlanNo.ToUpper = "D" AndAlso Me.txtCompanyid.Text.Trim <> "K077" Then ' 海外超值型 
                    'F115031200 新增卡控至少500萬
                    If AMT11 < 5000000 Then
                        err &= "15~70歲，海外超值型 身故失能保額建議至少500萬;" & Environment.NewLine
                    End If
                End If

            End If
        ElseIf AgeNow > 80 Then
            If checkChannelMN().Item1 AndAlso TRAVEL_A.SelectedValue = "400" Then
                'F115060500 K077 國內 81~100 身故保額限制50萬
                If AgeNow <= 100 Then
                    If AMT11 > 500000 Then
                        err &= "81~100歲，身故失能保額建議最高50萬;"
                    End If
                Else
                    err &= "超過100歲建議不投保此保險;"
                End If
            ElseIf checkChannelMN().Item1 AndAlso TRAVEL_A.SelectedValue <> "400" AndAlso Checkbox2.Checked = False Then
                'F115060500 K077 國外非申根 71~85 身故保額限制300萬
                If AgeNow <= 85 Then
                    If AMT11 > 3000000 Then
                        err &= "71~85歲，身故失能保額建議最高300萬;"
                    End If
                Else
                    err &= "超過85歲建議不投保此保險;"
                End If
            Else
                err &= "超過80歲建議不投保此保險;"
            End If
        ElseIf AgeNow >= 76 Then
            If Checkbox2.Checked = True Then '申根方案
                err &= "[申根方案]超過70歲建議不投保此保險;"
            Else
                If checkChannelMN().Item1 AndAlso TRAVEL_A.SelectedValue <> "400" Then
                    If AMT11 > 3000000 Then
                        err &= "71~85歲，身故失能保額建議最高300萬;"
                    End If
                Else
                    If AMT11 > 1000000 Then
                        err &= "76~80歲，身故失能保額建議最高100萬;" & Environment.NewLine
                    End If
                End If
            End If
        ElseIf AgeNow >= 71 Then
            If Checkbox2.Checked = True Then '申根方案
                err &= "[申根方案]超過70歲建議不投保此保險;"
            Else
                If checkChannelMN().Item1 AndAlso TRAVEL_A.SelectedValue <> "400" Then
                    If AMT11 > 3000000 Then
                        err &= "71~85歲，身故失能保額建議最高300萬;"
                    End If
                Else
                    If AMT11 > 2000000 Then
                        err &= "71~75歲，身故失能保額建議最高200萬;" & Environment.NewLine
                    End If
                End If
            End If
        Else
            err &= "[年齡計算錯誤];" & Environment.NewLine
        End If

        'If AgeNow >= 15 Then

        '	If TRAVEL_A.SelectedValue = "400" Then
        '		'國內           
        '		If AMT11 > 5000000 Then
        '			err &= "15~70歲，國內 身故失能保額建議最高500萬;" & Environment.NewLine
        '		End If
        '	Else
        '		If AMT11 > 15000000 Then
        '			err &= "15~70歲，國外 身故失能保額建議最高1500萬;" & Environment.NewLine
        '		End If
        '	End If
        'End If

        Return err

    End Function
    Public Function CheckAge_ITP_SPEChannel(ByVal DplyBgn As Date, ByVal sDBIRTH As Date, ByVal sISOLUTION As String, ByVal IADDSOL1 As String, ByVal IADDSOL2 As String, ByVal iPlanNo As String, ByVal IISSU_TP As String, ByVal xIPLAN As String) As String

        Dim err As String = ""

        Dim AgeNow As Decimal = B2BEntity.HasFunction.GetAgeInLaw(DplyBgn, sDBIRTH)
        Dim AMT11 As Decimal = GetDecimal(sISOLUTION) '死亡及失能
        Dim AMT21 As Decimal = GetDecimal(IADDSOL1) '實支實付型 
        Dim AMT31 As Decimal = GetDecimal(IADDSOL2) '海外
        If Not String.IsNullOrEmpty(xIPLAN) Then xIPLAN = xIPLAN.Substring(0, 1)

        If AgeNow < 15 Then
            If HasFunction.IsUnder15YearsOldInLaw(sDBIRTH, DplyBgn) Then
                If AMT11 > 690000 Then
                    err &= "未滿15足歲，身故失能保額建議最高69萬;" & Environment.NewLine
                End If
            End If
        ElseIf AgeNow >= 15 And AgeNow < 71 Then
            If TRAVEL_A.SelectedValue = "400" Then
                '國內      
                If IISSU_TP = "5" Then

                Else
                    If AMT11 > 5000000 Then
                        err &= "15~70歲，國內 身故失能保額建議最高500萬;" & Environment.NewLine
                    End If
                End If
            ElseIf Checkbox2.Checked Then '申根方案
                If AMT11 < 5000000 Then
                    err &= "[申根方案]15~70歲，國外 身故失能保額建議至少500萬;" & Environment.NewLine
                End If
                If AMT11 > 15000000 Then
                    err &= "[申根方案]15~70歲，國外 身故失能保額建議最高1500萬;" & Environment.NewLine
                End If
            Else
                If xIPLAN = "K" Then
                    If AMT11 < 5000000 Then
                        err &= "[海外豪華型]15~70歲，國外 身故失能保額建議至少500萬;" & Environment.NewLine
                    End If
                ElseIf xIPLAN = "O" Then
                    If AMT11 < 1000000 Then
                        err &= "15~70歲，國外 身故失能保額建議至少100萬;" & Environment.NewLine
                    End If
                Else
                    If AMT11 < 2000000 Then
                        err &= "15~70歲，國外 身故失能保額建議至少200萬;" & Environment.NewLine
                    End If
                End If

                If AMT11 > 15000000 Then
                    err &= "15~70歲，國外 身故失能保額建議最高1500萬;" & Environment.NewLine
                End If
            End If
        ElseIf AgeNow > 80 Then
            err &= "超過80歲建議不投保此保險;"
        ElseIf AgeNow >= 76 Then
            If TRAVEL_A.SelectedValue = "400" Then
                If AMT11 > 1000000 Then
                    err &= "76~80歲，國內 身故失能保額建議最高100萬;" & Environment.NewLine
                End If
            ElseIf Checkbox2.Checked = True AndAlso xIPLAN <> "O" Then '申根方案
                err &= "[申根方案]超過75歲建議不投保此保險;" & Environment.NewLine
            ElseIf Checkbox2.Checked = True AndAlso xIPLAN = "O" Then '申根方案 
                If AMT11 < 1000000 Then
                    err &= "76~80歲，國外 身故失能保額建議至少100萬;" & Environment.NewLine
                End If
                If AMT11 > 3000000 Then
                    err &= "[申根方案]76~80歲，身故失能保額建議最高300萬;" & Environment.NewLine
                End If
            Else
                If xIPLAN = "K" Then
                    err &= "超過75歲建議不投保此保險;" & Environment.NewLine
                ElseIf xIPLAN = "O" Then
                    If AMT11 < 1000000 Then
                        err &= "76~80歲，國外 身故失能保額建議至少100萬;" & Environment.NewLine
                    End If
                    If AMT11 > 3000000 Then
                        err &= "76~80歲，國外 身故失能保額建議最高300萬;" & Environment.NewLine
                    End If
                Else
                    If AMT11 < 2000000 Then
                        err &= "76~80歲，國外 身故失能保額建議至少200萬;" & Environment.NewLine
                    End If
                    If AMT11 > 3000000 Then
                        err &= "76~80歲，國外 身故失能保額建議最高300萬;" & Environment.NewLine
                    End If
                End If
            End If
        ElseIf AgeNow >= 71 Then
            If TRAVEL_A.SelectedValue = "400" Then
                If AMT11 > 2000000 Then
                    err &= "71~75歲，身故失能保額建議最高200萬;" & Environment.NewLine
                End If
            ElseIf Checkbox2.Checked = True Then '申根方案
                If AMT11 <> 5000000 Then
                    err &= "[申根方案]71~75歲，身故失能保額建議選擇500萬;" & Environment.NewLine
                End If
            Else
                If xIPLAN = "K" Then
                    If AMT11 <> 5000000 Then
                        err &= "[海外豪華型]71~75歲，國外 身故失能保額建議選擇500萬;" & Environment.NewLine
                    End If
                ElseIf xIPLAN = "O" Then
                    If AMT11 < 1000000 Then
                        err &= "71~75歲，國外 身故失能保額建議至少100萬;" & Environment.NewLine
                    End If
                    If AMT11 > 5000000 Then
                        err &= "71~75歲，國外 身故失能保額建議最高500萬;" & Environment.NewLine
                    End If
                Else
                    If AMT11 < 2000000 Then
                        err &= "71~75歲，國外 身故失能保額建議至少200萬;" & Environment.NewLine
                    End If
                    If AMT11 > 5000000 Then
                        err &= "71~75歲，國外 身故失能保額建議最高500萬;" & Environment.NewLine
                    End If
                End If
            End If
        Else
            err &= "[年齡計算錯誤];" & Environment.NewLine
        End If

        'If AgeNow >= 15 Then

        '	If TRAVEL_A.SelectedValue = "400" Then
        '		'國內           
        '		If AMT11 > 5000000 Then
        '			err &= "15~70歲，國內 身故失能保額建議最高500萬;" & Environment.NewLine
        '		End If
        '	Else
        '		If AMT11 > 15000000 Then
        '			err &= "15~70歲，國外 身故失能保額建議最高1500萬;" & Environment.NewLine
        '		End If
        '	End If
        'End If

        Return err

    End Function

    Public Function CheckiPlanNo(ByVal DplyBgn As Date, ByVal sDBIRTH As Date, ByVal iPlanNo As String, ByVal sISOLUTION As String, ByVal IADDSOL1 As String, ByVal IADDSOL2 As String, ByVal IT44 As String, ByVal IADDSOL3 As String, ByVal IISSU_TP As String, ByVal IT43 As String) As String

        Dim err As String = ""

        Dim AgeNow As Decimal = B2BEntity.HasFunction.GetAgeInLaw(DplyBgn, sDBIRTH)
        Dim AMT11 As Decimal = GetDecimal(sISOLUTION) '死亡及失能(意外身故)
        Dim AMT21 As Decimal = GetDecimal(IADDSOL1) '實支實付型 
        Dim AMT31 As Decimal = GetDecimal(IADDSOL2) '海外(不含傳染病)
        Dim AMT32 As Decimal = GetDecimal(IADDSOL3) '海外(含傳染病)
        Dim AMT44 As Decimal = GetDecimal(IT44) '重大傷燙
        Dim AMT43 As Decimal = GetDecimal(IT43) '大眾運輸
        Dim check1 As Decimal
        Dim check2 As Decimal
        '先檢核有沒有超過死亡及失能1/10
        check1 = AMT11 * 0.1

#Region "檢查旅遊地區是否可投保該方案(ITP可選地區較少)"

        If ddlTargetProject.SelectedValue = "ITP" Then

            Dim CounCode As String = "", MsgCoun As String = ""
            Dim CounA As String = ddlCountry_A2.SelectedValue.ToString.Trim
            Dim CounB As String = ddlCountry_B2.SelectedValue.ToString.Trim
            Dim CounC As String = ddlCountry_C2.SelectedValue.ToString.Trim
            Dim CounD As String = ddlCountry_D2.SelectedValue.ToString.Trim
            Dim CounE As String = ddlCountry_E2.SelectedValue.ToString.Trim
            Dim CounF As String = ddlCountry_F2.SelectedValue.ToString.Trim
            Dim CounG As String = ddlCountry_G2.SelectedValue.ToString.Trim
            Dim CounH As String = ddlCountry_H2.SelectedValue.ToString.Trim

            CounCode = "'" + CounA.Substring(0, 2) + "'" +
                        If(String.IsNullOrEmpty(CounB), "", (", '" + CounB.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounC), "", (", '" + CounC.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounD), "", (", '" + CounD.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounE), "", (", '" + CounE.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounF), "", (", '" + CounF.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounG), "", (", '" + CounG.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounH), "", (", '" + CounH.Substring(0, 2) + "'"))

            Dim dsTemp As DataSet = Me.InvokeCommonService("PUBService.bsSingleSQL", New String() {"skerp_b2b.ITP_COUNTRY", "COUNCODE in (" + CounCode + ") and PN = 'S'", "COUNTRYCHA"})
            If dsTemp.Tables(0).DefaultView.Count > 0 Then
                For Each dr As DataRow In dsTemp.Tables(0).Rows
                    MsgCoun = MsgCoun + If(String.IsNullOrEmpty(MsgCoun), dr("COUNTRYCHA"), ("、" + dr("COUNTRYCHA")))
                Next
                err &= "旅遊國家 " + MsgCoun + "，不可選擇旅綜方案；" & Environment.NewLine
            End If
        End If

#End Region

        If AMT43 <> 0 Then
            err &= "不建議保搭乘大眾運輸意外保險金;" & Environment.NewLine

        End If

        If AMT31 > 0 And AMT32 > 0 Then
            err &= "無法同時投保海外突發疾病保險金(不含傳染病)以及海外突發疾病保險金(含傳染病)，請擇一;" & Environment.NewLine
        End If

        If AgeNow < 15 Then
            If AMT11 > 690000 Then
                err &= "15歲以下" & If(IISSU_TP = "5", "、外國人持護照投保", "") & " 意外喪葬費用或意外失能保險金，建議保額最高 69萬;" & Environment.NewLine
            End If

            If IISSU_TP = "5" Then
                If AMT11 = 0 Then
                    err &= "國內地區喪葬費用保額不可為0;" & Environment.NewLine
                End If
            End If

            If AMT44 > 500000 Then
                err &= "15歲以下" & If(IISSU_TP = "5", "、外國人持護照投保", "") & " 重大燒燙傷保險金，建議保額最高 50萬;" & Environment.NewLine
            End If

            If Checkbox2.Checked = True Then '申根方案
                If AMT21 <> 1500000 Then
                    err &= "選擇[申根地區]時 傷害醫療保險金 建議保額 150萬;" & Environment.NewLine
                Else
                    If AMT32 <> 0 And AMT32 <> 1500000 Then
                        err &= "選擇[申根地區]時 傷害醫療保險金 建議保額 150萬; 海外突發疾病醫療保險金(含傳染病)，建議投保 150萬;" & Environment.NewLine
                    End If

                    If AMT31 <> 0 And AMT31 <> 1000000 Then
                        err &= "選擇[申根地區]時 傷害醫療保險金 建議保額 150萬; 海外突發疾病醫療保險金(不含傳染病)，建議投保 100萬;" & Environment.NewLine
                    End If
                End If
            Else
                If IISSU_TP = "5" Then
                    If AMT21 > 50000 Then
                        err &= "15歲以下、外國人持護照投保 傷害醫療保險金，建議保額最高 5萬;" & Environment.NewLine
                    End If
                Else
                    If AMT21 > 200000 Then
                        err &= "15歲以下 傷害醫療保險金，建議保額最高 20萬;" & Environment.NewLine
                    End If
                End If

                If AMT31 > 200000 Then
                    err &= "15歲以下 海外突發疾病醫療保險金，建議保額最高 20萬;" & Environment.NewLine
                End If
                If AMT32 > 200000 Then
                    err &= "15歲以下 海外突發疾病醫療保險金，建議保額最高 20萬;" & Environment.NewLine
                End If
            End If
        Else
            If Checkbox2.Checked = True Then '申根方案
                If AgeNow <= 70 Then
                    If AMT11 < 5000000 Then
                        err &= "15~70歲，選擇[申根地區]時 身故失能保額建議保額最低 500萬;" & Environment.NewLine
                    End If
                End If

                If AMT21 = 1500000 Then
                    If AgeNow > 70 Then
                        err &= "[申根地區]超過70歲建議不投保此保險;"
                    End If
                Else
                    err &= "選擇[申根地區]時 傷害醫療保險金 建議保額 150萬; " & Environment.NewLine
                End If

                If AMT32 <> 0 Then
                    If AMT32 <> 1500000 Then
                        err &= "海外突發疾病醫療保險金(含傳染病)，建議投保 150萬;" & Environment.NewLine
                    End If
                End If

                If AMT31 <> 0 Then
                    If AMT31 <> 1000000 Then
                        err &= "海外突發疾病醫療保險金(含傳染病)，建議投保 100萬;" & Environment.NewLine
                    End If
                End If
            Else
                If AMT32 > 1000000 Then
                    err &= "海外突發疾病醫療保險金(含傳染病)，建議保額最高 100萬;" & Environment.NewLine
                End If

                If AMT31 > 1000000 Then
                    err &= "海外突發疾病醫療保險金(不含傳染病)，建議保額最高 100萬;" & Environment.NewLine
                End If

                If IISSU_TP = "5" Then
                    If AMT11 > 2000000 Then
                        err &= "外國人持護照投保 意外身故失能保險金，建議保額最高 200萬;" & Environment.NewLine
                    End If

                    If AMT21 > 50000 Then
                        err &= "外國人持護照投保 傷害醫療保險金，建議保額最高 5萬;" & Environment.NewLine
                    End If
                Else
                    If AMT21 > check1 Then
                        err &= "傷害醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                    End If
                End If

                If AMT31 > check1 Then
                    err &= "海外突發疾病醫療保險金(不含傳染病)，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If

                If AMT32 > check1 Then
                    err &= "海外突發疾病醫療保險金(含傳染病)，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If

            End If

            If IISSU_TP = "5" Then
                If AgeNow <= 75 Then
                    If AMT44 > 500000 Then
                        err &= "外國人持護照投保 重大燒燙傷保險金，建議保額最高 50萬;" & Environment.NewLine
                    End If
                End If
            Else
                '檢核有沒有超過死亡及失能1/4
                check2 = AMT11 * 0.25
                If AMT44 > check2 Then
                    err &= "重大燒燙傷保險金，建議不超過身故失能保險金4分之1;" & Environment.NewLine
                End If
            End If

        End If

        If AgeNow > 70 Then
            If AMT32 > 0 Then
                If checkChannelMN().Item1 Then
                    ' F115060500 K077 國外 71~75可保含傳染病
                    If AgeNow > 75 Then
                        err &= "76歲(含)以上 不建議保海外突發疾病醫療保險金(含傳染病);" & Environment.NewLine
                    End If
                Else
                    err &= "70歲以上 不建議保海外突發疾病醫療保險金(含傳染病);" & Environment.NewLine
                End If
            End If

            If AMT31 > 0 Then
                If Checkbox2.Checked = True Then '申根方案
                    err &= "70歲以上 選擇[申根地區]時 不提供投保[海外突發疾病醫療保險金(不含傳染病)];" & Environment.NewLine
                ElseIf AgeNow <= 80 And AMT31 > check1 Then
                    err &= "71~80歲海外突發疾病醫療保險金(不含傳染病)，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                ElseIf AgeNow > 80 Then
                    If checkChannelMN().Item1 Then
                        ' F115060500 K077 國外 85歲以上不可保(不含傳染病)
                        If AgeNow > 85 Then
                            err &= "85歲以上 不建議保海外突發疾病醫療保險金(不含傳染病);" & Environment.NewLine
                        End If
                    Else
                        err &= "80歲以上 不建議保海外突發疾病醫療保險金(不含傳染病);" & Environment.NewLine
                    End If
                End If
            End If

            '外籍人士用護照投保 年齡76~80 身故100萬為上限 醫療上限5萬 重大燒燙傷 上限25萬
            If IISSU_TP = "5" AndAlso AgeNow > 75 Then
                If AMT11 > 1000000 Then
                    err &= "外國人持護照投保 意外身故失能保險金，建議保額最高 100萬;" & Environment.NewLine
                End If

                If AMT44 > 250000 Then
                    err &= "外國人持護照投保 重大燒燙傷保險金，建議保額最高 25萬;" & Environment.NewLine
                End If
            End If
        End If



        Return err

    End Function

    Public Function CheckiPlanNo_SPEChannel(ByVal DplyBgn As Date, ByVal sDBIRTH As Date, ByVal iPlanNo As String, ByVal sISOLUTION As String, ByVal IADDSOL1 As String, ByVal IADDSOL2 As String, ByVal IT44 As String, ByVal IADDSOL3 As String, ByVal IISSU_TP As String, ByVal IT43 As String) As String

        Dim err As String = ""

        Dim AgeNow As Decimal = B2BEntity.HasFunction.GetAgeInLaw(DplyBgn, sDBIRTH)
        Dim AMT11 As Decimal = GetDecimal(sISOLUTION) '死亡及失能(意外身故)
        Dim AMT21 As Decimal = GetDecimal(IADDSOL1) '實支實付型 
        Dim AMT31 As Decimal = GetDecimal(IADDSOL2) '海外(不含傳染病)
        Dim AMT32 As Decimal = GetDecimal(IADDSOL3) '海外(含傳染病)
        Dim AMT44 As Decimal = GetDecimal(IT44) '重大傷燙
        Dim AMT43 As Decimal = GetDecimal(IT43) '大眾運輸
        Dim check1 As Decimal
        Dim check2 As Decimal
        If Not String.IsNullOrEmpty(iPlanNo) Then iPlanNo = iPlanNo.Substring(0, 1)
        '先檢核有沒有超過死亡及失能1/10
        check1 = AMT11 * 0.1

#Region "檢查旅遊地區是否可投保該方案(ITP可選地區較少)"

        If ddlTargetProject.SelectedValue = "ITP" Then

            Dim CounCode As String = "", MsgCoun As String = ""
            Dim CounA As String = ddlCountry_A2.SelectedValue.ToString.Trim
            Dim CounB As String = ddlCountry_B2.SelectedValue.ToString.Trim
            Dim CounC As String = ddlCountry_C2.SelectedValue.ToString.Trim
            Dim CounD As String = ddlCountry_D2.SelectedValue.ToString.Trim
            Dim CounE As String = ddlCountry_E2.SelectedValue.ToString.Trim
            Dim CounF As String = ddlCountry_F2.SelectedValue.ToString.Trim
            Dim CounG As String = ddlCountry_G2.SelectedValue.ToString.Trim
            Dim CounH As String = ddlCountry_H2.SelectedValue.ToString.Trim

            CounCode = "'" + CounA.Substring(0, 2) + "'" +
                        If(String.IsNullOrEmpty(CounB), "", (", '" + CounB.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounC), "", (", '" + CounC.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounD), "", (", '" + CounD.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounE), "", (", '" + CounE.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounF), "", (", '" + CounF.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounG), "", (", '" + CounG.Substring(0, 2) + "'")) +
                        If(String.IsNullOrEmpty(CounH), "", (", '" + CounH.Substring(0, 2) + "'"))

            Dim dsTemp As DataSet = Me.InvokeCommonService("PUBService.bsSingleSQL", New String() {"skerp_b2b.ITP_COUNTRY", "COUNCODE in (" + CounCode + ") and PN = 'S'", "COUNTRYCHA"})
            If dsTemp.Tables(0).DefaultView.Count > 0 Then
                For Each dr As DataRow In dsTemp.Tables(0).Rows
                    MsgCoun = MsgCoun + If(String.IsNullOrEmpty(MsgCoun), dr("COUNTRYCHA"), ("、" + dr("COUNTRYCHA")))
                Next
                err &= "旅遊國家 " + MsgCoun + "，不可選擇旅綜方案；" & Environment.NewLine
            End If
        End If

#End Region


        If iPlanNo = "O" Then
            If AMT43 <> 0 Then
                'O 海外小資型 沒有大眾運輸
                err &= "海外小資型 建議不投保大眾運輸意外保險金;" & Environment.NewLine
            End If


            If AMT31 > 0 And AMT32 > 0 Then
                err &= "無法同時投保海外突發疾病保險金(不含傳染病)以及海外突發疾病保險金(含傳染病)，請擇一;" & Environment.NewLine
            End If
        End If


        If AgeNow < 15 Then
            If AMT11 > 690000 Then
                err &= "未滿15足歲，" & If(IISSU_TP = "5", "、外國人持護照投保", "") & " 意外喪葬費用或意外失能保險金，建議保額最高 69萬;" & Environment.NewLine
            End If

            If IISSU_TP = "5" Then
                If AMT11 = 0 Then
                    err &= "國內地區喪葬費用保額不可為0;" & Environment.NewLine
                End If
            End If

            If AMT43 <> 0 Then
                err &= "未滿15足歲，不建議保搭乘大眾運輸意外保險金;" & Environment.NewLine
            End If

            If Checkbox2.Checked = True Then '申根方案
                If iPlanNo = "O" Then
                    If AMT44 > 500000 Then
                        err &= "[申根地區]未滿15足歲，重大燒燙傷保險金，建議保額最高 50萬;" & Environment.NewLine
                    End If
                Else
                    If AMT44 <> 500000 Then
                        err &= "[申根地區]未滿15足歲，重大燒燙傷保險金，建議投保 50萬;" & Environment.NewLine
                    End If
                End If

                If AMT21 <> 1500000 Then
                    err &= "[申根地區]未滿15足歲，傷害醫療保險金 建議保額 150萬;" & Environment.NewLine
                End If

                If iPlanNo = "O" Then
                    If AMT31 = 0 And AMT32 = 0 Then
                        err &= "[申根地區]未滿15足歲，海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    ElseIf AMT31 <> 0 AndAlso AMT31 <> 1000000 Then
                        err &= "[申根地區]未滿15足歲，海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    ElseIf AMT32 <> 0 AndAlso AMT32 <> 1000000 Then
                        err &= "[申根地區]未滿15足歲，海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    End If
                Else
                    If AMT32 <> 1000000 Then
                        err &= "[申根地區]未滿15足歲，海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    End If
                End If

            ElseIf TRAVEL_A.SelectedValue = "400" Then '國內
                If IISSU_TP = "5" Then
                    If AMT21 > 50000 Then
                        err &= "未滿15足歲，外國人持護照投保 傷害醫療保險金，建議保額最高 5萬;" & Environment.NewLine
                    End If

                    If AMT44 > 500000 Then
                        err &= "未滿15足歲，外國人持護照投保 重大燒燙傷保險金，建議保額最高 50萬;" & Environment.NewLine
                    End If
                Else
                    If AMT44 <> 500000 Then
                        err &= "未滿15足歲，重大燒燙傷保險金，建議投保 50萬;" & Environment.NewLine
                    End If
                    If AMT21 > 200000 Then
                        err &= "未滿15足歲，傷害醫療保險金，建議保額最高 20萬;" & Environment.NewLine
                    End If
                End If
            Else '國外
                If iPlanNo = "O" Then
                    If AMT44 > 500000 Then
                        err &= "未滿15足歲，重大燒燙傷保險金，建議保額最高 50萬;" & Environment.NewLine
                    End If
                Else
                    If AMT44 <> 500000 Then
                        err &= "未滿15足歲，重大燒燙傷保險金，建議投保 50萬;" & Environment.NewLine
                    End If
                End If

                If AMT21 < 200000 Then
                    err &= "未滿15足歲，國外 傷害醫療保險金，保額建議至少 20萬;" & Environment.NewLine
                End If
                If AMT21 > 1000000 Then
                    err &= "未滿15足歲，國外 傷害醫療保險金，建議保額最高 100萬;" & Environment.NewLine
                End If

                If iPlanNo = "O" Then
                    If AMT31 = 0 And AMT32 = 0 Then
                        err &= "未滿15足歲，海外突發疾病住院醫療保險金，保額建議至少 20萬;" & Environment.NewLine

                    ElseIf AMT32 = 0 Then
                        If AMT31 < 200000 Then
                            err &= "未滿15足歲，海外突發疾病住院醫療保險金，保額建議至少 20萬;" & Environment.NewLine
                        End If
                        If AMT31 > 1000000 Then
                            err &= "未滿15足歲，海外突發疾病住院醫療保險金，建議保額最高 100萬;" & Environment.NewLine
                        End If

                    ElseIf AMT31 = 0 Then
                        If AMT32 < 200000 Then
                            err &= "未滿15足歲，海外突發疾病住院醫療保險金，保額建議至少 20萬;" & Environment.NewLine
                        End If
                        If AMT32 > 1000000 Then
                            err &= "未滿15足歲，海外突發疾病住院醫療保險金，建議保額最高 100萬;" & Environment.NewLine
                        End If
                    End If
                Else
                    If AMT32 < 200000 Then
                        err &= "未滿15足歲，海外突發疾病住院醫療保險金，保額建議至少 20萬;" & Environment.NewLine
                    End If
                    If AMT32 > 1000000 Then
                        err &= "未滿15足歲，海外突發疾病住院醫療保險金，建議保額最高 100萬;" & Environment.NewLine
                    End If
                End If
            End If
        ElseIf AgeNow >= 15 And AgeNow < 71 Then
            'O 海外小資型 沒有大眾運輸

            'AgeNow>15
            '檢核是否為死亡及失能的1/5
            check2 = AMT11 * 0.2
            If iPlanNo <> "O" AndAlso (AMT43 <> check2) Then
                err &= "搭乘大眾運輸意外保險金建議投保 " & check2 & "萬;" & Environment.NewLine
            End If

            If TRAVEL_A.SelectedValue = "400" Then
                If IISSU_TP = "5" Then
                    If AMT11 > 2000000 Then
                        err &= "外國人持護照投保 意外身故失能保險金，建議保額最高 200萬;" & Environment.NewLine
                    End If

                    If AMT21 > 50000 Then
                        err &= "外國人持護照投保 傷害醫療保險金，建議保額最高 5萬;" & Environment.NewLine
                    End If

                    If AMT44 > 500000 Then
                        err &= "外國人持護照投保 重大燒燙傷保險金，建議保額最高 50萬;" & Environment.NewLine
                    End If
                Else
                    If iPlanNo = "O" Then
                        '檢核有沒有超過死亡及失能1/4
                        check2 = AMT11 * 0.25
                        If AMT44 > check2 Then
                            err &= "重大燒燙傷保險金，建議不超過身故失能保險金4分之1;" & Environment.NewLine
                        End If
                    Else
                        '檢核是否為死亡及失能的1/4
                        check2 = AMT11 * 0.25
                        If AMT44 <> check2 Then
                            err &= "重大燒燙傷保險金建議投保 " & check2 & Environment.NewLine
                        End If
                    End If

                    If AMT21 > check1 Then
                        err &= "傷害醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                    End If
                End If
            ElseIf Checkbox2.Checked = True Then '申根方案
                If iPlanNo = "O" Then
                    '檢核有沒有超過死亡及失能1/4
                    check2 = AMT11 * 0.25
                    If AMT44 > check2 Then
                        err &= "重大燒燙傷保險金，建議不超過身故失能保險金4分之1;" & Environment.NewLine
                    End If
                Else
                    '檢核是否為死亡及失能的1/4
                    check2 = AMT11 * 0.25
                    If AMT44 <> check2 Then
                        err &= "[申根地區]重大燒燙傷保險金建議投保 " & check2 & Environment.NewLine
                    End If
                End If

                If AMT21 <> 1500000 Then
                    err &= "[申根地區]傷害醫療保險金，建議投保 150萬; " & Environment.NewLine
                End If

                If iPlanNo = "O" Then
                    If AMT31 = 0 And AMT32 = 0 Then
                        err &= "[申根地區]海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    ElseIf AMT31 <> 0 AndAlso AMT31 <> 1000000 Then
                        err &= "[申根地區]海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    ElseIf AMT32 <> 0 AndAlso AMT32 <> 1000000 Then
                        err &= "[申根地區]海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    End If
                Else
                    If AMT32 <> 1000000 Then
                        err &= "[申根地區]海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    End If
                End If
            Else
                If iPlanNo = "O" Then
                    '檢核有沒有超過死亡及失能1/4
                    check2 = AMT11 * 0.25
                    If AMT44 > check2 Then
                        err &= "重大燒燙傷保險金，建議不超過身故失能保險金4分之1;" & Environment.NewLine
                    End If
                Else
                    '檢核是否為死亡及失能的1/4
                    check2 = AMT11 * 0.25
                    If AMT44 <> check2 Then
                        err &= "重大燒燙傷保險金建議投保 " & check2 & Environment.NewLine
                    End If
                End If

                If AMT21 > check1 Then
                    err &= "傷害醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If

                If AMT32 > check1 Then
                    err &= "海外突發疾病住院醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If

                If iPlanNo = "O" AndAlso AMT31 > check1 Then
                    err &= "海外突發疾病住院醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If

            End If
        ElseIf AgeNow > 80 Then
            err &= "超過80歲建議不投保此保險;"

        ElseIf AgeNow >= 76 Then
            If TRAVEL_A.SelectedValue = "400" Then
                'O 海外小資型 沒有大眾運輸
                '檢核是否為死亡及失能的1/5
                check2 = AMT11 * 0.2
                If iPlanNo <> "O" AndAlso (AMT43 <> check2) Then
                    err &= "搭乘大眾運輸意外保險金建議投保 " & check2 & "萬;" & Environment.NewLine
                End If

                '外籍人士用護照投保 年齡76~80 身故100萬為上限 醫療上限5萬 重大燒燙傷 上限25萬
                If IISSU_TP = "5" Then
                    If AMT11 > 1000000 Then
                        err &= "外國人持護照投保 意外身故失能保險金，建議保額最高 100萬;" & Environment.NewLine
                    End If

                    If AMT44 > 250000 Then
                        err &= "外國人持護照投保 重大燒燙傷保險金，建議保額最高 25萬;" & Environment.NewLine
                    End If

                    If AMT21 > 50000 Then
                        err &= "外國人持護照投保 傷害醫療保險金，建議保額最高 5萬;" & Environment.NewLine
                    End If
                Else
                    If AMT11 > 1000000 Then
                        err &= "76~80歲，國內 身故失能保額建議最高100萬;" & Environment.NewLine
                    End If

                    If iPlanNo = "O" Then
                        '檢核有沒有超過死亡及失能1/4
                        check2 = AMT11 * 0.25
                        If AMT44 > check2 Then
                            err &= "重大燒燙傷保險金，建議不超過身故失能保險金4分之1;" & Environment.NewLine
                        End If
                    Else
                        '檢核是否為死亡及失能的1/4
                        check2 = AMT11 * 0.25
                        If AMT44 <> check2 Then
                            err &= "重大燒燙傷保險金建議投保 " & check2 & "萬;" & Environment.NewLine
                        End If
                    End If

                    If AMT21 > check1 Then
                        err &= "傷害醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                    End If
                End If

            ElseIf Checkbox2.Checked = True AndAlso iPlanNo = "O" Then '申根方案

                If AMT21 > check1 Then
                    err &= "傷害醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If

                If AMT32 > check1 Then
                    err &= "海外突發疾病住院醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If

                If iPlanNo = "O" AndAlso AMT31 > check1 Then
                    err &= "海外突發疾病住院醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If

                '檢核有沒有超過死亡及失能1/4
                check2 = AMT11 * 0.25
                If AMT44 > check2 Then
                    err &= "重大燒燙傷保險金，建議不超過身故失能保險金4分之1;" & Environment.NewLine
                End If

            ElseIf Checkbox2.Checked = True AndAlso iPlanNo <> "O" Then '申根方案
                err &= "[申根方案]超過75歲建議不投保此保險;" & Environment.NewLine
            Else
                If iPlanNo = "K" Then
                    err &= "超過75歲建議不投保此保險;" & Environment.NewLine
                Else
                    'O 海外小資型 沒有大眾運輸
                    '檢核是否為死亡及失能的1/5
                    check2 = AMT11 * 0.2
                    If iPlanNo <> "O" AndAlso (AMT43 <> check2) Then
                        err &= "搭乘大眾運輸意外保險金建議投保 " & check2 & "萬;" & Environment.NewLine
                    End If

                    If iPlanNo = "O" Then
                        '檢核有沒有超過死亡及失能1/4
                        check2 = AMT11 * 0.25
                        If AMT44 > check2 Then
                            err &= "重大燒燙傷保險金，建議不超過身故失能保險金4分之1;" & Environment.NewLine
                        End If
                    Else
                        '檢核是否為死亡及失能的1/4
                        check2 = AMT11 * 0.25
                        If AMT44 <> check2 Then
                            err &= "重大燒燙傷保險金建議投保 " & check2 & "萬;" & Environment.NewLine
                        End If
                    End If

                    If AMT21 > check1 Then
                        err &= "傷害醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                    End If

                    If AMT32 > check1 Then
                        err &= "海外突發疾病住院醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                    End If

                    If iPlanNo = "O" AndAlso AMT31 > check1 Then
                        err &= "海外突發疾病住院醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                    End If
                End If
            End If

            '---------------------------------------------------------------------------------------------------------------------
        ElseIf AgeNow >= 71 Then
            'O 海外小資型 沒有大眾運輸
            '檢核是否為死亡及失能的1/5
            check2 = AMT11 * 0.2
            If iPlanNo <> "O" AndAlso (AMT43 <> check2) Then
                err &= "搭乘大眾運輸意外保險金建議投保 " & check2 & "萬;" & Environment.NewLine
            End If

            If TRAVEL_A.SelectedValue = "400" Then

                If IISSU_TP = "5" Then
                    If AMT11 > 2000000 Then
                        err &= "外國人持護照投保 意外身故失能保險金，建議保額最高 200萬;" & Environment.NewLine
                    End If

                    If AMT44 > 500000 Then
                        err &= "外國人持護照投保 重大燒燙傷保險金，建議保額最高 50萬;" & Environment.NewLine
                    End If

                    If AMT21 > 50000 Then
                        err &= "外國人持護照投保 傷害醫療保險金，建議保額最高 5萬;" & Environment.NewLine
                    End If
                Else
                    If iPlanNo = "O" Then
                        '檢核有沒有超過死亡及失能1/4
                        check2 = AMT11 * 0.25
                        If AMT44 > check2 Then
                            err &= "重大燒燙傷保險金，建議不超過身故失能保險金4分之1;" & Environment.NewLine
                        End If
                    Else
                        '檢核是否為死亡及失能的1/4
                        check2 = AMT11 * 0.25
                        If AMT44 <> check2 Then
                            err &= "重大燒燙傷保險金建議投保 " & check2 & "萬;" & Environment.NewLine
                        End If
                    End If

                    If AMT21 > check1 Then
                        err &= "傷害醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                    End If
                End If

            ElseIf Checkbox2.Checked = True Then '申根方案
                If iPlanNo = "O" Then
                    '檢核有沒有超過死亡及失能1/4
                    check2 = AMT11 * 0.25
                    If AMT44 > check2 Then
                        err &= "重大燒燙傷保險金，建議不超過身故失能保險金4分之1;" & Environment.NewLine
                    End If
                Else
                    '檢核是否為死亡及失能的1/4
                    check2 = AMT11 * 0.25
                    If AMT44 <> check2 Then
                        err &= "[申根地區]重大燒燙傷保險金建議投保 " & check2 & "萬;" & Environment.NewLine
                    End If
                End If

                If AMT21 <> 1500000 Then
                    err &= "[申根地區]傷害醫療保險金，建議投保 150萬; " & Environment.NewLine
                End If

                If iPlanNo = "O" Then
                    If AMT31 = 0 And AMT32 = 0 Then
                        err &= "[申根地區]海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    ElseIf AMT31 <> 0 AndAlso AMT31 <> 1000000 Then
                        err &= "[申根地區]海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    ElseIf AMT32 <> 0 AndAlso AMT32 <> 1000000 Then
                        err &= "[申根地區]海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    End If
                Else
                    If AMT32 <> 1000000 Then
                        err &= "[申根地區]海外突發疾病住院醫療保險金，建議投保 100萬;" & Environment.NewLine
                    End If
                End If

            Else
                If iPlanNo = "O" Then
                    '檢核有沒有超過死亡及失能1/4
                    check2 = AMT11 * 0.25
                    If AMT44 > check2 Then
                        err &= "重大燒燙傷保險金，建議不超過身故失能保險金4分之1;" & Environment.NewLine
                    End If
                Else
                    '檢核是否為死亡及失能的1/4
                    check2 = AMT11 * 0.25
                    If AMT44 <> check2 Then
                        err &= "重大燒燙傷保險金建議投保 " & check2 & "萬;" & Environment.NewLine
                    End If
                End If

                If AMT21 > check1 Then
                    err &= "傷害醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If

                If AMT32 > check1 Then
                    err &= "海外突發疾病住院醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If

                If iPlanNo = "O" AndAlso AMT31 > check1 Then
                    err &= "海外突發疾病住院醫療保險金，建議不超過身故失能保險金10分之1;" & Environment.NewLine
                End If
            End If
        End If
        Return err

    End Function


    Private Sub ErrMail(ByVal mTitle As String, ByVal mBody As String)
        mTitle = "B2B(BHAS101003)錯誤_" + mTitle
        If Not CBool(ConfigurationManager.AppSettings("bolWsITPSKL")) Then
            mTitle = "【測試信件】" + mTitle
        End If
        Dim mFrom As String = "SKERP<skerp@skinsurance.com.tw>"
        Dim mTo As String = "ski93217@skinsurance.com.tw;dave@skinsurance.com.tw"

        Me.InvokeCommonService("PUBService.bsCMNSendMail", mTitle, mBody, mFrom, mTo, "", "", "", "", "")
    End Sub

    Public Function CheckFireBurn(ByVal IT414Z As String, ByVal IT44 As String) As String

        Dim err As String = ""

        Dim AMT11 As Decimal = GetDecimal(IT414Z) '死亡及失能
        Dim AMT44 As Decimal = GetDecimal(IT44) '重大傷燙

        Dim check1 As Decimal
        '先檢核有沒有超過死亡及失能1/4
        check1 = AMT11 * 0.25

        If AMT44 > check1 Then
            err &= "重大燒燙傷保險金，上限不得超過死亡及失能保額4分之1;"
        End If

        Return err

    End Function
    Function TransRocDate(ByVal obj As Object) As String
        If obj Is Nothing Then
            Return ""
        End If
        If obj.ToString.Trim() = "" Then
            Return ""
        End If

        Dim dTmp As DateTime = CType(obj, DateTime)
        Return (dTmp.Year - 1911).ToString() + "/" + dTmp.Month.ToString("00") + "/" + dTmp.Day.ToString("00")

    End Function

#Region "投保方案Function"
    'BT1
    Public Sub BT1_BUTTON(ByVal age As String)
        Try
            GetDay() '計算天數
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "A 國內海陸型"
            Me.lblTargetPlan.Text = "A 國內海陸型"
            Me.titIT6D.Attributes.Add("style", "display:none")
            Me.lblIT6D.Attributes.Add("style", "display:none")
            Me.lblUnitIT6D.Attributes.Add("style", "display:none")

            IISSU_N.Enabled = True

            If ddlIT41.SelectedValue <> "" Then
                ddlIT4Y.SelectedValue = ""
                GetIT44FromIT41()
                GetIT43FromIT41()
            ElseIf ddlIT4Y.SelectedValue <> "" Then
                ddlIT41.SelectedValue = ""
                GetIT44FromIT4Y()
                GetIT43FromIT4Y()
            End If

            'ITP 不便險取保費的方式
            '1:InsType(險種)
            '2:InsDay(天數)
            '3.InsCoverage 保額 元
            '4.InsMdeduct 自負額 元
            '5.CountryType 1=國內 2=國外

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "50000"
            InsMdeduct = "2500"

            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '==========IT28==========
            InsType = "IT28"
            InsCoverage = "3000"
            InsMdeduct = "0"
            Dim ds2 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds2.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If

            '==========IT2D==========
            'InsType = "IT2D"
            'InsCoverage = "3000"
            'InsMdeduct = "0"
            'Dim ds3 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds3.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT2D.Text = ds3.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT2D.Text = "0"
            'End If
            '==========IT2G==========
            InsType = "IT2G"
            InsCoverage = "3000"
            InsMdeduct = "0"
            Dim ds4 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds4.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2G.Text = ds4.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2G.Text = "0"
            End If
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    'BT2
    Public Sub BT2_BUTTON(ByVal age As String)
        Try

            GetDay() '計算天數
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "B 國內飛翔型"
            IISSU_N.Enabled = True
            Me.titIT6D.Attributes.Add("style", "display:none")
            Me.lblIT6D.Attributes.Add("style", "display:none")
            Me.lblUnitIT6D.Attributes.Add("style", "display:none")

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "50000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "1500"
            InsMdeduct = "0"

            Dim ds2 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds2.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "1000"
            InsMdeduct = "0"

            Dim dsA As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsA.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = dsA.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If
            '====================================
            InsType = "IT27"
            InsCoverage = "2000"
            InsMdeduct = "0"

            Dim ds4 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds4.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT27.Text = ds4.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT27.Text = "0"
            End If
            '====================================
            InsType = "IT28"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds5 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds5.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds5.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If
            '==========IT2D==========
            'InsType = "IT2D"
            'InsCoverage = "3000"
            'InsMdeduct = "0"
            'Dim dsE As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If dsE.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT2D.Text = dsE.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT2D.Text = "0"
            'End If
            '==========IT2G==========
            InsType = "IT2G"
            InsCoverage = "3000"
            InsMdeduct = "0"
            Dim dsG As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsG.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2G.Text = dsG.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2G.Text = "0"
            End If
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    'BT3
    Public Sub BT3_BUTTON(ByVal age As String)

        Try

            GetDay() '計算天數
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "C 海外基本型"
            Me.lblTargetPlan.Text = "C 海外基本型"
            IISSU_N.Enabled = True

            '設定海突相關控制項顯示
            Dim dobj As DiseaseFields = New DiseaseFields()
            dobj.IT61 = True
            dobj.IT63 = False
            dobj.IT6E = True
            dobj.IT64 = False
            dobj.IT65 = False
            dobj.IT66 = False
            dobj.IT6D = True
            dobj.IT6F = True
            dobj.IT6G = True
            FieldsSwtich_Disease(dobj)

            Me.ddlIT43.Enabled = False
            Me.ddlIT43.SelectedValue = ""
            Me.txtIT43.Text = ""

            If ddlIT41.SelectedValue <> "" Then
                ddlIT4Y.SelectedValue = ""
                GetIT44FromIT41()
            ElseIf ddlIT4Y.SelectedValue <> "" Then
                ddlIT41.SelectedValue = ""
                GetIT44FromIT4Y()
            End If

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "200000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '====================================
            InsType = "IT2R"
            InsCoverage = "1000000"
            InsMdeduct = "0"

            Dim ds2R As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2R.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2R.Text = ds2R.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2R.Text = "0"
            End If
            '====================================
            InsType = "IT37"
            InsCoverage = "10000"
            InsMdeduct = "0"

            Dim ds2K As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2K.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT37.Text = ds2K.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT37.Text = "0"
            End If
            '====================================
            InsType = "IT36"
            InsCoverage = "20000"
            InsMdeduct = "0"

            Dim ds2J As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2J.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT36.Text = ds2J.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT36.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "4000"
            InsMdeduct = "0"

            Dim ds2A As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2A.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = ds2A.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If
            '====================================
            InsType = "IT3B"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2L As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2L.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3B.Text = ds2L.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3B.Text = "0"
            End If
            '====================================
            InsType = "IT3C"
            InsCoverage = "10000"
            InsMdeduct = "0"

            Dim ds3C As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3C.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3C.Text = ds3C.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3C.Text = "0"
            End If
            '====================================
            InsType = "IT3D"
            InsCoverage = "1000"
            InsMdeduct = "0"

            Dim ds3D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3D.Text = ds3D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3D.Text = "0"
            End If
            '====================================
            InsType = "IT3E"
            InsCoverage = "1000"
            InsMdeduct = "0"

            Dim ds3E As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3E.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3E.Text = ds3E.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3E.Text = "0"
            End If
            '====================================
            InsType = "IT39"
            InsCoverage = "4000"
            InsMdeduct = "0"

            Dim ds29 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds29.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT39.Text = ds29.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT39.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "4000"
            InsMdeduct = "0"

            Dim ds2X As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2X.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds2X.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If

            '====================================
            InsType = "IT2M"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds25 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds25.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2M.Text = ds25.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2M.Text = "0"
            End If
            '====================================
            InsType = "IT27"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds27 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds27.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT27.Text = ds27.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT27.Text = "0"
            End If
            '====================================
            InsType = "IT2D"
            InsCoverage = "20000"
            InsMdeduct = "0"

            Dim dsIT2D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsIT2D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2D.Text = dsIT2D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2D.Text = "0"
            End If
            '====================================
            InsType = "IT28"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds28 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds28.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds28.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If
            '====================================
            InsType = "IT31"
            InsCoverage = "50000"
            InsMdeduct = "0"

            Dim ds31 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds31.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT31.Text = ds31.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT31.Text = "0"
            End If

            '====================================

        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    'BT4
    Public Sub BT4_BUTTON(ByVal age As String, ByVal instyp As String)
        Try

            GetDay() '計算天數
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "D 海外自駕型"

            IISSU_N.Enabled = True

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "500000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '====================================
            InsType = "IT3B"
            InsCoverage = "1000"
            InsMdeduct = "0"

            Dim ds2L As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2L.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3B.Text = ds2L.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3B.Text = "0"
            End If

            '====================================
            InsType = "IT39"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds29 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds29.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT39.Text = ds29.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT39.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2A As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2A.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = ds2A.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If
            '====================================
            InsType = "IT37"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim ds2K As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2K.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT37.Text = ds2K.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT37.Text = "0"
            End If
            '====================================
            InsType = "IT36"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim ds2J As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2J.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT36.Text = ds2J.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT36.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds2X As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2X.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds2X.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If
            '====================================
            InsType = "IT2M"
            InsCoverage = "2000"
            InsMdeduct = "0"

            Dim ds5 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds5.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2M.Text = ds5.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2M.Text = "0"
            End If
            '====================================
            InsType = "IT27"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds4 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds4.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT27.Text = ds4.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT27.Text = "0"
            End If
            '====================================
            InsType = "IT2N"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim dsN As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsN.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2N.Text = dsN.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2N.Text = "0"
            End If
            '====================================

        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    'BT5
    Public Sub BT5_BUTTON(ByVal age As String)
        Try
            GetDay() '計算天數
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "D 海外超值型"
            Me.lblTargetPlan.Text = "D 海外超值型"

            '設定海突相關控制項顯示
            Dim dobj As DiseaseFields = New DiseaseFields()
            dobj.IT61 = True
            dobj.IT63 = False
            dobj.IT6E = True
            dobj.IT64 = False
            dobj.IT65 = False
            dobj.IT66 = False
            dobj.IT6D = True
            dobj.IT6F = True
            dobj.IT6G = True
            FieldsSwtich_Disease(dobj)

            Me.ddlIT43.Enabled = False
            Me.ddlIT43.SelectedValue = ""
            Me.txtIT43.Text = ""


            IISSU_N.Enabled = True

            If ddlIT41.SelectedValue <> "" Then
                ddlIT4Y.SelectedValue = ""
                GetIT44FromIT41()
            ElseIf ddlIT4Y.SelectedValue <> "" Then
                ddlIT41.SelectedValue = ""
                GetIT44FromIT4Y()
            End If

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "200000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '====================================
            InsType = "IT2R"
            InsCoverage = "3000000"
            InsMdeduct = "0"

            Dim ds2R As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2R.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2R.Text = ds2R.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2R.Text = "0"
            End If
            '====================================
            InsType = "IT37"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim ds2K As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2K.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT37.Text = ds2K.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT37.Text = "0"
            End If
            '====================================
            InsType = "IT36"
            InsCoverage = "60000"
            InsMdeduct = "0"

            Dim ds2J As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2J.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT36.Text = ds2J.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT36.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds2A As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2A.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = ds2A.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If
            '====================================
            InsType = "IT3B"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2L As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2L.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3B.Text = ds2L.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3B.Text = "0"
            End If
            '====================================
            InsType = "IT3C"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim ds3C As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3C.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3C.Text = ds3C.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3C.Text = "0"
            End If
            '====================================
            InsType = "IT3D"
            InsCoverage = "2000"
            InsMdeduct = "0"

            Dim ds3D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3D.Text = ds3D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3D.Text = "0"
            End If
            '====================================
            InsType = "IT3E"
            InsCoverage = "2000"
            InsMdeduct = "0"

            Dim ds3E As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3E.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3E.Text = ds3E.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3E.Text = "0"
            End If
            '====================================
            InsType = "IT39"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds29 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds29.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT39.Text = ds29.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT39.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds2X As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2X.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds2X.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If

            '====================================
            InsType = "IT2M"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds25 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds25.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2M.Text = ds25.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2M.Text = "0"
            End If
            '====================================
            InsType = "IT27"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds27 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds27.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT27.Text = ds27.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT27.Text = "0"
            End If
            '====================================
            InsType = "IT2D"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim dsIT2D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsIT2D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2D.Text = dsIT2D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2D.Text = "0"
            End If
            '====================================
            InsType = "IT28"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds28 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds28.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds28.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If
            '====================================
            InsType = "IT31"
            InsCoverage = "50000"
            InsMdeduct = "0"

            Dim ds31 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds31.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT31.Text = ds31.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT31.Text = "0"
            End If

            '====================================


        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Public Sub BT6_BUTTON(ByVal age As String)
        Try
            GetDay() '計算天數
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "J 海外標準型"
            Me.lblTargetPlan.Text = "J 海外標準型"
            IISSU_N.Enabled = True

            '設定海突相關控制項顯示
            Dim dobj As DiseaseFields = New DiseaseFields()
            dobj.IT61 = False
            dobj.IT63 = False
            dobj.IT6E = False
            dobj.IT64 = True
            dobj.IT65 = True
            dobj.IT66 = True
            dobj.IT6D = False
            dobj.IT6F = False
            dobj.IT6G = False
            FieldsSwtich_Disease(dobj)

            If ddlIT41.SelectedValue <> "" Then
                ddlIT4Y.SelectedValue = ""
                GetIT44FromIT41(25)
                GetIT43FromIT41(20)
            ElseIf ddlIT4Y.SelectedValue <> "" Then
                ddlIT41.SelectedValue = ""
                GetIT44FromIT4Y(0)
                GetIT43FromIT4Y(0)
            End If

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "3000000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '====================================
            InsType = "IT2R"
            InsCoverage = "2500000"
            InsMdeduct = "0"

            Dim ds2R As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2R.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2R.Text = ds2R.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2R.Text = "0"
            End If

            '====================================
            InsType = "IT36" '原IT32
            InsCoverage = "120000"
            InsMdeduct = "0"

            Dim ds32 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds32.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT36.Text = ds32.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT36.Text = "0"
            End If

            '====================================
            InsType = "IT37" '原IT33
            InsCoverage = "120000"
            InsMdeduct = "0"

            Dim ds33 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds33.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT37.Text = ds33.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT37.Text = "0"
            End If
            '====================================
            InsType = "IT3B"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2L As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2L.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3B.Text = ds2L.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3B.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "4000"
            InsMdeduct = "0"

            Dim ds2A As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2A.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = ds2A.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If
            '====================================
            InsType = "IT39"
            InsCoverage = "4000"
            InsMdeduct = "0"

            Dim ds29 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds29.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT39.Text = ds29.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT39.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "4000"
            InsMdeduct = "0"

            Dim ds2X As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2X.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds2X.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If

            '====================================
            InsType = "IT2M"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds25 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds25.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2M.Text = ds25.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2M.Text = "0"
            End If
            '====================================
            InsType = "IT27"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds27 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds27.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT27.Text = ds27.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT27.Text = "0"
            End If
            '====================================
            InsType = "IT2D"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim dsIT2D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsIT2D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2D.Text = dsIT2D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2D.Text = "0"
            End If
            '====================================
            InsType = "IT28"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds28 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds28.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds28.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If
            '====================================
            InsType = "IT31"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim ds31 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds31.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT31.Text = ds31.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT31.Text = "0"
            End If

            '====================================
            InsType = "IT34"
            InsCoverage = "120000"
            InsMdeduct = "0"

            Dim ds34 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds34.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT34.Text = ds34.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT34.Text = "0"
            End If

            '====================================
            InsType = "IT35"
            InsCoverage = "120000"
            InsMdeduct = "0"

            Dim ds35 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds35.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT35.Text = ds35.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT35.Text = "0"
            End If

            '====================================
            InsType = "IT2P"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2P As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2P.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2P.Text = ds2P.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2P.Text = "0"
            End If

            '====================================
            InsType = "IT2U"
            InsCoverage = "10000"
            InsMdeduct = "0"

            Dim ds2U As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2U.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2U.Text = ds2U.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2U.Text = "0"
            End If

            '====================================
            InsType = "IT2T"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2T As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2T.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2T.Text = ds2T.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2T.Text = "0"
            End If

            '====================================
            InsType = "IT2S"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2S As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2S.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2S.Text = ds2S.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2S.Text = "0"
            End If

            '====================================

        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Public Sub BT7_BUTTON(ByVal age As String)
        Try
            GetDay() '計算天數
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "K 海外豪華型"
            Me.lblTargetPlan.Text = "K 海外豪華型"

            '設定海突相關控制項顯示
            Dim dobj As DiseaseFields = New DiseaseFields()
            dobj.IT61 = False
            dobj.IT63 = False
            dobj.IT6E = False
            dobj.IT64 = True
            dobj.IT65 = True
            dobj.IT66 = True
            dobj.IT6D = False
            dobj.IT6F = False
            dobj.IT6G = False
            FieldsSwtich_Disease(dobj)

            IISSU_N.Enabled = True

            If ddlIT41.SelectedValue <> "" Then
                ddlIT4Y.SelectedValue = ""
                GetIT44FromIT41(25)
                GetIT43FromIT41(20)
            ElseIf ddlIT4Y.SelectedValue <> "" Then
                ddlIT41.SelectedValue = ""
                GetIT44FromIT4Y(0)
                GetIT43FromIT4Y(0)
            End If

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "5000000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '====================================
            InsType = "IT2R"
            InsCoverage = "2500000"
            InsMdeduct = "0"

            Dim ds2R As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2R.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2R.Text = ds2R.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2R.Text = "0"
            End If

            '====================================
            InsType = "IT36" '原IT32
            InsCoverage = "150000"
            InsMdeduct = "0"

            Dim ds32 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds32.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT36.Text = ds32.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT36.Text = "0"
            End If

            '====================================
            InsType = "IT37" '原IT33
            InsCoverage = "150000"
            InsMdeduct = "0"

            Dim ds33 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds33.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT37.Text = ds33.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT37.Text = "0"
            End If
            '====================================
            InsType = "IT3B"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2L As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2L.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3B.Text = ds2L.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3B.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds2A As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2A.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = ds2A.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If
            '====================================
            InsType = "IT39"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds29 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds29.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT39.Text = ds29.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT39.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds2X As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2X.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds2X.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If

            '====================================
            InsType = "IT2M"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds25 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds25.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2M.Text = ds25.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2M.Text = "0"
            End If
            '====================================
            InsType = "IT27"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds27 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds27.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT27.Text = ds27.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT27.Text = "0"
            End If
            '====================================
            InsType = "IT2D"
            InsCoverage = "50000"
            InsMdeduct = "0"

            Dim dsIT2D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsIT2D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2D.Text = dsIT2D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2D.Text = "0"
            End If
            '====================================
            InsType = "IT28"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds28 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds28.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds28.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If
            '====================================
            InsType = "IT31"
            InsCoverage = "50000"
            InsMdeduct = "0"

            Dim ds31 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds31.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT31.Text = ds31.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT31.Text = "0"
            End If

            '====================================
            InsType = "IT34"
            InsCoverage = "150000"
            InsMdeduct = "0"

            Dim ds34 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds34.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT34.Text = ds34.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT34.Text = "0"
            End If

            '====================================
            InsType = "IT35"
            InsCoverage = "150000"
            InsMdeduct = "0"

            Dim ds35 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds35.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT35.Text = ds35.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT35.Text = "0"
            End If

            '====================================
            InsType = "IT2P"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds2P As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2P.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2P.Text = ds2P.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2P.Text = "0"
            End If

            '====================================
            InsType = "IT2U"
            InsCoverage = "10000"
            InsMdeduct = "0"

            Dim ds2U As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2U.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2U.Text = ds2U.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2U.Text = "0"
            End If

            '====================================
            InsType = "IT2T"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds2T As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2T.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2T.Text = ds2T.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2T.Text = "0"
            End If

            '====================================
            InsType = "IT2S"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds2S As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2S.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2S.Text = ds2S.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2S.Text = "0"
            End If

            '====================================


        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Public Sub BT8_BUTTON(ByVal age As String)
        Try
            GetDay() '計算天數
            Me.txtCase.Text = "L 一般型或申根型"
            Me.lblTargetPlan.Text = "L 一般型或申根型"

            '設定海突相關控制項顯示
            Dim dobj As DiseaseFields = New DiseaseFields()
            dobj.IT61 = False
            dobj.IT63 = False
            dobj.IT6E = False
            dobj.IT64 = True
            dobj.IT65 = True
            dobj.IT66 = True
            dobj.IT6D = False
            dobj.IT6F = False
            dobj.IT6G = False
            FieldsSwtich_Disease(dobj)

            IISSU_N.Enabled = True

            If ddlIT41.SelectedValue <> "" Then
                ddlIT4Y.SelectedValue = ""
                GetIT44FromIT41(25)
                GetIT43FromIT41(20)
            ElseIf ddlIT4Y.SelectedValue <> "" Then
                ddlIT41.SelectedValue = ""
                GetIT44FromIT4Y(0)
                GetIT43FromIT4Y(0)
            End If
        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Public Sub BT9_BUTTON(ByVal age As String)

        Try

            GetDay() '計算天數
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "M 海外安心型"
            Me.lblTargetPlan.Text = "M 海外安心型"
            IISSU_N.Enabled = True

            '設定海突相關控制項顯示
            Dim dobj As DiseaseFields = New DiseaseFields()
            dobj.IT61 = True
            dobj.IT63 = False
            dobj.IT6E = False
            dobj.IT64 = True
            dobj.IT65 = True
            dobj.IT66 = True
            dobj.IT6D = True
            dobj.IT6F = False
            dobj.IT6G = False
            FieldsSwtich_Disease(dobj)

            Me.ddlIT43.Enabled = False
            Me.ddlIT43.SelectedValue = ""
            Me.txtIT43.Text = ""

            If ddlIT41.SelectedValue <> "" Then
                ddlIT4Y.SelectedValue = ""
                GetIT44FromIT41()
            ElseIf ddlIT4Y.SelectedValue <> "" Then
                ddlIT41.SelectedValue = ""
                GetIT44FromIT4Y()
            End If

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "200000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '====================================
            InsType = "IT2R"
            InsCoverage = "1000000"
            InsMdeduct = "0"

            Dim ds2R As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2R.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2R.Text = ds2R.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2R.Text = "0"
            End If
            '====================================
            InsType = "IT37"
            InsCoverage = "10000"
            InsMdeduct = "0"

            Dim ds2K As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2K.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT37.Text = ds2K.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT37.Text = "0"
            End If
            '====================================
            InsType = "IT36"
            InsCoverage = "20000"
            InsMdeduct = "0"

            Dim ds2J As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2J.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT36.Text = ds2J.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT36.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "4000"
            InsMdeduct = "0"

            Dim ds2A As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2A.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = ds2A.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If
            '====================================
            InsType = "IT3B"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2L As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2L.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3B.Text = ds2L.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3B.Text = "0"
            End If
            '====================================
            InsType = "IT3C"
            InsCoverage = "10000"
            InsMdeduct = "0"

            Dim ds3C As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3C.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3C.Text = ds3C.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3C.Text = "0"
            End If
            '====================================
            InsType = "IT3D"
            InsCoverage = "1000"
            InsMdeduct = "0"

            Dim ds3D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3D.Text = ds3D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3D.Text = "0"
            End If
            '====================================
            InsType = "IT3E"
            InsCoverage = "1000"
            InsMdeduct = "0"

            Dim ds3E As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3E.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3E.Text = ds3E.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3E.Text = "0"
            End If
            '====================================
            InsType = "IT39"
            InsCoverage = "4000"
            InsMdeduct = "0"

            Dim ds29 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds29.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT39.Text = ds29.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT39.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "4000"
            InsMdeduct = "0"

            Dim ds2X As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2X.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds2X.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If

            '====================================
            InsType = "IT2M"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds25 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds25.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2M.Text = ds25.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2M.Text = "0"
            End If
            '====================================
            InsType = "IT27"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds27 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds27.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT27.Text = ds27.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT27.Text = "0"
            End If
            '====================================
            InsType = "IT2D"
            InsCoverage = "20000"
            InsMdeduct = "0"

            Dim dsIT2D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsIT2D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2D.Text = dsIT2D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2D.Text = "0"
            End If
            '====================================
            InsType = "IT28"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds28 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds28.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds28.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If
            '====================================
            InsType = "IT31"
            InsCoverage = "50000"
            InsMdeduct = "0"

            Dim ds31 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds31.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT31.Text = ds31.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT31.Text = "0"
            End If

            '====================================

        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    Public Sub BT10_BUTTON(ByVal age As String)
        Try
            GetDay() '計算天數.
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "N 海外特選型"
            Me.lblTargetPlan.Text = "N 海外特選型"

            '設定海突相關控制項顯示
            Dim dobj As DiseaseFields = New DiseaseFields()
            dobj.IT61 = True
            dobj.IT63 = False
            dobj.IT6E = False
            dobj.IT64 = True
            dobj.IT65 = True
            dobj.IT66 = True
            dobj.IT6D = True
            dobj.IT6F = False
            dobj.IT6G = False
            FieldsSwtich_Disease(dobj)

            Me.ddlIT43.Enabled = False
            Me.ddlIT43.SelectedValue = ""
            Me.txtIT43.Text = ""


            IISSU_N.Enabled = True

            If ddlIT41.SelectedValue <> "" Then
                ddlIT4Y.SelectedValue = ""
                GetIT44FromIT41()
            ElseIf ddlIT4Y.SelectedValue <> "" Then
                ddlIT41.SelectedValue = ""
                GetIT44FromIT4Y()
            End If

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "200000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '====================================
            InsType = "IT2R"
            InsCoverage = "3000000"
            InsMdeduct = "0"

            Dim ds2R As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2R.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2R.Text = ds2R.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2R.Text = "0"
            End If
            '====================================
            InsType = "IT37"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim ds2K As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2K.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT37.Text = ds2K.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT37.Text = "0"
            End If
            '====================================
            InsType = "IT36"
            InsCoverage = "60000"
            InsMdeduct = "0"

            Dim ds2J As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2J.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT36.Text = ds2J.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT36.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds2A As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2A.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = ds2A.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If
            '====================================
            InsType = "IT3B"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2L As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2L.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3B.Text = ds2L.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3B.Text = "0"
            End If
            '====================================
            InsType = "IT3C"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim ds3C As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3C.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3C.Text = ds3C.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3C.Text = "0"
            End If
            '====================================
            InsType = "IT3D"
            InsCoverage = "2000"
            InsMdeduct = "0"

            Dim ds3D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3D.Text = ds3D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3D.Text = "0"
            End If
            '====================================
            InsType = "IT3E"
            InsCoverage = "2000"
            InsMdeduct = "0"

            Dim ds3E As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3E.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3E.Text = ds3E.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3E.Text = "0"
            End If
            '====================================
            InsType = "IT39"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds29 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds29.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT39.Text = ds29.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT39.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds2X As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2X.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds2X.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If

            '====================================
            InsType = "IT2M"
            InsCoverage = "6000"
            InsMdeduct = "0"

            Dim ds25 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds25.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2M.Text = ds25.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2M.Text = "0"
            End If
            '====================================
            InsType = "IT27"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds27 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds27.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT27.Text = ds27.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT27.Text = "0"
            End If
            '====================================
            InsType = "IT2D"
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim dsIT2D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsIT2D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2D.Text = dsIT2D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2D.Text = "0"
            End If
            '====================================
            InsType = "IT28"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds28 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds28.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds28.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If
            '====================================
            InsType = "IT31"
            InsCoverage = "50000"
            InsMdeduct = "0"

            Dim ds31 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds31.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT31.Text = ds31.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT31.Text = "0"
            End If

            '====================================


        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Public Sub BT11_BUTTON(ByVal age As String)
        Try
            GetDay() '計算天數
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            Me.txtCase.Text = "O 海外小資型"
            Me.lblTargetPlan.Text = "O 海外小資型"

            '設定海突相關控制項顯示
            Dim dobj As DiseaseFields = New DiseaseFields()
            dobj.IT61 = True
            dobj.IT63 = False
            dobj.IT6E = False
            dobj.IT64 = True
            dobj.IT65 = True
            dobj.IT66 = True
            dobj.IT6D = True
            dobj.IT6F = False
            dobj.IT6G = False
            FieldsSwtich_Disease(dobj)

            Me.ddlIT43.Enabled = False
            Me.ddlIT43.SelectedValue = ""
            Me.txtIT43.Text = ""

            IISSU_N.Enabled = True

            If ddlIT41.SelectedValue <> "" Then
                ddlIT4Y.SelectedValue = ""
                GetIT44FromIT41()
            ElseIf ddlIT4Y.SelectedValue <> "" Then
                ddlIT41.SelectedValue = ""
                GetIT44FromIT4Y()
            End If

            InsType = "IT11"
            InsDay = txtIQPITP_Days.Text.Trim
            InsCoverage = "200000"
            InsMdeduct = "2500"
            '2016.2.16 (5集體國內 6集體國外) CHK5
            'ITP費率檔(個人集體直接) 1啟用 0停用 
            If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                If CHK5.Checked = True Then
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "5" '國外5 外6 別
                    Else
                        CountryType = "6" '國外5 外6 別
                    End If
                Else
                    If TRAVEL_A.SelectedValue = "400" Then
                        CountryType = "1" '國內1 外2 別
                    Else
                        CountryType = "2" '國內1 外2 別
                    End If
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If

            Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT11.Text = ds.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT11.Text = "0"
            End If

            '====================================
            InsType = "IT3D"
            InsCoverage = "1000"
            InsMdeduct = "0"

            Dim ds3D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds3D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3D.Text = ds3D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3D.Text = "0"
            End If

            '====================================
            InsType = "IT2R"
            InsCoverage = "2500000"
            InsMdeduct = "0"

            txtIT2R.Text = "0"
            'Dim ds2R As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds2R.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT2R.Text = ds2R.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT2R.Text = "0"
            'End If

            '====================================
            InsType = "IT36" '原IT32
            InsCoverage = "60000"
            InsMdeduct = "0"

            Dim ds32 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds32.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT36.Text = ds32.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT36.Text = "0"
            End If

            '====================================
            InsType = "IT37" '原IT33
            InsCoverage = "30000"
            InsMdeduct = "0"

            Dim ds33 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds33.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT37.Text = ds33.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT37.Text = "0"
            End If
            '====================================
            InsType = "IT3B"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds2L As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2L.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3B.Text = ds2L.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3B.Text = "0"
            End If
            '====================================
            InsType = "IT3A"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds2A As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2A.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT3A.Text = ds2A.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT3A.Text = "0"
            End If
            '====================================
            InsType = "IT39"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds29 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds29.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT39.Text = ds29.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT39.Text = "0"
            End If
            '====================================
            InsType = "IT38"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds2X As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds2X.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT38.Text = ds2X.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT38.Text = "0"
            End If

            '====================================
            InsType = "IT2M"
            InsCoverage = "5000"
            InsMdeduct = "0"

            Dim ds25 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds25.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2M.Text = ds25.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2M.Text = "0"
            End If
            '====================================
            InsType = "IT27"
            InsCoverage = "6000"
            InsMdeduct = "0"
            txtIT27.Text = "0"

            'Dim ds27 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds27.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT27.Text = ds27.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT27.Text = "0"
            'End If
            '====================================
            InsType = "IT2D"
            InsCoverage = "50000"
            InsMdeduct = "0"

            Dim dsIT2D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If dsIT2D.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT2D.Text = dsIT2D.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT2D.Text = "0"
            End If
            '====================================
            InsType = "IT28"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim ds28 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds28.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT28.Text = ds28.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT28.Text = "0"
            End If
            '====================================
            InsType = "IT31"
            InsCoverage = "50000"
            InsMdeduct = "0"

            Dim ds31 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            If ds31.Tables("ITPPrem").Rows.Count > 0 Then
                txtIT31.Text = ds31.Tables("ITPPrem").Rows(0)("PREM")
            Else
                txtIT31.Text = "0"
            End If

            '====================================
            InsType = "IT34"
            InsCoverage = "150000"
            InsMdeduct = "0"
            txtIT34.Text = "0"

            'Dim ds34 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds34.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT34.Text = ds34.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT34.Text = "0"
            'End If

            '====================================
            InsType = "IT35"
            InsCoverage = "150000"
            InsMdeduct = "0"
            txtIT35.Text = "0"

            'Dim ds35 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds35.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT35.Text = ds35.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT35.Text = "0"
            'End If

            '====================================
            InsType = "IT2P"
            InsCoverage = "5000"
            InsMdeduct = "0"
            txtIT2P.Text = "0"

            'Dim ds2P As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds2P.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT2P.Text = ds2P.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT2P.Text = "0"
            'End If

            '====================================
            InsType = "IT2U"
            InsCoverage = "10000"
            InsMdeduct = "0"
            txtIT2U.Text = "0"

            'Dim ds2U As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds2U.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT2U.Text = ds2U.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT2U.Text = "0"
            'End If

            '====================================
            InsType = "IT2T"
            InsCoverage = "5000"
            InsMdeduct = "0"
            txtIT2T.Text = "0"

            'Dim ds2T As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds2T.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT2T.Text = ds2T.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT2T.Text = "0"
            'End If

            '====================================
            InsType = "IT2S"
            InsCoverage = "5000"
            InsMdeduct = "0"
            txtIT2S.Text = "0"

            'Dim ds2S As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

            'If ds2S.Tables("ITPPrem").Rows.Count > 0 Then
            '    txtIT2S.Text = ds2S.Tables("ITPPrem").Rows(0)("PREM")
            'Else
            '    txtIT2S.Text = "0"
            'End If

            '====================================


        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

#End Region
    Public Function GetDecimal(ByVal oParameter As Object) As Decimal
        Try
            If oParameter Is Nothing OrElse oParameter Is DBNull.Value Then
                Return 0
            End If

            Return CType(oParameter, Decimal)
        Catch ex As Exception
            Return 0
        End Try

    End Function
    Public Sub LockPage()
        Try
            '查詢出來W C Y 的要鎖
            txtIins_IINS2.Enabled = False
            UcRocDobjbgn.Enabled = False
            UcRocDobjend.Enabled = False
            UcRocDobjbgn.Enabled =
            txtTobjbgn.Enabled = False
            txtTobjend.Enabled = False
            CHK5.Enabled = False
            TRAVEL_A.Enabled = False
            ddl_IINS.Enabled = False

            ddlCountry_A1.Enabled = False
            ddlCountry_A2.Enabled = False
            ddlCountry_B1.Enabled = False
            ddlCountry_B2.Enabled = False
            ddlCountry_C1.Enabled = False
            ddlCountry_C2.Enabled = False
            ddlCountry_D1.Enabled = False
            ddlCountry_D2.Enabled = False
            ddlCountry_E1.Enabled = False
            ddlCountry_E2.Enabled = False
            ddlCountry_F1.Enabled = False
            ddlCountry_F2.Enabled = False
            ddlCountry_G1.Enabled = False
            ddlCountry_G2.Enabled = False
            ddlCountry_H1.Enabled = False
            ddlCountry_H2.Enabled = False
            Checkbox2.Enabled = False
            Checkbox4.Enabled = False
            Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csUpdate, "").Enabled = False
            Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csQuery, "").Enabled = False
            Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csQuery, "").Enabled = False
            IISSU_N.Enabled = False

            ddlChkMailYN.Enabled = False
            txtIins_IINS2.Enabled = False
            txtIins_ISSU2.Enabled = False
            IINS_ADDR.Enabled = False
            TRAVEL_BEGIN.Enabled = False
            TRAVEL_END.Enabled = False
            CK1.Enabled = False
            ddl_IINS.Enabled = False
            NAME_IINS.Enabled = False
            Birtdy_IINS.Enabled = False
            IINS_AREA.Enabled = False
            ddl_IINS2.Enabled = False
            IINS_TEL.Enabled = False
            UcRocDobjbgn.Enabled = False
            UcRocDobjend.Enabled = False

            ddl_IISSU2.Enabled = False
            TRAVEL_A.Enabled = False
            ddl_IISSU.Enabled = False
            NAME_IISSU.Enabled = False
            UcRocDbirth.Enabled = False
            IISSU_AREA.Enabled = False
            IISSU_ADDR.Enabled = False
            IISSU_TEL.Enabled = False
            EMAIL.Enabled = False
            Relationship.Enabled = False
            NAME_IISSU.Enabled = False
            IISSU_N.Enabled = False
            IISSU_AREA.Enabled = False
            IISSU_ADDR.Enabled = False
            IISSU_TEL.Enabled = False
            EMAIL.Enabled = False
            Benefi.Enabled = False
            Benefi_ID.Enabled = False
            Relationship2.Enabled = False

            txtTELREL.Enabled = False
            txtZIPREL.Enabled = False
            txtAREL.Enabled = False
            Button3.Visible = False
            Button0.Visible = False
            Button2.Visible = False
            UP_BTN.Enabled = False

            IINSMAIL.Enabled = False
            INSMOBILE.Enabled = False
            ddlSex.Enabled = False
            btMore.Enabled = False
            txtPrincipal.Enabled = False
            Checkbox2.Enabled = False
            Checkbox4.Enabled = False
            btnGetFlightInfo.Enabled = False
        Catch ex As Exception
            Me.ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub
    Public Sub NOLockPage()
        Try
            '查詢出來R的不要鎖
            txtIins_IINS2.Enabled = True
            UcRocDobjbgn.Enabled = True
            UcRocDobjend.Enabled = True
            txtTobjbgn.Enabled = True
            txtTobjend.Enabled = True
            CHK5.Enabled = True
            TRAVEL_A.Enabled = True
            ddl_IINS.Enabled = True

            ddlCountry_A1.Enabled = True
            ddlCountry_A2.Enabled = True
            ddlCountry_B1.Enabled = True
            ddlCountry_B2.Enabled = True
            ddlCountry_C1.Enabled = True
            ddlCountry_C2.Enabled = True
            ddlCountry_D1.Enabled = True
            ddlCountry_D2.Enabled = True
            ddlCountry_E1.Enabled = True
            ddlCountry_E2.Enabled = True
            ddlCountry_F1.Enabled = True
            ddlCountry_F2.Enabled = True
            ddlCountry_G1.Enabled = True
            ddlCountry_G2.Enabled = True
            ddlCountry_H1.Enabled = True
            ddlCountry_H2.Enabled = True
            Checkbox2.Enabled = True
            Checkbox4.Enabled = True
            Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csUpdate, "").Enabled = True
            Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csQuery, "").Enabled = True
            Me.StateCommandBar.CommandState.FindByNameArgument(UICommands.csQuery, "").Enabled = True
            IISSU_N.Enabled = True

            '20260923 文彥郵件回饋，鎖死不可選，預設為"是"
            'ddlChkMailYN.Enabled = True
            txtIins_IINS2.Enabled = True
            txtIins_ISSU2.Enabled = True
            IINS_ADDR.Enabled = True
            TRAVEL_BEGIN.Enabled = True
            TRAVEL_END.Enabled = True
            CK1.Enabled = True
            ddl_IINS.Enabled = True
            NAME_IINS.Enabled = True
            Birtdy_IINS.Enabled = True
            IINS_AREA.Enabled = True
            ddl_IINS2.Enabled = True
            IINS_TEL.Enabled = True
            UcRocDobjbgn.Enabled = True
            UcRocDobjend.Enabled = True

            ddl_IISSU2.Enabled = True
            TRAVEL_A.Enabled = True
            ddl_IISSU.Enabled = True
            NAME_IISSU.Enabled = True
            UcRocDbirth.Enabled = True
            IISSU_AREA.Enabled = True
            IISSU_ADDR.Enabled = True
            IISSU_TEL.Enabled = True
            EMAIL.Enabled = True
            Relationship.Enabled = True
            NAME_IISSU.Enabled = True
            IISSU_N.Enabled = True
            IISSU_AREA.Enabled = True
            IISSU_ADDR.Enabled = True
            IISSU_TEL.Enabled = True
            EMAIL.Enabled = True
            Benefi.Enabled = True
            Benefi_ID.Enabled = True
            Relationship2.Enabled = True

            txtTELREL.Enabled = True
            txtZIPREL.Enabled = True
            txtAREL.Enabled = True
            Button3.Visible = True
            Button0.Visible = True
            Button2.Visible = True
            UP_BTN.Enabled = True

            IINSMAIL.Enabled = True
            INSMOBILE.Enabled = True
            ddlSex.Enabled = True
            btMore.Enabled = True

            txtPrincipal.Enabled = True
        Catch ex As Exception
            Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub

    Public Sub IISSU_Clear()
        ddl_IISSU.SelectedValue = ""
        txtIins_ISSU2.Text = ""
        ddl_IISSU2.SelectedValue = ""
        NAME_IISSU.Text = ""
        UcRocDbirth.Text = ""
        IISSU_AREA.Text = ""
        IISSU_ADDR.Text = ""
        IISSU_TEL.Text = ""
        EMAIL.Text = ""
        Relationship.SelectedValue = ""
        Benefi.Text = "法定繼承人"
        Benefi_ID.Text = ""
        Relationship2.SelectedValue = ""

        Benefi2.Text = ""
        Benefi_ID2.Text = ""
        Relationship22.SelectedValue = ""

        Benefi3.Text = ""
        Benefi_ID3.Text = ""
        Relationship23.SelectedValue = ""

        txtTELREL.Text = ""
        txtZIPREL.Text = ""
        txtAREL.Text = ""
        txtDBIRTH.Text = ""
        txtISEQ.Text = ""
        txtIISSU_Ori.Text = ""

        PASSPORT_NAME.Text = ""
        PASSPORT_NUMBER.Text = ""

        '清空判斷[可保特殊方案(豪華型、標準型)]的變數，清空後會抓通路代碼重新判斷
        txtIsSPEChannel.Text = ""

        '250613回復清空，原改成不清空
        ddlIT41.SelectedValue = ""
        txtIT41.Text = ""
        ddlIT42.SelectedValue = ""
        txtIT42.Text = ""
        ddlIT43.SelectedValue = ""
        txtIT43.Text = ""
        ddlIT44.SelectedValue = ""
        txtIT44.Text = ""
        ddlIT4Y.SelectedValue = ""
        txtIT4Y.Text = ""
        ddlIT61.SelectedValue = ""
        txtIT61.Text = ""
        ddlIT63.SelectedValue = ""
        txtIT63.Text = ""
        ddlIT64.SelectedValue = ""
        txtIT64.Text = ""
        ddlIT6E.SelectedValue = ""
        txtIT6E.Text = ""
        lblIT65.Text = ""
        txtIT65.Text = ""
        lblIT66.Text = ""
        txtIT66.Text = ""
        lblIT6D.Text = ""
        txtIT6D.Text = ""
        lblIT6F.Text = ""
        txtIT6F.Text = ""
        lblIT6G.Text = ""
        txtIT6G.Text = ""
        CHK_P1.Checked = False
        CHK_P2.Checked = False
    End Sub

    '要、被保人國別下拉選單給值
    Private Sub InitCountryDDL()
        Dim dsA As DataSet = Me.InvokeCommonService("B2BService.bsB2BCommonService", "GetCOUNTRYCODE")
        For Each dr As DataRow In dsA.Tables("CMNP_COUNTRYCODE").Rows
            Dim b As New ListItem
            b.Text = dr("ICOUNTRY") + " " + dr("NCCOUNTRY")
            b.Value = dr("ICOUNTRY")
            ddlIISSU_COUNTRY.Items.Add(b)
            ddlIINS_COUNTRY.Items.Add(b)
        Next
    End Sub

    Function IsAllENG(str As String) As Boolean
        Return Regex.IsMatch(str, "^[A-Z'.\- ,]+$")
    End Function

    Function IsAllENGnINT(str As String) As Boolean
        Return Regex.IsMatch(str, "^[a-zA-Z0-9]+$")
    End Function

    ''' <summary>
    ''' 點擊方案按鈕時，將各細項方案的金額清空
    ''' </summary>
    ''' <param name="include65_66"></param>
    Private Sub ResetFields(include65_66 As Boolean, include6D As Boolean, include6F_6G As Boolean)

        ' 基本欄位（全部共用）
        Dim baseFields As String() = {
            "txtIT11", "txtIT27", "txtIT28",
            "txtIT2D", "txtIT2E", "txtIT2G", "txtIT2I",
            "txtIT2M", "txtIT2N", "txtIT2P", "txtIT2R",
            "txtIT2S", "txtIT2T", "txtIT2U",
            "txtIT31", "txtIT32", "txtIT33", "txtIT34", "txtIT35",
            "txtIT36", "txtIT37", "txtIT38", "txtIT39",
            "txtIT3A", "txtIT3B", "txtIT3C", "txtIT3D", "txtIT3E"
        }

        ' 設定為 0
        For Each txtid In baseFields
            Dim ctrl = CType(FindControl(txtid), TextBox)
            If ctrl IsNot Nothing Then ctrl.Text = "0"
        Next

        '' 條件欄位：2R / 31
        'If include2R_31 Then
        '    Dim extraFields As String() = {"txtIT2R", "txtIT31"}
        '    For Each txtid In extraFields
        '        Dim ctrl = CType(FindControl(txtid), TextBox)
        '        If ctrl IsNot Nothing Then ctrl.Text = "0"
        '    Next
        'End If

        ' 條件欄位：65 / 66（清空）
        If include65_66 Then
            Dim ctrl65 = CType(FindControl("txtIT65"), TextBox)
            Dim ctrl66 = CType(FindControl("txtIT66"), TextBox)

            If ctrl65 IsNot Nothing Then ctrl65.Text = ""
            If ctrl66 IsNot Nothing Then ctrl66.Text = ""
        End If

        ' 條件欄位：6D（清空）
        If include6D Then
            Dim ctrl6D = CType(FindControl("txtIT6D"), TextBox)

            If ctrl6D IsNot Nothing Then ctrl6D.Text = ""
        End If

        ' 條件欄位：6F / 6G（清空）
        If include6F_6G Then
            Dim ctrl6F = CType(FindControl("txtIT6F"), TextBox)
            Dim ctrl6G = CType(FindControl("txtIT6G"), TextBox)

            If ctrl6F IsNot Nothing Then ctrl6F.Text = ""
            If ctrl6G IsNot Nothing Then ctrl6G.Text = ""
        End If
    End Sub

    ''' <summary>
    ''' 點擊方案按鈕時，將海突含傳染相關顯示
    ''' </summary>
    Private Sub FieldsSwtich_Disease(DObj As DiseaseFields)
        If DObj.IT61 Then
            Me.ddlIT61.Enabled = True
        Else
            Me.ddlIT61.Enabled = False
            Me.ddlIT61.SelectedValue = ""
            Me.txtIT61.Text = ""
        End If

        If DObj.IT63 Then
            Me.titIT63.Attributes.Add("style", "display:inline")
            Me.ddlIT63.Attributes.Add("style", "display:inline")
            Me.ddlIT63.Enabled = True
        Else
            If Not DObj.IT6E Then
                Me.titIT63.Attributes.Add("style", "display:none")
            End If
            Me.ddlIT63.Attributes.Add("style", "display:none")
            Me.ddlIT63.Enabled = False
            Me.ddlIT63.SelectedValue = ""
            Me.txtIT63.Text = ""
        End If

        If DObj.IT64 Then
            Me.titIT64.Attributes.Add("style", "display:inline")
            Me.ddlIT64.Attributes.Add("style", "display:inline")
            Me.ddlIT64.Enabled = True
        Else
            Me.titIT64.Attributes.Add("style", "display:none")
            Me.ddlIT64.Attributes.Add("style", "display:none")
            Me.ddlIT64.Enabled = False
            Me.ddlIT64.SelectedValue = ""
            Me.txtIT64.Text = ""
        End If

        If DObj.IT6E Then
            Me.titIT63.Attributes.Add("style", "display:inline")
            Me.ddlIT6E.Attributes.Add("style", "display:inline")
            Me.ddlIT6E.Enabled = True
        Else
            If Not DObj.IT63 Then
                Me.titIT63.Attributes.Add("style", "display:none")
            End If
            Me.ddlIT6E.Attributes.Add("style", "display:none")
            Me.ddlIT6E.Enabled = False
            Me.ddlIT6E.SelectedValue = ""
            Me.txtIT6E.Text = ""
        End If

        If DObj.IT65 Then
            Me.titIT65.Attributes.Add("style", "display:inline")
            Me.lblIT65.Attributes.Add("style", "display:inline")
        Else
            Me.titIT65.Attributes.Add("style", "display:none")
            Me.lblIT65.Attributes.Add("style", "display:none")
            Me.lblIT65.Text = ""
            Me.txtIT65.Text = ""
        End If

        If DObj.IT66 Then
            Me.titIT66.Attributes.Add("style", "display:inline")
            Me.lblIT66.Attributes.Add("style", "display:inline")
        Else
            Me.titIT66.Attributes.Add("style", "display:none")
            Me.lblIT66.Attributes.Add("style", "display:none")
            Me.lblIT66.Text = ""
            Me.txtIT66.Text = ""
        End If

        If DObj.IT6F Then
            Me.titIT6F.Attributes.Add("style", "display:inline")
            Me.lblIT6F.Attributes.Add("style", "display:inline")
        Else
            Me.titIT6F.Attributes.Add("style", "display:none")
            Me.lblIT6F.Attributes.Add("style", "display:none")
            Me.lblIT6F.Text = ""
            Me.txtIT6F.Text = ""
        End If

        If DObj.IT6G Then
            Me.titIT6G.Attributes.Add("style", "display:inline")
            Me.lblIT6G.Attributes.Add("style", "display:inline")
        Else
            Me.titIT6G.Attributes.Add("style", "display:none")
            Me.lblIT6G.Attributes.Add("style", "display:none")
            Me.lblIT6G.Text = ""
            Me.txtIT6G.Text = ""
        End If

        If DObj.IT6D Then
            Me.titIT6D.Attributes.Add("style", "display:inline")
            Me.lblIT6D.Attributes.Add("style", "display:inline")
            Me.lblUnitIT6D.Attributes.Add("style", "display:inline")
        Else
            Me.titIT6D.Attributes.Add("style", "display:none")
            Me.lblIT6D.Attributes.Add("style", "display:none")
            Me.lblUnitIT6D.Attributes.Add("style", "display:none")
            Me.txtIT6D.Text = ""
        End If

    End Sub


    Public Sub GET_IISSU_Default()

        ddlIT41.SelectedValue = ""
        txtIT41.Text = ""
        ddlIT42.SelectedValue = ""
        txtIT42.Text = ""
        ddlIT43.SelectedValue = ""
        txtIT43.Text = ""
        ddlIT44.SelectedValue = ""
        txtIT44.Text = ""
        ddlIT4Y.SelectedValue = ""
        txtIT4Y.Text = ""
        ddlIT61.SelectedValue = ""
        ddlIT63.SelectedValue = ""
        txtIT61.Text = ""
        txtIT63.Text = ""
        ddlIT64.SelectedValue = ""
        txtIT64.Text = ""
        ddlIT6E.SelectedValue = ""
        txtIT6E.Text = ""
        lblIT65.Text = ""
        txtIT65.Text = ""
        lblIT66.Text = ""
        txtIT66.Text = ""
        lblIT6D.Text = ""
        txtIT6D.Text = ""
        lblIT6F.Text = ""
        txtIT6F.Text = ""
        lblIT6G.Text = ""
        txtIT6G.Text = ""
        'Label1.Text = ""
        'Label2.Text = ""
        'Label3.Text = ""
        'Label5.Text = ""
        'Label6.Text = ""
        'Label7.Text = ""
        'Label8.Text = ""
        'Label9.Text = ""
        'Label10.Text = ""
        'Label11.Text = ""

        txtIT11.Text = ""
        'txtIT21.Text = ""
        txtIT3A.Text = ""
        txtIT3B.Text = ""
        txtIT3C.Text = ""
        txtIT3D.Text = ""
        txtIT3E.Text = ""
        txtIT2N.Text = ""
        txtIT2M.Text = ""
        'txtIT2B.Text = ""
        txtIT27.Text = ""
        txtIT28.Text = ""
        txtIT31.Text = ""
        txtIT32.Text = ""
        txtIT33.Text = ""
        txtIT34.Text = ""
        txtIT35.Text = ""
        txtIT36.Text = ""
        txtIT37.Text = ""
        txtIT38.Text = ""
        txtIT39.Text = ""
        txtIT2D.Text = ""
        txtIT2P.Text = ""
        txtIT2S.Text = ""
        txtIT2R.Text = ""
        txtIT2U.Text = "0"
        txtIT2T.Text = "0"
        txtTotal.Text = ""

    End Sub
    Function GetiPlanNo(ByVal obj As Object) As String

        If obj Is Nothing Then
            Return ""
        Else
            If obj.ToString.Trim() = "" Then
                Return ""
            End If
            If obj.ToString.Trim() = "A" Then
                Return "A 國內海陸型"
            ElseIf obj.ToString.Trim() = "B" Then
                Return "B 國內飛翔型"
            ElseIf obj.ToString.Trim() = "C" Then
                Return "C 海外基本型"
            ElseIf obj.ToString.Trim() = "D" Then
                Return "D 海外超值型"
            ElseIf obj.ToString.Trim() = "J" Then
                Return "J 海外標準型"
            ElseIf obj.ToString.Trim() = "K" Then
                Return "K 海外豪華型"
            ElseIf obj.ToString.Trim() = "L" Then
                Return "L 一般型或申根型"
            ElseIf obj.ToString.Trim() = "M" Then
                Return "M 海外安心型"
            ElseIf obj.ToString.Trim() = "N" Then
                Return "N 海外特選型"
            ElseIf obj.ToString.Trim() = "O" Then
                Return "O 海外小資型"
            End If

        End If
    End Function

    Private Function divfuncinfoDisplay(ByVal ShowOrNot As Boolean)
        If ShowOrNot Then
            divfuncinfo.Attributes.Add("style", "display:inhert")
        Else
            divfuncinfo.Attributes.Add("style", "display:none")
        End If
    End Function

    Function GetTYPEREL(ByVal obj As Object) As String
        Return CommonFunc.Relationship_to_Name(obj)
    End Function
    Function GetTYPEREL2(ByVal obj As Object) As String

        If obj Is Nothing Then
            Return ""
        Else
            If obj.ToString.Trim() = "" Then
                Return ""
            End If
            If obj.ToString.Trim() = "1" Then
                Return "1 配偶"
            ElseIf obj.ToString.Trim() = "2" Then
                Return "2 子女"
            ElseIf obj.ToString.Trim() = "3" Then
                Return "3 父母"
            ElseIf obj.ToString.Trim() = "4" Then
                Return "4 兄弟姐妹"
            ElseIf obj.ToString.Trim() = "5" Then
                Return "5 費用仰給人"
            ElseIf obj.ToString.Trim() = "6" Then
                Return "6 債務人"
            ElseIf obj.ToString.Trim() = "7" Then
                Return "7 財產管理人"
            ElseIf obj.ToString.Trim() = "9" Then
                Return "9 祖父母"
            ElseIf obj.ToString.Trim() = "A" Then
                Return "A 孫子女"
            End If

        End If
    End Function
    Public Sub InsertIQPOBJ(iplanno As String, type As String)
        Dim ds As beITP101000 = CType(Me.StateMain, beITP101000)
        Dim TargetProject As String = ddlTargetProject.SelectedValue.Trim

        'ITP 不便險取保費的方式 
        'IT41/IT4Y/IT42/IT62 這四檔 旅平的相關險種
        '1:InsType(險種)
        '2:InsDay(天數)
        '3.InsCoverage 保額 元
        '4.InsMdeduct 自負額 元
        '5.CountryType  代表費率別 ：3=個人費率、4=團費率
        'Dflag(要抓該費率檔內小於系統日的最大日)
        Dim CountryType As String = ""
        Dim InsMdeduct As String = ""
        Dim IT2RFree = False
        If CHK5.Checked = True Then
            CountryType = "4"
        Else
            CountryType = "3"
        End If

        If txtIT41.Text.Trim <> "" Then
            InsMdeduct = "0"
            Dim dr1 As beITP101000.IQPOBJRow
            dr1 = ds.IQPOBJ.NewIQPOBJRow
            dr1("IISSU") = txtIins_ISSU2.Text
            dr1("INSTyp") = "IT41"
            dr1("CONTENT") = "意外身故失能保險金"
            dr1("INSAMT") = ddlIT41.SelectedValue + "0000"
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney1 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr1("INSTyp").ToString, txtIQPITP_Days.Text, dr1("INSAMT"), InsMdeduct, CountryType)
            If dsMoney1.Tables("ITPPrem").Rows.Count > 0 Then
                dr1("INSPREM") = CInt(dsMoney1.Tables("ITPPrem").Rows(0)("PREM"))

                '20251218
                '	身故500萬以上，方案J K優惠IT2R不算錢
                If iplanno = "J" Or iplanno = "K" Then
                    If CType(dr1("INSAMT"), Int32) >= 5000000 Then IT2RFree = True
                End If
            Else
                dr1("INSPREM") = 0
            End If
            '================================
            'dr1("INSPREM") = txtIT41.Text
            If TargetProject = "IQP" Then dr1("INSTyp") = "IQP01"

            ds.IQPOBJ.AddIQPOBJRow(dr1)
            BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
        End If
        'Case 2
        If txtIT4Y.Text.Trim <> "" And txtIT4Y.Text.Trim <> "0" Then
            InsMdeduct = "0"
            Dim dr2 As beITP101000.IQPOBJRow
            dr2 = ds.IQPOBJ.NewIQPOBJRow
            dr2("IISSU") = txtIins_ISSU2.Text
            dr2("INSTyp") = "IT4Y"
            dr2("CONTENT") = "意外喪葬費用"
            dr2("INSAMT") = CInt(CDec(ddlIT4Y.SelectedValue) * 10000)
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney2 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr2("INSTyp").ToString, txtIQPITP_Days.Text, dr2("INSAMT").ToString, InsMdeduct, CountryType)
            If dsMoney2.Tables("ITPPrem").Rows.Count > 0 Then
                dr2("INSPREM") = CInt(dsMoney2.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr2("INSPREM") = 0
            End If
            '================================
            'dr2("INSPREM") = txtIT4Y.Text
            If TargetProject = "IQP" Then dr2("INSTyp") = "IQP0Y" '未滿15歲

            ds.IQPOBJ.AddIQPOBJRow(dr2)
            BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
        End If
        'Case 3
        If txtIT42.Text.Trim <> "" Then
            InsMdeduct = "0"
            Dim dr3 As beITP101000.IQPOBJRow
            dr3 = ds.IQPOBJ.NewIQPOBJRow
            dr3("IISSU") = txtIins_ISSU2.Text
            dr3("INSTyp") = "IT42"
            dr3("CONTENT") = "傷害醫療保險金"
            dr3("INSAMT") = ddlIT42.SelectedValue + "0000"
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney3 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr3("INSTyp").ToString, txtIQPITP_Days.Text, dr3("INSAMT"), InsMdeduct, CountryType)
            If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                dr3("INSPREM") = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr3("INSPREM") = 0
            End If
            '================================
            'dr3.Item("INSPREM") = txtIT42.Text
            If TargetProject = "IQP" Then dr3("INSTyp") = "IQP02"

            ds.IQPOBJ.AddIQPOBJRow(dr3)
            BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
        End If

        If txtIT43.Text.Trim <> "" Then
            InsMdeduct = "0"
            Dim dr43 As beITP101000.IQPOBJRow
            dr43 = ds.IQPOBJ.NewIQPOBJRow
            dr43("IISSU") = txtIins_ISSU2.Text
            dr43("INSTyp") = "IT43"
            dr43("CONTENT") = "大眾運輸保險金"
            dr43("INSAMT") = ddlIT43.SelectedValue + "0000"
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney43 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr43("INSTyp").ToString, txtIQPITP_Days.Text, dr43("INSAMT"), InsMdeduct, CountryType)
            If dsMoney43.Tables("ITPPrem").Rows.Count > 0 Then
                dr43("INSPREM") = CInt(dsMoney43.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr43("INSPREM") = 0
            End If
            '================================
            'dr3.Item("INSPREM") = txtIT42.Text
            If TargetProject = "IQP" Then dr43("INSTyp") = "IQP03"

            ds.IQPOBJ.AddIQPOBJRow(dr43)
            BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
        End If
        'Case 3.5
        If txtIT44.Text.Trim <> "" Then
            InsMdeduct = "0"
            Dim dr33 As beITP101000.IQPOBJRow
            dr33 = ds.IQPOBJ.NewIQPOBJRow
            dr33("IISSU") = txtIins_ISSU2.Text
            dr33("INSTyp") = "IT44"
            dr33("CONTENT") = "重大燒燙傷保險金"
            dr33("INSAMT") = ddlIT44.SelectedValue + "0000"
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney4 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr33("INSTyp").ToString, txtIQPITP_Days.Text, dr33("INSAMT"), InsMdeduct, CountryType)
            If dsMoney4.Tables("ITPPrem").Rows.Count > 0 Then
                dr33("INSPREM") = CInt(dsMoney4.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr33("INSPREM") = 0
            End If
            '================================
            'dr33("INSPREM") = txtIT44.Text
            If TargetProject = "IQP" Then dr33("INSTyp") = "IQP04"

            ds.IQPOBJ.AddIQPOBJRow(dr33)
            BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
        End If
        'Case 4
        If txtIT61.Text.Trim <> "" Then
            InsMdeduct = "0"
            Dim dr4 As beITP101000.IQPOBJRow
            dr4 = ds.IQPOBJ.NewIQPOBJRow
            dr4("IISSU") = txtIins_ISSU2.Text
            dr4("INSTyp") = "IT62"
            dr4("CONTENT") = "海外突發疾病醫療保險金(不含傳染病)"
            dr4("INSAMT") = ddlIT61.SelectedValue + "0000"
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney4 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr4("INSTyp").ToString, txtIQPITP_Days.Text, dr4("INSAMT"), InsMdeduct, CountryType)
            If dsMoney4.Tables("ITPPrem").Rows.Count > 0 Then
                dr4("INSPREM") = CInt(dsMoney4.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr4("INSPREM") = 0
            End If
            '================================

            ' 若IQP存檔時改回， 因費率表(ITPPrem_new)僅記代碼IT62;不記IQP11
            If TargetProject = "IQP" Then
                dr4("INSTyp") = "IQP11"
            End If
            'dr4("INSPREM") = txtIT61.Text

            'dr4("INSTyp") = "IQP11"
            'If TargetProject = "IQP" Then ds.IQPOBJ.AddIQPOBJRow(dr4)
            ds.IQPOBJ.AddIQPOBJRow(dr4)

            BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
        End If
        If txtIT63.Text.Trim <> "" Then
            InsMdeduct = "0"
            Dim dr4 As beITP101000.IQPOBJRow
            dr4 = ds.IQPOBJ.NewIQPOBJRow
            dr4("IISSU") = txtIins_ISSU2.Text
            dr4("INSTyp") = "IT63"
            dr4("CONTENT") = "海外突發疾病醫療保險金(含傳染病)"
            dr4("INSAMT") = ddlIT63.SelectedValue + "0000"
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney4 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr4("INSTyp").ToString, txtIQPITP_Days.Text, dr4("INSAMT"), InsMdeduct, CountryType)
            If dsMoney4.Tables("ITPPrem").Rows.Count > 0 Then
                dr4("INSPREM") = CInt(dsMoney4.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr4("INSPREM") = 0
            End If
            '================================

            ' 若IQP存檔時改回， 因費率表(ITPPrem_new)僅記代碼IT63;不記IQP12
            If TargetProject = "IQP" Then
                dr4("INSTyp") = "IQP12"
            End If

            ds.IQPOBJ.AddIQPOBJRow(dr4)

            'IT6D
            'F115060600 新增IT6D，方案基本型 & 超值型才有的附加保險金，而J、K、L無IT63，不用特別判斷
            If TargetProject = "ITP" Then
                Dim dr6D As beITP101000.IQPOBJRow
                dr6D = ds.IQPOBJ.NewIQPOBJRow
                dr6D("IISSU") = txtIins_ISSU2.Text
                dr6D("INSTyp") = "IT6D"
                dr6D("CONTENT") = "返國繼續住院醫療保險金"
                dr6D("INSAMT") = (CInt(ddlIT63.SelectedValue) * 10000 * 0.05).ToString
                '金額要重算 免得有人用滾輪亂滾
                '================================
                Dim dsMoney6D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr6D("INSTyp").ToString, txtIQPITP_Days.Text, dr6D("INSAMT"), InsMdeduct, CountryType)
                If dsMoney6D.Tables("ITPPrem").Rows.Count > 0 Then
                    dr6D("INSPREM") = CInt(dsMoney6D.Tables("ITPPrem").Rows(0)("PREM"))
                Else
                    dr6D("INSPREM") = 0
                End If
                '================================

                ds.IQPOBJ.AddIQPOBJRow(dr6D)
            End If

            BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
        End If
        If txtIT64.Text.Trim <> "" AndAlso txtIT64.Text.Trim <> "0" Then
            InsMdeduct = "0"
            '20251218 海突新增IT64、IT65、IT66
            'IT64
            Dim dr64 As beITP101000.IQPOBJRow
            dr64 = ds.IQPOBJ.NewIQPOBJRow
            dr64("IISSU") = txtIins_ISSU2.Text
            dr64("INSTyp") = "IT64"
            dr64("CONTENT") = "海外突發疾病住院醫療保險金"
            dr64("INSAMT") = ddlIT64.SelectedValue + "0000"
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney64 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr64("INSTyp").ToString, txtIQPITP_Days.Text, dr64("INSAMT"), InsMdeduct, CountryType)
            If dsMoney64.Tables("ITPPrem").Rows.Count > 0 Then
                dr64("INSPREM") = CInt(dsMoney64.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr64("INSPREM") = 0
            End If
            '================================

            ' 若IQP存檔時改回， 因費率表(ITPPrem_new)僅記代碼IT64;不記IQP13
            If TargetProject = "IQP" Then
                dr64("INSTyp") = "IQP13"
            End If

            ds.IQPOBJ.AddIQPOBJRow(dr64)

            'IT65
            Dim dr65 As beITP101000.IQPOBJRow
            dr65 = ds.IQPOBJ.NewIQPOBJRow
            dr65("IISSU") = txtIins_ISSU2.Text
            dr65("INSTyp") = "IT65"
            dr65("CONTENT") = "海外突發疾病門診醫療保險金"
            dr65("INSAMT") = (CInt(ddlIT64.SelectedValue) * 10000 * 0.05).ToString
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney65 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr65("INSTyp").ToString, txtIQPITP_Days.Text, dr65("INSAMT"), InsMdeduct, CountryType)
            If dsMoney65.Tables("ITPPrem").Rows.Count > 0 Then
                dr65("INSPREM") = CInt(dsMoney65.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr65("INSPREM") = 0
            End If
            '================================

            ' 若IQP存檔時改回， 因費率表(ITPPrem_new)僅記代碼IT65;不記IQP14
            If TargetProject = "IQP" Then
                dr65("INSTyp") = "IQP14"
            End If

            ds.IQPOBJ.AddIQPOBJRow(dr65)

            'IT66
            Dim dr66 As beITP101000.IQPOBJRow
            dr66 = ds.IQPOBJ.NewIQPOBJRow
            dr66("IISSU") = txtIins_ISSU2.Text
            dr66("INSTyp") = "IT66"
            dr66("CONTENT") = "海外突發疾病急診醫療保險金"
            dr66("INSAMT") = (CInt(ddlIT64.SelectedValue) * 10000 * 0.05).ToString
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney66 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr66("INSTyp").ToString, txtIQPITP_Days.Text, dr66("INSAMT"), InsMdeduct, CountryType)
            If dsMoney66.Tables("ITPPrem").Rows.Count > 0 Then
                dr66("INSPREM") = CInt(dsMoney66.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr66("INSPREM") = 0
            End If
            '================================

            ' 若IQP存檔時改回， 因費率表(ITPPrem_new)僅記代碼IT66;不記IQP15
            If TargetProject = "IQP" Then
                dr66("INSTyp") = "IQP15"
            End If

            ds.IQPOBJ.AddIQPOBJRow(dr66)

            If checkChannelMN().Item1 Or (checkChannelO().Item1 AndAlso iplanno = "O") Then
                'IT6D
                'F115060600 新增K077附加IT6D
                'F115082802 新增O 海外小資型附加IT6D
                If TargetProject = "ITP" Then
                    Dim dr6D As beITP101000.IQPOBJRow
                    dr6D = ds.IQPOBJ.NewIQPOBJRow
                    dr6D("IISSU") = txtIins_ISSU2.Text
                    dr6D("INSTyp") = "IT6D"
                    dr6D("CONTENT") = "返國繼續住院醫療保險金"
                    dr6D("INSAMT") = (CInt(ddlIT64.SelectedValue) * 10000 * 0.05).ToString
                    '金額要重算 免得有人用滾輪亂滾
                    '================================
                    Dim dsMoney6D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr6D("INSTyp").ToString, txtIQPITP_Days.Text, dr6D("INSAMT"), InsMdeduct, CountryType)
                    If dsMoney6D.Tables("ITPPrem").Rows.Count > 0 Then
                        dr6D("INSPREM") = CInt(dsMoney6D.Tables("ITPPrem").Rows(0)("PREM"))
                    Else
                        dr6D("INSPREM") = 0
                    End If
                    '================================

                    ds.IQPOBJ.AddIQPOBJRow(dr6D)
                End If
            End If

            BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
        End If
        If txtIT6E.Text.Trim <> "" And TargetProject = "ITP" Then
            InsMdeduct = "0"
            Dim dr6E As beITP101000.IQPOBJRow
            dr6E = ds.IQPOBJ.NewIQPOBJRow
            dr6E("IISSU") = txtIins_ISSU2.Text
            dr6E("INSTyp") = "IT6E"
            dr6E("CONTENT") = "海外突發疾病醫療保險金(含傳染病)"
            dr6E("INSAMT") = ddlIT6E.SelectedValue + "0000"
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney4 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr6E("INSTyp").ToString, txtIQPITP_Days.Text, dr6E("INSAMT"), InsMdeduct, CountryType)
            If dsMoney4.Tables("ITPPrem").Rows.Count > 0 Then
                dr6E("INSPREM") = CInt(dsMoney4.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr6E("INSPREM") = 0
            End If
            '================================

            ds.IQPOBJ.AddIQPOBJRow(dr6E)

            'IT6F
            Dim dr6F As beITP101000.IQPOBJRow
            dr6F = ds.IQPOBJ.NewIQPOBJRow
            dr6F("IISSU") = txtIins_ISSU2.Text
            dr6F("INSTyp") = "IT6F"
            dr6F("CONTENT") = "海外突發疾病急診醫療保險金"
            dr6F("INSAMT") = (CInt(ddlIT6E.SelectedValue) * 10000 * 0.01).ToString
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney6F As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr6F("INSTyp").ToString, txtIQPITP_Days.Text, dr6F("INSAMT"), InsMdeduct, CountryType)
            If dsMoney6F.Tables("ITPPrem").Rows.Count > 0 Then
                dr6F("INSPREM") = CInt(dsMoney6F.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr6F("INSPREM") = 0
            End If
            '================================

            ds.IQPOBJ.AddIQPOBJRow(dr6F)

            'IT6G
            Dim dr6G As beITP101000.IQPOBJRow
            dr6G = ds.IQPOBJ.NewIQPOBJRow
            dr6G("IISSU") = txtIins_ISSU2.Text
            dr6G("INSTyp") = "IT6G"
            dr6G("CONTENT") = "海外突發疾病門診醫療保險金"
            dr6G("INSAMT") = (CInt(ddlIT6E.SelectedValue) * 10000 * 0.01).ToString
            '金額要重算 免得有人用滾輪亂滾
            '================================
            Dim dsMoney66 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr6G("INSTyp").ToString, txtIQPITP_Days.Text, dr6G("INSAMT"), InsMdeduct, CountryType)
            If dsMoney66.Tables("ITPPrem").Rows.Count > 0 Then
                dr6G("INSPREM") = CInt(dsMoney66.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                dr6G("INSPREM") = 0
            End If
            '================================

            ds.IQPOBJ.AddIQPOBJRow(dr6G)

            'IT6D
            'F115060600 新增IT6D，方案基本型 & 超值型才有的附加保險金，而J、K、L無IT6E，不用特別判斷
            If TargetProject = "ITP" Then
                Dim dr6D As beITP101000.IQPOBJRow
                dr6D = ds.IQPOBJ.NewIQPOBJRow
                dr6D("IISSU") = txtIins_ISSU2.Text
                dr6D("INSTyp") = "IT6D"
                dr6D("CONTENT") = "返國繼續住院醫療保險金"
                dr6D("INSAMT") = (CInt(ddlIT6E.SelectedValue) * 10000 * 0.05).ToString
                '金額要重算 免得有人用滾輪亂滾
                '================================
                Dim dsMoney6D As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", dr6D("INSTyp").ToString, txtIQPITP_Days.Text, dr6D("INSAMT"), InsMdeduct, CountryType)
                If dsMoney6D.Tables("ITPPrem").Rows.Count > 0 Then
                    dr6D("INSPREM") = CInt(dsMoney6D.Tables("ITPPrem").Rows(0)("PREM"))
                Else
                    dr6D("INSPREM") = 0
                End If
                '================================

                ds.IQPOBJ.AddIQPOBJRow(dr6D)
            End If

            BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
        End If
        '20250320 只有ITP有方案
        If ddlTargetProject.SelectedValue = "ITP" And iplanno <> "L" Then
            '險種代號       海外基本型		海外超值型
            'IT11		    限額200000元	限額200000元
            'IT2R		    限額100萬		限額300萬
            'IT37(原IT2K)   限額10000元		限額30000元
            'IT36(原IT2J)   限額20000元		限額60000元 
            'IT3A(原IT2A)   每次4000元		每次6000元
            'IT3B(原IT2L)   每次3000元		每次3000元
            'IT3C           每次10000元		每次30000元
            'IT3D           每次1000元		每次2000元
            'IT3E           每次1000元		每次2000元
            'IT39(原IT29)   每次4000元		每次6000元
            'IT38(原IT2X)	每次4000元		每次6000元
            'IT2M		    每次3000元		每次6000元
            'IT27		    每日3000元		每日5000元
            'IT2D		    限額20000元		限額30000元
            'IT28		    每次3000元		每次5000元
            'IT31		    限額50000元		限額50000元

            '險種代號       海外標準型		海外豪華型
            'IT11		    限額3000000元	限額5000000元
            'IT2R		    限額250萬		限額250萬
            'IT3B(原IT2L)   每次3000元		每次3000元
            'IT3A(原IT2A)   每次4000元		每次6000元
            'IT39(原IT29)   每次4000元		每次6000元
            'IT38(原IT2X)   每次4000元		每次6000元
            'IT2M		    每次3000元		每次5000元
            'IT27		    每日3000元		每日6000元
            'IT2D		    限額30000元		限額50000元
            'IT28		    每次3000元		每次3000元
            'IT31		    限額30000元		限額50000元
            'IT36(原IT32)   限額120000元	限額150000元
            'IT37(原IT33)   限額120000元	限額150000元
            'IT34		    限額120000元	限額150000元
            'IT35		    限額120000元	限額150000元
            'IT2P		    限額3000元		限額5000元
            'IT2T		    每日10000元		每日10000元
            'IT2U		    每次3000元		每次5000元
            'IT2S		    每次3000元		每次5000元

            '險種代號       O 海外小資型	
            'IT11		    限額200000元		
            'IT36(原IT2J)   限額60000元		
            'IT37(原IT2K)   限額30000元
            'IT38(原IT2X)	每次5000元
            'IT39(原IT29)   每次5000元
            'IT3A(原IT2A)   每次5000元		
            'IT3B(原IT2L)   每次3000元	
            'IT2M		    每次5000元		
            'IT28		    每次3000元	
            'IT31		    限額50000元	
            'IT2D		    限額50000元	
            'IT3D           每次1000元	

            'IT11
            If (txtIT11.Text <> "0" And txtIT11.Text <> "") Then
                Dim dr5 As beITP101000.IQPOBJRow
                dr5 = ds.IQPOBJ.NewIQPOBJRow
                dr5("IISSU") = txtIins_ISSU2.Text
                dr5("INSTyp") = "IT11"
                dr5("INSPREM") = txtIT11.Text
                'MDEDUCT 自負額
                If iplanno = "A" Or iplanno = "B" Then
                    dr5("INSAMT") = "50000"
                ElseIf iplanno = "J" Then
                    dr5("INSAMT") = "3000000"
                ElseIf iplanno = "K" Then
                    dr5("INSAMT") = "5000000"
                Else
                    dr5("INSAMT") = "200000"
                End If

                dr5("MDEDUCT") = "2500"
                ds.IQPOBJ.AddIQPOBJRow(dr5)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If

            'IT2R
            If (txtIT2R.Text <> "0" And txtIT2R.Text <> "") Then
                Dim dr As beITP101000.IQPOBJRow
                dr = ds.IQPOBJ.NewIQPOBJRow
                dr("IISSU") = txtIins_ISSU2.Text
                dr("INSTyp") = "IT2R"
                If iplanno = "C" Or iplanno = "M" Then
                    dr("INSAMT") = "1000000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr("INSAMT") = "3000000"
                ElseIf iplanno = "J" Or iplanno = "K" Then
                    dr("INSAMT") = "2500000"
                Else
                    dr("INSAMT") = "0"
                End If

                If IT2RFree Then
                    dr("INSPREM") = "0"
                Else
                    dr("INSPREM") = txtIT2R.Text
                End If

                ds.IQPOBJ.AddIQPOBJRow(dr)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT37
            If (txtIT37.Text <> "0" And txtIT37.Text <> "") Then
                Dim dr As beITP101000.IQPOBJRow
                dr = ds.IQPOBJ.NewIQPOBJRow
                dr("IISSU") = txtIins_ISSU2.Text
                dr("INSTyp") = "IT37"
                If iplanno = "C" Or iplanno = "M" Then
                    dr("INSAMT") = "10000"
                ElseIf iplanno = "D" Or iplanno = "N" Or iplanno = "O" Then
                    dr("INSAMT") = "30000"
                ElseIf iplanno = "J" Then
                    dr("INSAMT") = "120000"
                ElseIf iplanno = "K" Then
                    dr("INSAMT") = "150000"
                Else
                    dr("INSAMT") = "0"
                End If
                dr("INSPREM") = txtIT37.Text
                ds.IQPOBJ.AddIQPOBJRow(dr)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT36
            If (txtIT36.Text <> "0" And txtIT36.Text <> "") Then
                Dim dr As beITP101000.IQPOBJRow
                dr = ds.IQPOBJ.NewIQPOBJRow
                dr("IISSU") = txtIins_ISSU2.Text
                dr("INSTyp") = "IT36"
                If iplanno = "C" Or iplanno = "M" Then
                    dr("INSAMT") = "20000"
                ElseIf iplanno = "D" Or iplanno = "N" Or iplanno = "O" Then
                    dr("INSAMT") = "60000"
                ElseIf iplanno = "J" Then
                    dr("INSAMT") = "120000"
                ElseIf iplanno = "K" Then
                    dr("INSAMT") = "150000"
                Else
                    dr("INSAMT") = "0"
                End If
                dr("INSPREM") = txtIT36.Text
                ds.IQPOBJ.AddIQPOBJRow(dr)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT3A  
            If (txtIT3A.Text <> "0" And txtIT3A.Text <> "") Then
                Dim dr2A As beITP101000.IQPOBJRow
                dr2A = ds.IQPOBJ.NewIQPOBJRow
                dr2A("IISSU") = txtIins_ISSU2.Text
                dr2A("INSTyp") = "IT3A"
                If iplanno = "A" Then
                    dr2A("INSAMT") = "0"
                ElseIf iplanno = "B" Then
                    dr2A("INSAMT") = "1000"
                ElseIf iplanno = "C" Or iplanno = "M" Then
                    dr2A("INSAMT") = "4000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr2A("INSAMT") = "6000"
                ElseIf iplanno = "J" Then
                    dr2A("INSAMT") = "4000"
                ElseIf iplanno = "K" Then
                    dr2A("INSAMT") = "6000"
                ElseIf iplanno = "O" Then
                    dr2A("INSAMT") = "5000"
                End If
                dr2A("INSPREM") = txtIT3A.Text
                ds.IQPOBJ.AddIQPOBJRow(dr2A)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT3B
            If (txtIT3B.Text <> "0" And txtIT3B.Text <> "") Then
                Dim drL As beITP101000.IQPOBJRow
                drL = ds.IQPOBJ.NewIQPOBJRow
                drL("IISSU") = txtIins_ISSU2.Text
                drL("INSTyp") = "IT3B"
                If iplanno = "C" Or iplanno = "D" Or iplanno = "J" Or iplanno = "K" Or iplanno = "M" Or iplanno = "N" Or iplanno = "O" Then
                    drL("INSAMT") = "3000"
                Else
                    drL("INSAMT") = "0"
                End If
                drL("INSPREM") = txtIT3B.Text
                ds.IQPOBJ.AddIQPOBJRow(drL)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT3C
            If (txtIT3C.Text <> "0" And txtIT3C.Text <> "") Then
                Dim dr As beITP101000.IQPOBJRow
                dr = ds.IQPOBJ.NewIQPOBJRow
                dr("IISSU") = txtIins_ISSU2.Text
                dr("INSTyp") = "IT3C"
                If iplanno = "C" Or iplanno = "M" Then
                    dr("INSAMT") = "10000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr("INSAMT") = "30000"
                Else
                    dr("INSAMT") = "0"
                End If
                dr("INSPREM") = txtIT3C.Text
                ds.IQPOBJ.AddIQPOBJRow(dr)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT3D
            If (txtIT3D.Text <> "0" And txtIT3D.Text <> "") Then
                Dim dr As beITP101000.IQPOBJRow
                dr = ds.IQPOBJ.NewIQPOBJRow
                dr("IISSU") = txtIins_ISSU2.Text
                dr("INSTyp") = "IT3D"
                If iplanno = "C" Or iplanno = "M" Or iplanno = "O" Then
                    dr("INSAMT") = "1000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr("INSAMT") = "2000"
                Else
                    dr("INSAMT") = "0"
                End If
                dr("INSPREM") = txtIT3D.Text
                ds.IQPOBJ.AddIQPOBJRow(dr)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT3E
            If (txtIT3E.Text <> "0" And txtIT3E.Text <> "") Then
                Dim dr As beITP101000.IQPOBJRow
                dr = ds.IQPOBJ.NewIQPOBJRow
                dr("IISSU") = txtIins_ISSU2.Text
                dr("INSTyp") = "IT3E"
                If iplanno = "C" Or iplanno = "M" Then
                    dr("INSAMT") = "1000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr("INSAMT") = "2000"
                Else
                    dr("INSAMT") = "0"
                End If
                dr("INSPREM") = txtIT3E.Text
                ds.IQPOBJ.AddIQPOBJRow(dr)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If

            'IT39
            If (txtIT39.Text <> "0" And txtIT39.Text <> "") Then
                Dim dr As beITP101000.IQPOBJRow
                dr = ds.IQPOBJ.NewIQPOBJRow
                dr("IISSU") = txtIins_ISSU2.Text
                dr("INSTyp") = "IT39"
                If iplanno = "C" Or iplanno = "M" Then
                    dr("INSAMT") = "4000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr("INSAMT") = "6000"
                ElseIf iplanno = "J" Then
                    dr("INSAMT") = "4000"
                ElseIf iplanno = "K" Then
                    dr("INSAMT") = "6000"
                ElseIf iplanno = "O" Then
                    dr("INSAMT") = "5000"
                Else
                    dr("INSAMT") = "0"
                End If
                dr("INSPREM") = txtIT39.Text
                ds.IQPOBJ.AddIQPOBJRow(dr)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT38
            If (txtIT38.Text <> "0" And txtIT38.Text <> "") Then
                Dim dr9 As beITP101000.IQPOBJRow
                dr9 = ds.IQPOBJ.NewIQPOBJRow
                dr9("IISSU") = txtIins_ISSU2.Text
                dr9("INSTyp") = "IT38"
                If iplanno = "A" Then
                    dr9("INSAMT") = "0"
                ElseIf iplanno = "B" Then
                    dr9("INSAMT") = "1500"
                ElseIf iplanno = "C" Or iplanno = "M" Then
                    dr9("INSAMT") = "4000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr9("INSAMT") = "6000"
                ElseIf iplanno = "J" Then
                    dr9("INSAMT") = "4000"
                ElseIf iplanno = "K" Then
                    dr9("INSAMT") = "6000"
                ElseIf iplanno = "O" Then
                    dr9("INSAMT") = "5000"
                End If
                dr9("INSPREM") = txtIT38.Text
                ds.IQPOBJ.AddIQPOBJRow(dr9)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If

            'IT2M
            If (txtIT2M.Text <> "0" And txtIT2M.Text <> "") Then
                Dim dr As beITP101000.IQPOBJRow
                dr = ds.IQPOBJ.NewIQPOBJRow
                dr("IISSU") = txtIins_ISSU2.Text
                dr("INSTyp") = "IT2M"
                If iplanno = "C" Or iplanno = "M" Then
                    dr("INSAMT") = "3000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr("INSAMT") = "6000"
                ElseIf iplanno = "J" Then
                    dr("INSAMT") = "3000"
                ElseIf iplanno = "K" Or iplanno = "O" Then
                    dr("INSAMT") = "5000"
                Else
                    dr("INSAMT") = "0"
                End If
                dr("INSPREM") = txtIT2M.Text
                ds.IQPOBJ.AddIQPOBJRow(dr)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If

            'IT27
            If (txtIT27.Text <> "0" And txtIT27.Text <> "") Then
                Dim dr12 As beITP101000.IQPOBJRow
                dr12 = ds.IQPOBJ.NewIQPOBJRow
                dr12("IISSU") = txtIins_ISSU2.Text
                dr12("INSTyp") = "IT27"
                If iplanno = "A" Then
                    dr12("INSAMT") = "0"
                ElseIf iplanno = "B" Then
                    dr12("INSAMT") = "2000"
                ElseIf iplanno = "C" Or iplanno = "M" Then
                    dr12("INSAMT") = "3000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr12("INSAMT") = "5000"
                ElseIf iplanno = "J" Then
                    dr12("INSAMT") = "3000"
                ElseIf iplanno = "K" Then
                    dr12("INSAMT") = "6000"
                End If
                dr12("INSPREM") = txtIT27.Text
                ds.IQPOBJ.AddIQPOBJRow(dr12)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT2D
            If (txtIT2D.Text <> "0" And txtIT2D.Text <> "") Then
                Dim dr2D As beITP101000.IQPOBJRow
                dr2D = ds.IQPOBJ.NewIQPOBJRow
                dr2D("IISSU") = txtIins_ISSU2.Text
                dr2D("INSTyp") = "IT2D"
                If iplanno = "C" Or iplanno = "M" Then
                    dr2D("INSAMT") = "20000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr2D("INSAMT") = "30000"
                ElseIf iplanno = "J" Then
                    dr2D("INSAMT") = "30000"
                ElseIf iplanno = "K" Or iplanno = "O" Then
                    dr2D("INSAMT") = "50000"
                End If
                dr2D("INSPREM") = txtIT2D.Text
                ds.IQPOBJ.AddIQPOBJRow(dr2D)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT28
            If (txtIT28.Text <> "0" And txtIT28.Text <> "") Then
                Dim dr13 As beITP101000.IQPOBJRow
                dr13 = ds.IQPOBJ.NewIQPOBJRow
                dr13("IISSU") = txtIins_ISSU2.Text
                dr13("INSTyp") = "IT28"
                If iplanno = "A" Or iplanno = "B" Then
                    dr13("INSAMT") = "3000"
                ElseIf iplanno = "C" Or iplanno = "M" Then
                    dr13("INSAMT") = "3000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr13("INSAMT") = "5000"
                ElseIf iplanno = "J" Then
                    dr13("INSAMT") = "3000"
                ElseIf iplanno = "K" Or iplanno = "O" Then
                    dr13("INSAMT") = "3000"
                End If
                dr13("INSPREM") = txtIT28.Text
                ds.IQPOBJ.AddIQPOBJRow(dr13)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT31
            If (txtIT31.Text <> "0" And txtIT31.Text <> "") Then

                Dim dr As beITP101000.IQPOBJRow
                dr = ds.IQPOBJ.NewIQPOBJRow
                dr("IISSU") = txtIins_ISSU2.Text
                dr("INSTyp") = "IT31"
                If iplanno = "C" Or iplanno = "M" Or iplanno = "O" Then
                    dr("INSAMT") = "50000"
                ElseIf iplanno = "D" Or iplanno = "N" Then
                    dr("INSAMT") = "50000"
                ElseIf iplanno = "J" Then
                    dr("INSAMT") = "30000"
                ElseIf iplanno = "K" Then
                    dr("INSAMT") = "50000"
                Else
                    dr("INSAMT") = "0"
                End If
                dr("INSPREM") = txtIT31.Text
                ds.IQPOBJ.AddIQPOBJRow(dr)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT34
            If (txtIT34.Text <> "0" And txtIT34.Text <> "") Then

                Dim dr34 As beITP101000.IQPOBJRow
                dr34 = ds.IQPOBJ.NewIQPOBJRow
                dr34("IISSU") = txtIins_ISSU2.Text
                dr34("INSTyp") = "IT34"
                If iplanno = "J" Then
                    dr34("INSAMT") = "120000"
                ElseIf iplanno = "K" Then
                    dr34("INSAMT") = "150000"
                Else
                    dr34("INSAMT") = "0"
                End If
                dr34("INSPREM") = txtIT34.Text
                ds.IQPOBJ.AddIQPOBJRow(dr34)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT35
            If (txtIT35.Text <> "0" And txtIT35.Text <> "") Then

                Dim dr35 As beITP101000.IQPOBJRow
                dr35 = ds.IQPOBJ.NewIQPOBJRow
                dr35("IISSU") = txtIins_ISSU2.Text
                dr35("INSTyp") = "IT35"
                If iplanno = "J" Then
                    dr35("INSAMT") = "120000"
                ElseIf iplanno = "K" Then
                    dr35("INSAMT") = "150000"
                Else
                    dr35("INSAMT") = "0"
                End If
                dr35("INSPREM") = txtIT35.Text
                ds.IQPOBJ.AddIQPOBJRow(dr35)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT2P
            If (txtIT2P.Text <> "0" And txtIT2P.Text <> "") Then

                Dim dr2P As beITP101000.IQPOBJRow
                dr2P = ds.IQPOBJ.NewIQPOBJRow
                dr2P("IISSU") = txtIins_ISSU2.Text
                dr2P("INSTyp") = "IT2P"
                If iplanno = "J" Then
                    dr2P("INSAMT") = "3000"
                ElseIf iplanno = "K" Then
                    dr2P("INSAMT") = "5000"
                Else
                    dr2P("INSAMT") = "0"
                End If
                dr2P("INSPREM") = txtIT2P.Text
                ds.IQPOBJ.AddIQPOBJRow(dr2P)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT2U
            If (txtIT2U.Text <> "0" And txtIT2U.Text <> "") Then

                Dim dr2U As beITP101000.IQPOBJRow
                dr2U = ds.IQPOBJ.NewIQPOBJRow
                dr2U("IISSU") = txtIins_ISSU2.Text
                dr2U("INSTyp") = "IT2U"
                If iplanno = "J" Then
                    dr2U("INSAMT") = "10000"
                ElseIf iplanno = "K" Then
                    dr2U("INSAMT") = "10000"
                Else
                    dr2U("INSAMT") = "0"
                End If
                dr2U("INSPREM") = txtIT2U.Text
                ds.IQPOBJ.AddIQPOBJRow(dr2U)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT2T
            If (txtIT2T.Text <> "0" And txtIT2T.Text <> "") Then

                Dim dr2T As beITP101000.IQPOBJRow
                dr2T = ds.IQPOBJ.NewIQPOBJRow
                dr2T("IISSU") = txtIins_ISSU2.Text
                dr2T("INSTyp") = "IT2T"
                If iplanno = "J" Then
                    dr2T("INSAMT") = "3000"
                ElseIf iplanno = "K" Then
                    dr2T("INSAMT") = "5000"
                Else
                    dr2T("INSAMT") = "0"
                End If
                dr2T("INSPREM") = txtIT2T.Text
                ds.IQPOBJ.AddIQPOBJRow(dr2T)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT2S
            If (txtIT2S.Text <> "0" And txtIT2S.Text <> "") Then

                Dim dr2S As beITP101000.IQPOBJRow
                dr2S = ds.IQPOBJ.NewIQPOBJRow
                dr2S("IISSU") = txtIins_ISSU2.Text
                dr2S("INSTyp") = "IT2S"
                If iplanno = "J" Then
                    dr2S("INSAMT") = "3000"
                ElseIf iplanno = "K" Then
                    dr2S("INSAMT") = "5000"
                Else
                    dr2S("INSAMT") = "0"
                End If
                dr2S("INSPREM") = txtIT2S.Text
                ds.IQPOBJ.AddIQPOBJRow(dr2S)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            'IT2G
            If (txtIT2G.Text <> "0" And txtIT2G.Text <> "") Then

                Dim drD As beITP101000.IQPOBJRow
                drD = ds.IQPOBJ.NewIQPOBJRow
                drD("IISSU") = txtIins_ISSU2.Text
                drD("INSTyp") = "IT2G"
                If iplanno = "A" Or iplanno = "B" Then
                    drD("INSAMT") = "3000"
                Else
                    drD("INSAMT") = "0"
                End If
                drD("INSPREM") = txtIT2G.Text
                ds.IQPOBJ.AddIQPOBJRow(drD)
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
        End If


    End Sub
    Sub GetDefault(ByVal age As Integer)
        ddlIT41.SelectedValue = ""
        txtIT41.Text = ""
        ddlIT42.SelectedValue = ""
        txtIT42.Text = ""
        ddlIT43.SelectedValue = ""
        txtIT43.Text = ""
        ddlIT44.SelectedValue = ""
        txtIT44.Text = ""
        ddlIT4Y.SelectedValue = ""
        txtIT4Y.Text = ""
        ddlIT61.SelectedValue = ""
        txtIT61.Text = ""
        ddlIT63.SelectedValue = ""
        txtIT63.Text = ""
        ddlIT64.SelectedValue = ""
        txtIT64.Text = ""
        ddlIT6E.SelectedValue = ""
        txtIT6E.Text = ""
        lblIT65.Text = ""
        txtIT65.Text = ""
        lblIT66.Text = ""
        txtIT66.Text = ""
        lblIT6F.Text = ""
        txtIT6F.Text = ""
        lblIT6G.Text = ""
        txtIT6G.Text = ""
        lblIT6D.Text = ""
        txtIT6D.Text = ""
        If age < 15 Then
            ddlIT4Y.SelectedValue = "69"
            ddlIT42.SelectedValue = "20"
            ddlIT44.SelectedValue = "25"
        ElseIf age >= 86 Then
            ddlIT41.SelectedValue = "50"
            ddlIT42.SelectedValue = "5"
            ddlIT44.SelectedValue = ""
        ElseIf age >= 81 Then
            ddlIT41.SelectedValue = "100"
            ddlIT42.SelectedValue = "10"
            ddlIT44.SelectedValue = "25"
        ElseIf age >= 76 Then
            ddlIT41.SelectedValue = "200"
            ddlIT42.SelectedValue = "20"
            ddlIT44.SelectedValue = "50"
        ElseIf age >= 71 Then
            ddlIT41.SelectedValue = "500"
            ddlIT42.SelectedValue = "50"
            ddlIT44.SelectedValue = "125"
        ElseIf age >= 21 Then
            ddlIT41.SelectedValue = "1000"
            ddlIT42.SelectedValue = "100"
            ddlIT44.SelectedValue = "250"
        ElseIf age >= 15 Then
            ddlIT41.SelectedValue = "500"
            ddlIT42.SelectedValue = "50"
            ddlIT44.SelectedValue = "125"
        End If

    End Sub
    Sub GetDay()

        Dim dateDOBJBGN As Date = Me.UcRocDobjbgn.GetDate
        Dim lQday As Integer = DateDiff("d", Me.UcRocDobjbgn.GetDate, Me.UcRocDobjend.GetDate)
        ' IQP 天數計算需考慮時間.
        Dim iTobjbgn As Integer = 0
        If txtTobjbgn.Text.TrimEnd <> String.Empty Then
            iTobjbgn = CInt(txtTobjbgn.Text.TrimEnd)
        End If
        Dim iTobjend As Integer = 0
        If txtTobjend.Text.TrimEnd <> String.Empty Then
            iTobjend = CInt(txtTobjend.Text.TrimEnd)
        End If
        If iTobjbgn < iTobjend Then
            lQday += 1
        End If
        txtIQPITP_Days.Text = lQday

    End Sub
    Private Sub CK1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CK1.CheckedChanged
        CHK5.Enabled = True
        txtTobjbgn.ReadOnly = False
        txtTobjend.ReadOnly = False
        UcRocDobjbgn.Enabled = True
        UcRocDobjend.Enabled = True
        UcRocDobjbgn.ReadOnly = False
        UcRocDobjend.ReadOnly = False
        'ReCaculate.Enabled = True
        ddl_IINS.Enabled = True
        ddl_IINS2.Enabled = True
        txtIins_IINS2.Enabled = True
        '不能清掉方案
        'Me.txtCase.Text = ""
        CheckS.Text = "1"
        CheckSNNN.Text = "請[重新計算保費]"
    End Sub
    Private Sub CK2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CK2.CheckedChanged
        CHK5.Enabled = True
        txtTobjbgn.ReadOnly = False
        txtTobjend.ReadOnly = False
        UcRocDobjbgn.Enabled = True
        UcRocDobjend.Enabled = True
        UcRocDobjbgn.ReadOnly = False
        UcRocDobjend.ReadOnly = False
        'ReCaculate.Enabled = True
        ddl_IINS.Enabled = True
        ddl_IINS2.Enabled = True
        txtIins_IINS2.Enabled = True
        Me.txtCase.Text = ""
        CheckS.Text = "1"
        CheckSNNN.Text = "請[重新計算保費]"
    End Sub
    Function CheckTime() As String
        Try

            Dim DateErr As String = ""
            Me.IsInputErr = False

            If IsNumeric(txtTobjbgn.Text) = True Then
                If CInt(txtTobjbgn.Text) >= 24 Then
                    'txtTobjbgn.Text = "24"
                    DateErr &= "時間起 請輸入時間區間00~23" & Environment.NewLine
                End If
                If CInt(txtTobjbgn.Text) < 0 Then
                    'txtTobjbgn.Text = "00"
                    DateErr &= "時間起 請輸入時間區間00~23" & Environment.NewLine
                End If
            Else
                DateErr &= "時間起 請輸入數字." & Environment.NewLine
            End If
            'txtTobjend
            If IsNumeric(txtTobjend.Text) = True Then
                If CInt(txtTobjend.Text) >= 24 Then
                    'txtTobjbgn.Text = "24"
                    DateErr &= "時間迄 請輸入時間區間00~23" & Environment.NewLine
                End If
                If CInt(txtTobjend.Text) < 0 Then
                    'txtTobjbgn.Text = "00"
                    DateErr &= "時間迄 請輸入時間區間00~23" & Environment.NewLine
                End If
            Else
                DateErr &= "時間迄 請輸入數字." & Environment.NewLine
            End If
            '202400115 F113011200 起始輸入天數(由原本的二個工作天改為投保前一個天)
            If UcRocDobjbgn.GetDate <= Now.Date().Date Then
                DateErr &= " 保險生效日最早須為 " & Now.Date().AddDays(1) & " 0時，若欲一日內生效之案件，則無法受理。" & Environment.NewLine
            End If
            'If Me.ScsOrgID = "SKI" Or Me.ScsOrgID = "A3RY" Or Me.ScsOrgID = "A045" Then
            '    '20230417 甲類業務員登打上述險種時，起保日限登打日+1天
            '    If UcRocDobjbgn.GetDate <= Now.Date().Date Then
            '        DateErr &= " 保險生效日最早須為 " & Now.Date().AddDays(1) & " 0時，若欲一日內生效之案件，則無法受理。" & Environment.NewLine
            '    End If
            'Else
            '    '來源：非SKI
            '    If Me.TRAVEL_A.SelectedValue = "400" Then
            '        Dim sql1 As String = " (select ddate from psnp_date where  ddate > to_date('"
            '        Dim sql2 As String = "', 'YYYY/MM/DD') and kwork = '1' and rownum <= 2  ) "
            '        Dim sql3 As String = " 1=1  ORDER BY ddate desc"
            '        Dim ds_now As DataSet = Me.CallCommonService("PUBService.bsPUBSingleSQL", New String() {sql1 + Now.Date().Date + sql2, sql3, "ddate"})
            '        If UcRocDobjbgn.GetDate < ds_now.Tables(0).Rows(0)("ddate") Then
            '            DateErr &= "國內旅遊 保險生效日最早須為 " & ds_now.Tables(0).Rows(0)("ddate").ToShortDateString & " 0時，若欲二日內(工作日)生效之案件，則無法受理。" & Environment.NewLine
            '        End If
            '    Else
            '        '來源：非SKI
            '        Dim sql1 As String = " (select ddate from psnp_date where  ddate > to_date('"
            '        Dim sql2 As String = "', 'YYYY/MM/DD') and kwork = '1' and rownum <= 2 ) "
            '        Dim sql3 As String = " 1=1  ORDER BY ddate desc"
            '        Dim ds_now As DataSet = Me.CallCommonService("PUBService.bsPUBSingleSQL", New String() {sql1 + Now.Date().Date + sql2, sql3, "ddate"})

            '        If UcRocDobjbgn.GetDate < ds_now.Tables(0).Rows(0)("ddate") Then
            '            DateErr &= "國外旅遊 保險生效日最早須為 " & ds_now.Tables(0).Rows(0)("ddate").ToShortDateString & " 0時，若欲二日內(工作日)生效之案件，則無法受理。" & Environment.NewLine
            '        End If
            '    End If
            'End If

            '生生效日、到期日同一天，要控管生效時需小於到期時。
            If UcRocDobjbgn.GetDate = UcRocDobjend.GetDate Then
                If CInt(txtTobjbgn.Text) >= CInt(txtTobjend.Text) Then
                    DateErr &= "保險期間為同一天，生效時需小於到期時" & Environment.NewLine
                End If
            End If

            '航班時間檢核，航班時間必須在保險期間
            '至少一個航班要從台灣出發TPE、TSA、KHH、RMQ
            '115/08/19說不檢核去、回航程的數量 （原設定_航程必須有去程 & 回程）
            Dim VDep As Boolean = False
            Dim VArr As Boolean = False
            Dim HasFlight As Boolean = False
            Dim TWDep As Boolean = False
            Dim TWAirports As String() = New String() {"TPE", "TSA", "KHH", "RMQ", "TTT", "TNN", "CYI", "KYD", "HUN", "MZG", "KNH", "MFK", "LZN", "GNI", "CMJ", "WOT", "HCN"}
            If dgdFlight.Items.Count > 0 Then
                For Each item As DataGridItem In dgdFlight.Items
                    If item.ItemType = ListItemType.Item OrElse item.ItemType = ListItemType.AlternatingItem Then
                        HasFlight = True
                        Dim labFDeparture As Label = CType(item.FindControl("labDeparturetime"), Label)
                        Dim labFArrival As Label = CType(item.FindControl("labArrivaltime"), Label)
                        Dim labVoyage As Label = CType(item.FindControl("labVoyage"), Label)
                        Dim labDepartureAirport_IATA As Label = CType(item.FindControl("labDepartureAirport_IATA"), Label)

                        '航班時間與保險期間檢核
                        Dim FTimeCheck As Tuple(Of Boolean, String) = ConfirmFlightDateTime(
                                labFDeparture.Text,
                                labFArrival.Text,
                                UcRocDobjbgn.GetDate.ToString("yyyy/MM/dd"),
                                UcRocDobjend.GetDate.ToString("yyyy/MM/dd"),
                                CInt(txtTobjbgn.Text),
                                CInt(txtTobjend.Text)
                            )
                        If FTimeCheck.Item1 = False Then
                            DateErr &= FTimeCheck.Item2 & Environment.NewLine
                            Exit For
                        End If

                        '檢查有無台灣出發的航班
                        If TWAirports.Contains(labDepartureAirport_IATA.Text.ToUpper) Then
                            TWDep = True
                        End If

                        ''取航程
                        'If labVoyage.Text = "去程" Then
                        '    VDep = True
                        'ElseIf labVoyage.Text = "回程" Then
                        '    VArr = True
                        'End If
                    End If
                Next
            End If

            If TWDep = False And dgdFlight.Items.Count > 0 Then
                DateErr &= "必須有台灣機場出發的航班" & Environment.NewLine
            End If

            'If HasFlight = True AndAlso (VDep = False Or VArr = False) Then
            '    DateErr &= "航班必須有 去程 與 回程" & Environment.NewLine
            'End If

            Return DateErr

        Catch ex As Exception
            'Me.ProcessException(ex)
            Me.UIAppendError(ex.Message.ToString)
        End Try
    End Function
    Function CheckInsert() As String
        Try

            Dim strErr As String = ""
            Me.IsInputErr = False

            If txtIins_IINS2.Text = "" Then
                Me.IsInputErr = True
                strErr &= "要保人ID 不可為空白." & Environment.NewLine
            End If
            If ddl_IINS.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "請重新輸入要保人ID，並注意ID填寫正確" & Environment.NewLine '"要保人類別 不可為空白." & Environment.NewLine
            End If
            'NAME_IINS
            If NAME_IINS.Text = "" Then
                Me.IsInputErr = True
                strErr &= "要保人姓名 不可為空白." & Environment.NewLine
            End If


            '選統編的 除外
            If ddl_IINS.SelectedValue <> "2" Then
                If Birtdy_IINS.Text = "" Then
                    Me.IsInputErr = True
                    strErr &= "要保人生日 不可為空白." & Environment.NewLine
                Else
                    '要保人要大於7歲
                    Dim age As Integer = B2BEntity.HasFunction.GetAgeInLaw(Me.Birtdy_IINS.GetDate(), Me.UcRocDobjbgn.GetDate())
                    If age <= 7 Then
                        Me.IsInputErr = True
                        strErr &= "要保人 應年滿7歲." & Environment.NewLine
                    End If
                End If
            End If
            '選統編的
            If ddl_IINS.SelectedValue = "2" Then
                If txtPrincipal.Text = "" Then
                    Me.IsInputErr = True
                    strErr &= "負責人姓名 不可為空白." & Environment.NewLine
                End If
            End If
            If IINS_AREA.Text = "" Then
                Me.IsInputErr = True
                strErr &= "要保人郵遞區號 不可為空白." & Environment.NewLine
            End If
            If IINS_ADDR.Text = "" Then
                Me.IsInputErr = True
                strErr &= "要保人地址 不可為空白." & Environment.NewLine
            End If
            '聯絡電話
            If IINS_TEL.Text = "" Then
                Me.IsInputErr = True
                strErr &= "要保人聯絡電話 不可為空白." & Environment.NewLine
            Else
                If IINS_TEL.Text.Trim.Length < 6 Then
                    Me.IsInputErr = True
                    strErr &= "要保人聯絡電話 格式錯誤." & Environment.NewLine
                End If
            End If

            If UcRocDobjbgn.Text = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間起 不可為空白." & Environment.NewLine
            Else
                If UcRocDobjbgn.GetDate > Now.AddMonths(3) Then
                    Me.IsInputErr = True
                    strErr &= "出發日 不得大於系統日3個月." & Environment.NewLine
                End If
            End If

            If txtTobjbgn.Text = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間時間起 不可為空白." & Environment.NewLine
            End If
            If UcRocDobjend.Text = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間訖 不可為空白." & Environment.NewLine
            End If
            If txtTobjend.Text = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間時間訖 不可為空白." & Environment.NewLine
            End If


            'TRAVEL_A
            If TRAVEL_A.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "旅遊地區 不可為空白." & Environment.NewLine
            End If
            '旅遊地點
            If TRAVEL_BEGIN.Text = "" Then
                ''Me.IsInputErr = True
                ''strErr &= "啟程地 不可為空白." & Environment.NewLine
                TRAVEL_BEGIN.Text = "台灣"
            End If

            TRAVEL_END.Text = ddlCountry_A2.SelectedItem.ToString
            If ddlCountry_B2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_B2.SelectedItem.ToString
            End If
            If ddlCountry_C2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_C2.SelectedItem.ToString
            End If
            If ddlCountry_D2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_D2.SelectedItem.ToString
            End If
            If ddlCountry_E2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_E2.SelectedItem.ToString
            End If
            If ddlCountry_F2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_F2.SelectedItem.ToString
            End If
            If ddlCountry_G2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_G2.SelectedItem.ToString
            End If
            If ddlCountry_H2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_H2.SelectedItem.ToString
            End If

            If ddlCountry_A2.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "旅遊國家 地區不可為空白." & Environment.NewLine
                'TRAVEL_END.Text = ddlCountry_A2.SelectedItem.ToString
            Else
                'Dim RedArea As String = ConfigurationManager.AppSettings("RedArea")
                'Dim AreaArray As String()
                'AreaArray = RedArea.Split("|")
                'For i As Integer = 0 To AreaArray.Length - 1
                '    If TRAVEL_END.Text = (AreaArray(i)) Then
                '        strErr &= "旅遊地區為紅色警戒地區，本公司不予承保!." & Environment.NewLine
                '    End If
                'Next
            End If


            '被保人 
            'txtIins_ISSU2
            If txtIins_ISSU2.Text = "" Then
                Me.IsInputErr = True
                strErr &= "被保人ID 不可為空白." & Environment.NewLine
            End If
            If ddl_IISSU.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "請重新輸入被保人ID，並注意ID填寫正確" & Environment.NewLine '"被保人類別 不可為空白." & Environment.NewLine
            End If

            If NAME_IISSU.Text = "" Then
                Me.IsInputErr = True
                strErr &= "被保人姓名 不可為空白." & Environment.NewLine
            End If
            'UcRocDbirth
            If UcRocDbirth.Text = "" Then
                Me.IsInputErr = True
                strErr &= "被保人生日 不可為空白." & Environment.NewLine
            End If
            If ddlSex.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "被保人性別 不可為空白." & Environment.NewLine
            End If
            If IISSU_AREA.Text = "" Then
                Me.IsInputErr = True
                strErr &= "被保人郵遞區號 不可為空白." & Environment.NewLine
            End If
            If IISSU_ADDR.Text = "" Then
                Me.IsInputErr = True
                strErr &= "被保人地址 不可為空白." & Environment.NewLine
            End If
            'IISSU_TEL
            If IISSU_TEL.Text = "" Then
                Me.IsInputErr = True
                strErr &= "被保人行動電話 不可為空白." & Environment.NewLine
            Else
                If IISSU_TEL.Text.Trim.Length < 6 Then
                    Me.IsInputErr = True
                    strErr &= "被保人行動電話 不可為空白." & Environment.NewLine
                ElseIf IsMobileNumber(IISSU_TEL.Text) = False Then
                    Me.IsInputErr = True
                    strErr &= "被保人行動電話 格式錯誤." & Environment.NewLine
                End If
            End If
            If Relationship.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "與要保人關係 不可為空白." & Environment.NewLine
            End If

            If Benefi.Text.Trim = "法定繼承人" Then
                '法定繼承人 無身故保險金 不用檢核 受益人ＩＤ姓名關係
            ElseIf Benefi.Text = "" Then
                Me.IsInputErr = True
                strErr &= "受益人 不可為空白." & Environment.NewLine
            Else
                If Benefi_ID.Text = "" Then
                    Me.IsInputErr = True
                    strErr &= "受益人ID 不可為空白." & Environment.NewLine
                End If
                'Relationship2
                If Relationship2.SelectedValue = "" Then
                    Me.IsInputErr = True
                    strErr &= "與被保人關係ID 不可為空白." & Environment.NewLine
                End If
                If txtZIPREL.Text = "" Then
                    Me.IsInputErr = True
                    strErr &= "受益人郵遞區號 不可為空白." & Environment.NewLine
                End If
                If txtAREL.Text = "" Then
                    Me.IsInputErr = True
                    strErr &= "受益人地址 不可為空白." & Environment.NewLine
                End If
                If txtTELREL.Text = "" Then
                    Me.IsInputErr = True
                    strErr &= "受益人聯絡電話 不可為空白." & Environment.NewLine
                Else
                    If txtTELREL.Text.Trim.Length < 6 Then
                        Me.IsInputErr = True
                        strErr &= "受益人聯絡電話 格式錯誤." & Environment.NewLine
                    End If
                End If
            End If
            If Benefi2.Text.Trim <> "" Then
                If Benefi2.Text.Trim = "法定繼承人" Then
                    '法定繼承人 無身故保險金 不用檢核 受益人ＩＤ姓名關係
                Else
                    If Benefi_ID2.Text = "" Then
                        Me.IsInputErr = True
                        strErr &= "受益人2 ID 不可為空白." & Environment.NewLine
                    End If
                    'Relationship2
                    If Relationship22.SelectedValue = "" Then
                        Me.IsInputErr = True
                        strErr &= "與被保人關係2 ID 不可為空白." & Environment.NewLine
                    End If
                End If
            End If
            If Benefi3.Text.Trim <> "" Then
                If Benefi3.Text.Trim = "法定繼承人" Then
                    '法定繼承人 無身故保險金 不用檢核 受益人ＩＤ姓名關係
                Else
                    If Benefi_ID3.Text = "" Then
                        Me.IsInputErr = True
                        strErr &= "受益人3 ID 不可為空白." & Environment.NewLine
                    End If
                    If Relationship23.SelectedValue = "" Then
                        Me.IsInputErr = True
                        strErr &= "與被保人關係3 ID 不可為空白." & Environment.NewLine
                    End If
                End If
            End If

            Dim CaseCode As String = ""
            If ddlTargetProject.SelectedValue.Trim = "ITP" Then
                '投保方案
                If txtCase.Text = "" Or lblTargetPlan.Text.Trim = "" Then '新增新光金保代可以單出不便險
                    Me.IsInputErr = True
                    strErr &= "請選擇投保方案." & Environment.NewLine
                ElseIf txtCase.Text <> "" Then
                    CaseCode = (Me.txtCase.Text).Substring(0, 1)
                End If
            End If

            Dim age2 As Integer = B2BEntity.HasFunction.GetAgeInLaw(Me.UcRocDbirth.GetDate(), Me.UcRocDobjbgn.GetDate())
            If age2 < 15 Then
                If ddlIT4Y.SelectedValue = "" Or ddlIT4Y.SelectedValue = "0" Then
                    If ddl_IISSU.SelectedValue = "5" And TRAVEL_A.SelectedValue = "400" Then
                        Me.IsInputErr = True
                        strErr &= "國內地區喪葬費用保額不可為0." & Environment.NewLine
                    Else
                        If (checkSpecialChannel().Item1 = True Or checkChannelO().Item1) AndAlso (CaseCode = "J" Or CaseCode = "K" Or CaseCode = "L") Then

                            If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or ddlIT64.SelectedValue = "") Then
                                'Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                Me.IsInputErr = True
                                strErr &= "未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金." & Environment.NewLine
                            End If

                        ElseIf checkChannelO().Item1 = True AndAlso CaseCode = "O" Then

                            If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or (ddlIT63.SelectedValue = "" And ddlIT61.SelectedValue = "" And ddlIT64.SelectedValue = "")) Then
                                'Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                Me.IsInputErr = True
                                strErr &= "未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金." & Environment.NewLine
                            End If

                        ElseIf (CaseCode <> "J" And CaseCode <> "K" And CaseCode <> "L") AndAlso checkChannelMN().Item1 Then

                            If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or (ddlIT63.SelectedValue = "" And ddlIT61.SelectedValue = "" And ddlIT64.SelectedValue = "")) Then
                                'Throw New Exception("未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金.")
                                Me.IsInputErr = True
                                strErr &= "未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金." & Environment.NewLine
                            End If

                        ElseIf (CaseCode <> "J" And CaseCode <> "K" And CaseCode <> "L") Then

                            If (ddlIT42.SelectedValue = "" Or ddlIT42.SelectedValue = "0" Or (ddlIT61.SelectedValue = "" And (ddlIT63.SelectedValue = "" And ddlIT6E.SelectedValue = ""))) Then
                                Me.IsInputErr = True
                                strErr &= "未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金." & Environment.NewLine
                            End If

                        End If
                    End If
                End If
            Else
                '主保險契約 旅行平安保險
                If ddlIT41.SelectedValue = "" And ddlIT4Y.SelectedValue = "" Then
                    Me.IsInputErr = True
                    strErr &= "請選擇身故失能保險金 或 失能保險金." & Environment.NewLine
                End If
            End If

            If ddl_IISSU.SelectedValue = "1" Then
                If B2BEntity.HasFunction._CheckCitizenID(txtIins_ISSU2.Text.Trim) Then
                Else
                    Me.IsInputErr = True
                    strErr &= "被保險人身分證號 錯誤." & Environment.NewLine
                End If
            End If
            'If Me.IsInputErr Then
            '    Me.UIAppendMessage("資料有誤, 請修正後重新上傳.")
            'End If
            Return strErr

        Catch ex As Exception
            Me.ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Function
    Function CheckBasic() As String
        Try

            Dim strErr As String = ""
            Me.IsInputErr = False

            If txtIins_IINS2.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人ID 不可為空白." & Environment.NewLine
            End If
            If ddl_IINS.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "請重新輸入要保人ID，並注意ID填寫正確" & Environment.NewLine '"要保人類別 不可為空白." & Environment.NewLine
            End If
            'NAME_IINS
            If NAME_IINS.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人姓名 不可為空白." & Environment.NewLine
            Else
                ' 將換行符號剃除
                NAME_IINS.Text = (NAME_IINS.Text).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
            End If
            '選統編的 除外
            If ddl_IINS.SelectedValue <> "2" Then
                If Birtdy_IINS.Text = "" Then
                    Me.IsInputErr = True
                    strErr &= "要保人生日 不可為空白." & Environment.NewLine
                Else
                    Dim ageIINS As Integer = B2BEntity.HasFunction.GetAgeInLaw(Me.Birtdy_IINS.GetDate(), Me.UcRocDobjbgn.GetDate())
                    If ageIINS <= 7 Then
                        Me.IsInputErr = True
                        strErr &= "要保人 應年滿7歲." & Environment.NewLine
                    End If
                End If
                '要保人要大於7歲
            End If
            '選統編的
            If ddl_IINS.SelectedValue = "2" Then
                If txtPrincipal.Text = "" Then
                    Me.IsInputErr = True
                    strErr &= "負責人姓名 不可為空白." & Environment.NewLine
                End If
            End If

            If IINS_AREA.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人郵遞區號 不可為空白." & Environment.NewLine
            End If
            If IINS_ADDR.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人地址 不可為空白." & Environment.NewLine
            End If
            '聯絡電話
            If IINS_TEL.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人聯絡電話 不可為空白." & Environment.NewLine
            Else
                If IINS_TEL.Text.Trim.Length < 6 Then
                    Me.IsInputErr = True
                    strErr &= "要保人聯絡電話 格式錯誤." & Environment.NewLine
                End If
            End If
            'UcRocDobjbgn
            If UcRocDobjbgn.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間起 不可為空白." & Environment.NewLine
            Else
                If UcRocDobjbgn.GetDate > Now.AddMonths(3) Then
                    Me.IsInputErr = True
                    strErr &= "出發日 不得大於系統日3個月." & Environment.NewLine
                End If
            End If

            If txtTobjbgn.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間時間起 不可為空白." & Environment.NewLine
            End If
            If UcRocDobjend.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間訖 不可為空白." & Environment.NewLine
            End If
            If txtTobjend.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間時間訖 不可為空白." & Environment.NewLine
            End If

            If txtIQPITP_Days.Text = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間 不可為空白." & Environment.NewLine
            Else
                ''時間檢核
                If CInt(txtIQPITP_Days.Text) < 0 Then
                    Me.IsInputErr = True
                    strErr &= "保險期間迄日必須大於起日." & Environment.NewLine
                End If
                'If CInt(txtIQPITP_Days.Text) > 60 Then
                '    Me.IsInputErr = True
                '    strErr &= "保險期間不得超過60天." & Environment.NewLine
                'End If
            End If


            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If
            '旅遊國家 組合判斷2

            If ddlCountry_B2.SelectedValue = "" And ddlCountry_C2.SelectedValue = "" And ddlCountry_D2.SelectedValue = "" And
                ddlCountry_E2.SelectedValue = "" And ddlCountry_F2.SelectedValue = "" And ddlCountry_G2.SelectedValue = "" And ddlCountry_H2.SelectedValue = "" Then
                '只有一組 用第一組判斷就好
                COUNCODE1 = HttpUtility.HtmlEncode(ddlCountry_A2.SelectedValue.Substring(2, 3))
                If COUNCODE1 = "750" Then
                    '申根地區判斷
                    If Checkbox2.Checked = True Then
                        If PASSPORT_NAME.Text = "" Or PASSPORT_NUMBER.Text = "" Then
                            Me.IsInputErr = True
                            strErr &= "勾選申根公約國證明文件 請輸入英文姓名 與 護照號碼。" & Environment.NewLine
                        Else
                            '申根地區
                            COUNCODE1 = "750"
                        End If
                    Else
                        '改成歐洲地區
                        COUNCODE1 = "740"
                    End If

                End If
            Else
                Dim ddlA, ddlB, ddlC, ddlD, ddlE, ddlF, ddlG, ddlH As String

                If ddlCountry_A2.SelectedValue <> "" Then
                    ddlA = HttpUtility.HtmlEncode(ddlCountry_A2.SelectedValue.Substring(2, 3))
                Else
                    ddlA = ""
                End If
                If ddlCountry_B2.SelectedValue <> "" Then
                    ddlB = HttpUtility.HtmlEncode(ddlCountry_B2.SelectedValue.Substring(2, 3))
                Else
                    ddlB = ""
                End If
                If ddlCountry_C2.SelectedValue <> "" Then
                    ddlC = HttpUtility.HtmlEncode(ddlCountry_C2.SelectedValue.Substring(2, 3))
                Else
                    ddlC = ""
                End If
                If ddlCountry_D2.SelectedValue <> "" Then
                    ddlD = HttpUtility.HtmlEncode(ddlCountry_D2.SelectedValue.Substring(2, 3))
                Else
                    ddlD = ""
                End If
                If ddlCountry_E2.SelectedValue <> "" Then
                    ddlE = HttpUtility.HtmlEncode(ddlCountry_E2.SelectedValue.Substring(2, 3))
                Else
                    ddlE = ""
                End If
                If ddlCountry_F2.SelectedValue <> "" Then
                    ddlF = HttpUtility.HtmlEncode(ddlCountry_F2.SelectedValue.Substring(2, 3))
                Else
                    ddlF = ""
                End If
                If ddlCountry_G2.SelectedValue <> "" Then
                    ddlG = HttpUtility.HtmlEncode(ddlCountry_G2.SelectedValue.Substring(2, 3))
                Else
                    ddlG = ""
                End If
                If ddlCountry_H2.SelectedValue <> "" Then
                    ddlH = HttpUtility.HtmlEncode(ddlCountry_H2.SelectedValue.Substring(2, 3))
                Else
                    ddlH = ""
                End If

                If ddlA = "400" Then
                    If (ddlB <> "400" And ddlB <> "") Then
                        Me.IsInputErr = True
                        strErr &= "選擇台灣地區時無法再選擇國外區域" & Environment.NewLine
                    End If
                    If (ddlC <> "400" And ddlC <> "") Then
                        Me.IsInputErr = True
                        strErr &= "選擇台灣地區時無法再選擇國外區域" & Environment.NewLine
                    End If
                    If (ddlD <> "400" And ddlD <> "") Then
                        Me.IsInputErr = True
                        strErr &= "選擇台灣地區時無法再選擇國外區域" & Environment.NewLine
                    End If
                    If (ddlE <> "400" And ddlE <> "") Then
                        Me.IsInputErr = True
                        strErr &= "選擇台灣地區時無法再選擇國外區域" & Environment.NewLine
                    End If
                    If (ddlF <> "400" And ddlF <> "") Then
                        Me.IsInputErr = True
                        strErr &= "選擇台灣地區時無法再選擇國外區域" & Environment.NewLine
                    End If
                    If (ddlG <> "400" And ddlG <> "") Then
                        Me.IsInputErr = True
                        strErr &= "選擇台灣地區時無法再選擇國外區域" & Environment.NewLine
                    End If
                    If (ddlH <> "400" And ddlH <> "") Then
                        Me.IsInputErr = True
                        strErr &= "選擇台灣地區時無法再選擇國外區域" & Environment.NewLine
                    End If
                    COUNCODE1 = "400"
                Else
                    If ddlB = "400" Or ddlC = "400" Or ddlD = "400" Or ddlE = "400" Or ddlF = "400" Or ddlG = "400" Or ddlH = "400" Then
                        Me.IsInputErr = True
                        strErr &= "選擇國外地區時無法再選擇台灣區域" & Environment.NewLine
                    End If
                    '全球各地
                    If ddlA = "750" Or ddlB = "750" Or ddlC = "750" Or ddlD = "750" Or ddlE = "750" Or ddlF = "750" Or ddlG = "750" Or ddlH = "750" Then
                        '申根地區判斷
                        If Checkbox2.Checked = True Then
                            If PASSPORT_NAME.Text = "" Or PASSPORT_NUMBER.Text = "" Then
                                Me.IsInputErr = True
                                strErr &= "勾選申根公約國證明文件 請輸入英文姓名 與 護照號碼。" & Environment.NewLine
                            Else
                                '申根地區
                                COUNCODE1 = "750"
                            End If
                        Else
                            '改成歐洲地區
                            COUNCODE1 = "740"
                        End If
                    Else
                        '改成全球地區
                        COUNCODE1 = "800"
                    End If
                End If

            End If
            TRAVEL_A.SelectedValue = COUNCODE1
            Me.Check750.Text = HttpUtility.HtmlEncode(TRAVEL_A.SelectedValue)
            'TRAVEL_A
            If TRAVEL_A.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "旅遊地區 不可為空白." & Environment.NewLine
            End If
            '旅遊地點
            If TRAVEL_BEGIN.Text.Trim = "" Then
                ''Me.IsInputErr = True
                ''strErr &= "啟程地 不可為空白." & Environment.NewLine
                TRAVEL_BEGIN.Text = "台灣"
            End If

            TRAVEL_END.Text = ddlCountry_A2.SelectedItem.ToString
            If ddlCountry_B2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_B2.SelectedItem.ToString
            End If
            If ddlCountry_C2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_C2.SelectedItem.ToString
            End If
            If ddlCountry_D2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_D2.SelectedItem.ToString
            End If
            If ddlCountry_E2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_E2.SelectedItem.ToString
            End If
            If ddlCountry_F2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_F2.SelectedItem.ToString
            End If
            If ddlCountry_G2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_G2.SelectedItem.ToString
            End If
            If ddlCountry_H2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_H2.SelectedItem.ToString
            End If

            If ddlCountry_A2.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "旅遊國家 地區不可為空白." & Environment.NewLine
                'TRAVEL_END.Text = ddlCountry_A2.SelectedItem.ToString
            Else
                'Dim RedArea As String = ConfigurationManager.AppSettings("RedArea")
                'Dim AreaArray As String()
                'AreaArray = RedArea.Split("|")
                'For i As Integer = 0 To AreaArray.Length - 1
                '    If TRAVEL_END.Text = (AreaArray(i)) Then
                '        strErr &= "旅遊地區為紅色警戒地區，本公司不予承保!." & Environment.NewLine
                '    End If
                'Next
            End If

            '20250508 提出移除中國大陸地區限制
            'If rad_Disease_Y.Checked And TRAVEL_END.Text.Contains("中國") Then
            '	ddlIT63.SelectedValue = ""
            '	strErr &= "選擇中國大陸地區時，不建議保﹝海外突發疾病保險金(含傳染病)﹞。" & Environment.NewLine
            'End If
            'If ddlIT63.SelectedValue <> "" And TRAVEL_END.Text.Contains("中國") Then
            '	ddlIT63.SelectedValue = ""
            '	strErr &= "選擇中國大陸地區時，不建議保﹝海外突發疾病保險金(含傳染病)﹞。" & Environment.NewLine
            'End If

            ' 檢查 要保人若使用護照號碼投保，資料是否填寫正確
            If ddl_IINS.SelectedValue.ToString = "5" Then
                If TRAVEL_A.SelectedValue <> "400" Then
                    Me.IsInputErr = True
                    strErr &= "要保人勾選「持外國人護照投保」，旅遊國家必須在「國內」." & Environment.NewLine
                Else
                    If ContainFullWidthChar(NAME_IINS.Text.Trim) = True Then
                        Me.IsInputErr = True
                        strErr &= "要保人勾選「持外國人護照投保」，要保人姓名必須為半形." & Environment.NewLine
                    End If

                    If Not IsAllENG(NAME_IINS.Text) Then
                        Me.IsInputErr = True
                        strErr &= "要保人勾選「持外國人護照投保」，要保人姓名必須為英文." & Environment.NewLine
                    End If

                    If ddlIINS_COUNTRY.SelectedValue = "TW" Or String.IsNullOrEmpty(ddlIINS_COUNTRY.SelectedValue.ToString) Then
                        Me.IsInputErr = True
                        strErr &= "要保人勾選「持外國人護照投保」，要保人國別不可為空、或為「TW」." & Environment.NewLine
                    End If
                End If
            End If

            '被保人
            'txtIins_ISSU2
            If txtIins_ISSU2.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "被保人ID 不可為空白." & Environment.NewLine
            End If
            If ddl_IISSU.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "請重新輸入被保人ID，並注意ID填寫正確" & Environment.NewLine '"被保人類別 不可為空白." & Environment.NewLine
            End If
            If NAME_IISSU.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "被保人姓名 不可為空白." & Environment.NewLine
            Else
                ' 將換行符號剃除
                NAME_IISSU.Text = (NAME_IISSU.Text).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
            End If
            'UcRocDbirth
            If UcRocDbirth.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "被保人生日 不可為空白." & Environment.NewLine
            End If
            If ddlSex.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "被保人性別 不可為空白." & Environment.NewLine
            End If
            If IISSU_AREA.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "被保人郵遞區號 不可為空白." & Environment.NewLine
            End If
            If IISSU_ADDR.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "被保人地址 不可為空白." & Environment.NewLine
            End If
            'IISSU_TEL
            If IISSU_TEL.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "被保人行動電話 不可為空白." & Environment.NewLine
            Else
                If IISSU_TEL.Text.Trim.Length < 6 Then
                    Me.IsInputErr = True
                    strErr &= "被保人行動電話 格式錯誤." & Environment.NewLine
                ElseIf IsMobileNumber(IISSU_TEL.Text) = False Then
                    Me.IsInputErr = True
                    strErr &= "被保人行動電話 格式錯誤." & Environment.NewLine
                End If
            End If
            'EMAIL
            If Relationship.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "與要保人關係 不可為空白." & Environment.NewLine
            End If

            ' 檢查 被保人若使用護照號碼投保，資料是否填寫正確
            If ddl_IISSU.SelectedValue.ToString = "5" Then
                If TRAVEL_A.SelectedValue <> "400" Then
                    Me.IsInputErr = True
                    strErr &= "被保人勾選「持外國人護照投保」，旅遊國家必須在「國內」." & Environment.NewLine
                Else
                    If ContainFullWidthChar(NAME_IISSU.Text.Trim) = True Then
                        Me.IsInputErr = True
                        strErr &= "被保人勾選「持外國人護照投保」，被保人姓名必須為半形." & Environment.NewLine
                    End If

                    If Not IsAllENG(NAME_IISSU.Text) Then
                        Me.IsInputErr = True
                        strErr &= "被保人勾選「持外國人護照投保」，被保人姓名必須為英文." & Environment.NewLine
                    End If

                    If ddlIISSU_COUNTRY.SelectedValue = "TW" Or String.IsNullOrWhiteSpace(ddlIISSU_COUNTRY.SelectedValue) Then
                        Me.IsInputErr = True
                        strErr &= "被保人勾選「持外國人護照投保」，被保人國別不可為空、或為「TW」." & Environment.NewLine
                    End If
                End If
            End If

            ' 檢查 海外旅遊時英文姓名 & 護照號碼，兩者選填，但是不可擇一填寫
            If TRAVEL_A.SelectedValue <> "400" Then
                Dim passportCheck As String = PASSPORT_NAME.Text.Trim & PASSPORT_NUMBER.Text.Trim
                If String.IsNullOrWhiteSpace(passportCheck) = False AndAlso (String.IsNullOrWhiteSpace(PASSPORT_NAME.Text.Trim) = True Or String.IsNullOrWhiteSpace(PASSPORT_NUMBER.Text.Trim) = True) Then
                    Me.IsInputErr = True
                    strErr &= "英文姓名 與 護照號碼 不可擇一填寫" & Environment.NewLine
                End If
            End If

            If Benefi.Text.Trim = "法定繼承人" Then
                '不用檢核 受益人ＩＤ姓名關係.
            ElseIf Benefi.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "受益人 不可為空白." & Environment.NewLine
            Else
                ' 將換行符號剃除
                Benefi.Text = (Benefi.Text).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")

                If Benefi_ID.Text.Trim = "" Then
                    Me.IsInputErr = True
                    strErr &= "受益人ID 不可為空白." & Environment.NewLine
                End If
                'Relationship2
                If Relationship2.SelectedValue = "" Then
                    Me.IsInputErr = True
                    strErr &= "與被保人關係ID 不可為空白." & Environment.NewLine
                End If

                If txtZIPREL.Text.Trim = "" Then
                    Me.IsInputErr = True
                    strErr &= "受益人郵遞區號 不可為空白." & Environment.NewLine
                End If
                If txtAREL.Text.Trim = "" Then
                    Me.IsInputErr = True
                    strErr &= "受益人地址 不可為空白." & Environment.NewLine
                End If
                If txtTELREL.Text.Trim = "" Then
                    Me.IsInputErr = True
                    strErr &= "受益人聯絡電話 不可為空白." & Environment.NewLine
                Else
                    If txtTELREL.Text.Trim.Length < 6 Then
                        Me.IsInputErr = True
                        strErr &= "受益人聯絡電話 格式錯誤." & Environment.NewLine
                    End If
                End If
            End If

            If Benefi2.Text.Trim <> "" Then
                If Benefi2.Text.Trim = "法定繼承人" Then
                    '法定繼承人 無身故保險金 不用檢核 受益人ＩＤ姓名關係
                Else
                    ' 將換行符號剃除
                    Benefi2.Text = (Benefi2.Text).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")

                    If Benefi_ID2.Text = "" Then
                        Me.IsInputErr = True
                        strErr &= "受益人2 ID 不可為空白." & Environment.NewLine
                    End If
                    'Relationship2
                    If Relationship22.SelectedValue = "" Then
                        Me.IsInputErr = True
                        strErr &= "與被保人關係2 ID 不可為空白." & Environment.NewLine
                    End If
                End If
            End If

            If Benefi3.Text.Trim <> "" Then
                If Benefi3.Text.Trim = "法定繼承人" Then
                    '法定繼承人 無身故保險金 不用檢核 受益人ＩＤ姓名關係
                Else
                    ' 將換行符號剃除
                    Benefi3.Text = (Benefi3.Text).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")

                    If Benefi_ID3.Text = "" Then
                        Me.IsInputErr = True
                        strErr &= "受益人3 ID 不可為空白." & Environment.NewLine
                    End If
                    If Relationship23.SelectedValue = "" Then
                        Me.IsInputErr = True
                        strErr &= "與被保人關係3 ID 不可為空白." & Environment.NewLine
                    End If
                End If
            End If

            If INSMOBILE.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "索取電子保單 請輸入要保人行動電話." & Environment.NewLine
            End If

            If INSMOBILE.Text.Trim <> "" Then
                If IsNumeric(INSMOBILE.Text.Trim) = True Then
                    If INSMOBILE.Text.Trim.Length <> 10 Then
                        Me.IsInputErr = True
                        strErr &= "行動電話請輸入數字10碼." & Environment.NewLine
                    End If
                    If INSMOBILE.Text.Substring(0, 2).Trim <> "09" Then
                        Me.IsInputErr = True
                        strErr &= "請輸入正確 行動電話 以方便簡訊傳送." & Environment.NewLine
                    End If
                Else
                    Me.IsInputErr = True
                    strErr &= "行動電話請輸入數字." & Environment.NewLine
                End If
            End If

            Return strErr

        Catch ex As Exception
            Me.ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Function
    Function CheckBasicRe() As String
        Try
            '當使用者查詢要保號後，直接執行"保費試算"，因為"被保人控制項"皆為空，下拉選單Relationship無值，
            '故透過StateMain.IQPISSU判斷是否存在【與被保人關係為"其他"】，有則tempDDL.selectedvalue為"其他"
            Dim tempDDL As New DropDownList()
            tempDDL = GetTempRelationshipDDL()

            CommonFunc.Relationship_addpeople_btn(CType(Me.StateMain, beITP101000).Tables("IQPISSU"), tempDDL, txtIins_IINS2, txtIins_ISSU2, txtRelationshipOther)

            Dim strErr As String = ""
            Me.IsInputErr = False

            If txtIins_IINS2.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人ID 不可為空白." & Environment.NewLine
            End If
            If ddl_IINS.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "請重新輸入要保人ID，並注意ID填寫正確" & Environment.NewLine '"要保人類別 不可為空白." & Environment.NewLine
            End If
            'NAME_IINS
            If NAME_IINS.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人姓名 不可為空白." & Environment.NewLine
            Else
                ' 將換行符號剃除
                NAME_IINS.Text = (NAME_IINS.Text).Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
            End If
            '選統編的 除外
            If ddl_IINS.SelectedValue <> "2" Then
                If Birtdy_IINS.Text.Trim = "" Then
                    Me.IsInputErr = True
                    strErr &= "要保人生日 不可為空白." & Environment.NewLine
                Else
                    '要保人要大於7歲
                    Dim ageIINS As Integer = B2BEntity.HasFunction.GetAgeInLaw(Me.Birtdy_IINS.GetDate(), Me.UcRocDobjbgn.GetDate())
                    If ageIINS <= 7 Then
                        Me.IsInputErr = True
                        strErr &= "要保人 應年滿7歲." & Environment.NewLine
                    End If
                End If
            End If
            '選統編的
            If ddl_IINS.SelectedValue = "2" Then
                If txtPrincipal.Text = "" Then
                    Me.IsInputErr = True
                    strErr &= "負責人姓名 不可為空白." & Environment.NewLine
                End If
            End If
            If IINS_AREA.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人郵遞區號 不可為空白." & Environment.NewLine
            End If
            If IINS_ADDR.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人地址 不可為空白." & Environment.NewLine
            End If
            '聯絡電話
            If IINS_TEL.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人聯絡電話 不可為空白." & Environment.NewLine
            Else
                If IINS_TEL.Text.Trim.Length < 6 Then
                    Me.IsInputErr = True
                    strErr &= "要保人聯絡電話 格式錯誤." & Environment.NewLine
                End If
            End If
            If UcRocDobjbgn.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間起 不可為空白." & Environment.NewLine
            Else
                If UcRocDobjbgn.GetDate > Now.AddMonths(3) Then
                    Me.IsInputErr = True
                    strErr &= "出發日 不得大於系統日3個月." & Environment.NewLine
                End If
            End If

            If txtTobjbgn.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間時間起 不可為空白." & Environment.NewLine
            End If
            If UcRocDobjend.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間訖 不可為空白." & Environment.NewLine
            End If
            If txtTobjend.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間時間訖 不可為空白." & Environment.NewLine
            End If

            'TRAVEL_A
            If TRAVEL_A.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "旅遊地區 不可為空白." & Environment.NewLine
            End If

            If ddlCountry_A2.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "旅遊國家 地區不可為空白." & Environment.NewLine
                'TRAVEL_END.Text = ddlCountry_A2.SelectedItem.ToString
            Else
                'Dim RedArea As String = ConfigurationManager.AppSettings("RedArea")
                'Dim AreaArray As String()
                'AreaArray = RedArea.Split("|")
                'For i As Integer = 0 To AreaArray.Length - 1
                '    If TRAVEL_END.Text = (AreaArray(i)) Then
                '        strErr &= "旅遊地區為紅色警戒地區，本公司不予承保!." & Environment.NewLine
                '    End If
                'Next
            End If

            If TRAVEL_A.SelectedValue = "400" Then
                If CInt(txtIQPITP_Days.Text) > 15 Then
                    strErr &= "國內旅遊 保險期間不得超過15天." & Environment.NewLine
                End If
            Else
                'F115060500 因投保項目而調整保險期間的天數
                Dim tmp_errStr As String
                For Each drE As DataRow In StateMain.Tables("IQPISSU").Rows
                    If drE.RowState <> DataRowState.Deleted Then
                        tmp_errStr = CheckPlanNTime(drE("ISSU_1"), drE("ISSU_2"), drE("ISSU_3"), drE("ISSU_3D"), drE("ISSU_IT64"), drE("ISSU_3D"))

                        If Not String.IsNullOrEmpty(tmp_errStr) Then
                            strErr &= tmp_errStr
                        End If
                    End If
                Next
            End If

            ' 檢查 要保人若使用護照號碼投保，資料是否填寫正確
            If ddl_IINS.SelectedValue.ToString = "5" Then
                If TRAVEL_A.SelectedValue <> "400" Then
                    Me.IsInputErr = True
                    strErr &= "要保人勾選「持外國人護照投保」，旅遊國家必須在「國內」." & Environment.NewLine
                Else
                    If ContainFullWidthChar(NAME_IINS.Text.Trim) = True Then
                        Me.IsInputErr = True
                        strErr &= "要保人勾選「持外國人護照投保」，要保人姓名必須為半形." & Environment.NewLine
                    End If

                    If Not IsAllENG(NAME_IINS.Text) Then
                        Me.IsInputErr = True
                        strErr &= "要保人勾選「持外國人護照投保」，要保人姓名必須為英文." & Environment.NewLine
                    End If

                    If ddlIINS_COUNTRY.SelectedValue = "TW" Or String.IsNullOrEmpty(ddlIINS_COUNTRY.SelectedValue.ToString) Then
                        Me.IsInputErr = True
                        strErr &= "要保人勾選「持外國人護照投保」，要保人國別不可為空、或為「TW」." & Environment.NewLine
                    End If
                End If
            End If

            '旅遊地點
            If TRAVEL_BEGIN.Text.Trim = "" Then
                TRAVEL_BEGIN.Text = "台灣"
            End If

            TRAVEL_END.Text = ddlCountry_A2.SelectedItem.ToString
            If ddlCountry_B2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_B2.SelectedItem.ToString
            End If
            If ddlCountry_C2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_C2.SelectedItem.ToString
            End If
            If ddlCountry_D2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_D2.SelectedItem.ToString
            End If
            If ddlCountry_E2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_E2.SelectedItem.ToString
            End If
            If ddlCountry_F2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_F2.SelectedItem.ToString
            End If
            If ddlCountry_G2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_G2.SelectedItem.ToString
            End If
            If ddlCountry_H2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_H2.SelectedItem.ToString
            End If

            '20250508 提出移除中國大陸地區限制
            'If rad_Disease_Y.Checked And TRAVEL_END.Text.Contains("中國") Then
            '	ddlIT63.SelectedValue = ""
            '	strErr &= "選擇中國大陸地區時，不建議保﹝海外突發疾病保險金(含傳染病)﹞。" & Environment.NewLine
            'End If
            'If ddlIT63.SelectedValue <> "" And TRAVEL_END.Text.Contains("中國") Then
            '	ddlIT63.SelectedValue = ""
            '	strErr &= "選擇中國大陸地區時，不建議保﹝海外突發疾病保險金(含傳染病)﹞。" & Environment.NewLine
            'End If

            If INSMOBILE.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "索取電子保單 請輸入要保人行動電話." & Environment.NewLine
            End If

            If INSMOBILE.Text.Trim <> "" Then
                If IsNumeric(INSMOBILE.Text.Trim) = True Then
                    If INSMOBILE.Text.Trim.Length <> 10 Then
                        Me.IsInputErr = True
                        strErr &= "行動電話請輸入數字10碼." & Environment.NewLine
                    End If
                    If INSMOBILE.Text.Substring(0, 2).Trim <> "09" Then
                        Me.IsInputErr = True
                        strErr &= "請輸入正確 行動電話 以方便簡訊傳送." & Environment.NewLine
                    End If
                Else
                    Me.IsInputErr = True
                    strErr &= "行動電話請輸入數字." & Environment.NewLine
                End If
            End If
            Return strErr

        Catch ex As Exception
            Me.ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Function
    Function CheckExlUP() As String
        Try

            Dim strErr As String = ""
            Me.IsInputErr = False

            If ddl_IINS.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "要保人類別 不可為空白." & Environment.NewLine
            End If
            If txtIins_IINS2.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人ID 不可為空白." & Environment.NewLine
            End If
            'NAME_IINS
            If NAME_IINS.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人姓名 不可為空白." & Environment.NewLine
            End If

            '選統編的 除外
            If ddl_IINS.SelectedValue <> "2" Then
                If Birtdy_IINS.Text.Trim = "" Then
                    Me.IsInputErr = True
                    strErr &= "要保人生日 不可為空白." & Environment.NewLine
                Else
                    '要保人要大於7歲
                    Dim ageIINS As Integer = B2BEntity.HasFunction.GetAgeInLaw(Me.Birtdy_IINS.GetDate(), Me.UcRocDobjbgn.GetDate())
                    If ageIINS <= 7 Then
                        Me.IsInputErr = True
                        strErr &= "要保人 應年滿7歲." & Environment.NewLine
                    End If
                End If
            End If

            '選統編的
            If ddl_IINS.SelectedValue = "2" Then
                If txtPrincipal.Text = "" Then
                    Me.IsInputErr = True
                    strErr &= "負責人姓名 不可為空白." & Environment.NewLine
                End If
            End If

            If IINS_AREA.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人郵遞區號 不可為空白." & Environment.NewLine
            End If
            If IINS_ADDR.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人地址 不可為空白." & Environment.NewLine
            End If
            '聯絡電話
            If IINS_TEL.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "要保人聯絡電話 不可為空白." & Environment.NewLine
            Else
                If IINS_TEL.Text.Trim.Length < 6 Then
                    Me.IsInputErr = True
                    strErr &= "要保人聯絡電話 格式錯誤." & Environment.NewLine
                End If
            End If
            If UcRocDobjbgn.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間起 不可為空白." & Environment.NewLine
            Else
                If UcRocDobjbgn.GetDate > Now.AddMonths(3) Then
                    Me.IsInputErr = True
                    strErr &= "出發日 不得大於系統日3個月." & Environment.NewLine
                End If
            End If

            If txtTobjbgn.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間時間起 不可為空白." & Environment.NewLine
            End If
            If UcRocDobjend.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間訖 不可為空白." & Environment.NewLine
            End If
            If txtTobjend.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "保險期間時間訖 不可為空白." & Environment.NewLine
            End If
            'TRAVEL_A
            If TRAVEL_A.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "旅遊地區 不可為空白." & Environment.NewLine
            End If

            Dim COUNCODE1 As String = ""
            '旅遊國家 組合判斷1
            If ddlCountry_A2.SelectedValue = "" Then
                Throw New Exception("請選擇旅遊國家之旅遊地區。")
            End If
            '旅遊國家 組合判斷2

            If ddlCountry_B2.SelectedValue = "" And ddlCountry_C2.SelectedValue = "" And ddlCountry_D2.SelectedValue = "" And
                ddlCountry_E2.SelectedValue = "" And ddlCountry_F2.SelectedValue = "" And ddlCountry_G2.SelectedValue = "" And ddlCountry_H2.SelectedValue = "" Then
                '只有一組 用第一組判斷就好
                COUNCODE1 = HttpUtility.HtmlEncode(ddlCountry_A2.SelectedValue.Substring(2, 3))
                If COUNCODE1 = "750" Then
                    '申根地區判斷
                    If Checkbox2.Checked = True Then
                        COUNCODE1 = "750"
                    Else
                        '改成歐洲地區
                        COUNCODE1 = "740"
                    End If

                End If
            Else
                Dim ddlA, ddlB, ddlC, ddlD, ddlE, ddlF, ddlG, ddlH As String

                If ddlCountry_A2.SelectedValue <> "" Then
                    ddlA = HttpUtility.HtmlEncode(ddlCountry_A2.SelectedValue.Substring(2, 3))
                Else
                    ddlA = ""
                End If
                If ddlCountry_B2.SelectedValue <> "" Then
                    ddlB = HttpUtility.HtmlEncode(ddlCountry_B2.SelectedValue.Substring(2, 3))
                Else
                    ddlB = ""
                End If
                If ddlCountry_C2.SelectedValue <> "" Then
                    ddlC = HttpUtility.HtmlEncode(ddlCountry_C2.SelectedValue.Substring(2, 3))
                Else
                    ddlC = ""
                End If
                If ddlCountry_D2.SelectedValue <> "" Then
                    ddlD = HttpUtility.HtmlEncode(ddlCountry_D2.SelectedValue.Substring(2, 3))
                Else
                    ddlD = ""
                End If
                If ddlCountry_E2.SelectedValue <> "" Then
                    ddlE = HttpUtility.HtmlEncode(ddlCountry_E2.SelectedValue.Substring(2, 3))
                Else
                    ddlE = ""
                End If
                If ddlCountry_F2.SelectedValue <> "" Then
                    ddlF = HttpUtility.HtmlEncode(ddlCountry_F2.SelectedValue.Substring(2, 3))
                Else
                    ddlF = ""
                End If
                If ddlCountry_G2.SelectedValue <> "" Then
                    ddlG = HttpUtility.HtmlEncode(ddlCountry_G2.SelectedValue.Substring(2, 3))
                Else
                    ddlG = ""
                End If
                If ddlCountry_H2.SelectedValue <> "" Then
                    ddlH = HttpUtility.HtmlEncode(ddlCountry_H2.SelectedValue.Substring(2, 3))
                Else
                    ddlH = ""
                End If

                If ddlA = "400" Then
                    If (ddlB <> "400" And ddlB <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlC <> "400" And ddlC <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlD <> "400" And ddlD <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlE <> "400" And ddlE <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlF <> "400" And ddlF <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlG <> "400" And ddlG <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlH <> "400" And ddlH <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    COUNCODE1 = "400"
                Else
                    If ddlB = "400" Or ddlC = "400" Or ddlD = "400" Or ddlE = "400" Or ddlF = "400" Or ddlG = "400" Or ddlH = "400" Then
                        Throw New Exception("選擇國外地區時無法再選擇台灣區域。")
                    End If
                    '全球各地
                    If ddlA = "750" Or ddlB = "750" Or ddlC = "750" Or ddlD = "750" Or ddlE = "750" Or ddlF = "750" Or ddlG = "750" Or ddlH = "750" Then
                        '申根地區判斷
                        If Checkbox2.Checked = True Then
                            '申根地區
                            COUNCODE1 = "750"
                        Else
                            '改成歐洲地區
                            COUNCODE1 = "740"
                        End If
                    Else
                        '改成全球地區
                        COUNCODE1 = "800"
                    End If
                End If

            End If
            TRAVEL_A.SelectedValue = COUNCODE1

            '外國人使用護照號碼投保
            If ddl_IINS.SelectedValue.ToString = "5" Then
                If TRAVEL_A.SelectedValue <> "400" Then
                    Me.IsInputErr = True
                    strErr &= "要保人勾選「持外國人護照投保」，旅遊國家必須在「國內」." & Environment.NewLine
                Else
                    If ContainFullWidthChar(NAME_IINS.Text.Trim) = True Then
                        Me.IsInputErr = True
                        strErr &= "要保人勾選「持外國人護照投保」，要保人姓名必須為半形." & Environment.NewLine
                    End If

                    If Not IsAllENG(NAME_IINS.Text) Then
                        Me.IsInputErr = True
                        strErr &= "要保人勾選「持外國人護照投保」，要保人姓名必須為英文." & Environment.NewLine
                    End If

                    If ddlIINS_COUNTRY.SelectedValue = "TW" Or String.IsNullOrEmpty(ddlIINS_COUNTRY.SelectedValue.ToString) Then
                        Me.IsInputErr = True
                        strErr &= "要保人勾選「持外國人護照投保」，要保人國別不可為空、或為「TW」." & Environment.NewLine
                    End If
                End If
            End If

            '旅遊地點
            If TRAVEL_BEGIN.Text.Trim = "" Then
                TRAVEL_BEGIN.Text = "台灣"
            End If

            TRAVEL_END.Text = ddlCountry_A2.SelectedItem.ToString
            If ddlCountry_B2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_B2.SelectedItem.ToString
            End If
            If ddlCountry_C2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_C2.SelectedItem.ToString
            End If
            If ddlCountry_D2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_D2.SelectedItem.ToString
            End If
            If ddlCountry_E2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_E2.SelectedItem.ToString
            End If
            If ddlCountry_F2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_F2.SelectedItem.ToString
            End If
            If ddlCountry_G2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_G2.SelectedItem.ToString
            End If
            If ddlCountry_H2.SelectedValue <> "" Then
                TRAVEL_END.Text = TRAVEL_END.Text.Trim + ddlCountry_H2.SelectedItem.ToString
            End If

            If ddlCountry_A2.SelectedValue = "" Then
                Me.IsInputErr = True
                strErr &= "旅遊國家 地區不可為空白." & Environment.NewLine
            Else
                'Dim RedArea As String = ConfigurationManager.AppSettings("RedArea")
                'Dim AreaArray As String()
                'AreaArray = RedArea.Split("|")
                'For i As Integer = 0 To AreaArray.Length - 1
                '    If TRAVEL_END.Text = (AreaArray(i)) Then
                '        strErr &= "旅遊地區為紅色警戒地區，本公司不予承保!." & Environment.NewLine
                '    End If
                'Next
            End If

            If INSMOBILE.Text.Trim = "" Then
                Me.IsInputErr = True
                strErr &= "索取電子保單 請輸入要保人行動電話." & Environment.NewLine
            End If

            If INSMOBILE.Text.Trim <> "" Then
                If IsNumeric(INSMOBILE.Text.Trim) = True Then
                    If INSMOBILE.Text.Trim.Length <> 10 Then
                        Me.IsInputErr = True
                        strErr &= "行動電話請輸入數字10碼." & Environment.NewLine
                    End If
                    If INSMOBILE.Text.Substring(0, 2).Trim <> "09" Then
                        Me.IsInputErr = True
                        strErr &= "請輸入正確 行動電話 以方便簡訊傳送." & Environment.NewLine
                    End If
                Else
                    Me.IsInputErr = True
                    strErr &= "行動電話請輸入數字." & Environment.NewLine
                End If
            End If
            Return strErr
        Catch ex As Exception
            Me.ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Function

    Function CheckPlanNTime(Chk_ISSU_1 As Object, Chk_ISSU_2 As Object, Chk_IT62 As Object, Chk_IT63 As Object, Chk_IT64 As Object, Chk_IT6E As Object) As String
        'F115060500 因投保項目而調整保險期間的天數

        Dim resultStr As String = ""
        Dim tmp_ISSU_1 As Decimal = GetDec(Chk_ISSU_1)
        Dim tmp_ISSU_2 As Decimal = GetDec(Chk_ISSU_2)
        Dim tmp_ISSU_62 As Decimal = GetDec(Chk_IT62)
        Dim tmp_ISSU_63 As Decimal = GetDec(Chk_IT63)
        Dim tmp_ISSU_64 As Decimal = GetDec(Chk_IT64)
        Dim tmp_ISSU_6E As Decimal = GetDec(Chk_IT6E)

        If tmp_ISSU_1 <= 5000000 And tmp_ISSU_2 <= 200000 And tmp_ISSU_62 <= 200000 And tmp_ISSU_63 <= 200000 And tmp_ISSU_64 <= 200000 And tmp_ISSU_6E <= 200000 Then
            If CInt(txtIQPITP_Days.Text) > 180 Then
                resultStr = "國外旅遊 保險期間不得超過180天."
            End If
        ElseIf tmp_ISSU_1 <= 10000000 And tmp_ISSU_2 <= 200000 And tmp_ISSU_62 <= 200000 And tmp_ISSU_63 <= 200000 And tmp_ISSU_64 <= 200000 And tmp_ISSU_6E <= 200000 Then
            If CInt(txtIQPITP_Days.Text) > 90 Then
                resultStr = "國外旅遊 保險期間不得超過90天."
            End If
        Else
            If CInt(txtIQPITP_Days.Text) > 60 Then
                resultStr = "國外旅遊 保險期間不得超過60天."
            End If
        End If

        Return resultStr
    End Function
    Private Sub UP_BTN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UP_BTN.Click
        '轉入 Excel 檔按鈕
        Try
            '因應需求單F115030200 【特定通路可以指定要保的經辦代號】，先檢查chkAssignACC是否勾選，因為勾選or取消勾選會執行對應檢核及動作
            If chkAssignACC.Checked = False Then
                Throw New Exception("請先勾選【確認經辦代號】。（保險期間上方）")
            End If

            If Not UploadFile.PostedFile Is Nothing AndAlso UploadFile.PostedFile.ContentLength > 0 Then
                Dim pfUploadFile As HttpPostedFile = UploadFile.PostedFile
                Dim dsNew As New DataSet
                Dim vu As VirtualUser = VirtualUser.CreateVirtualUserFromConfig(VirtualUser.ConfigKeyFileServiceUser)
                CheckSNNN.Text = ""
                Try
                    vu.Impersonate()

                    '記錄暫存檔案路徑
                    FileUrl = Server.MapPath(FSGetTempFileUrl(Path.GetFileName(pfUploadFile.FileName.Trim)))

                    '上傳檔案資料
                    Dim stFile As IO.Stream = pfUploadFile.InputStream()
                    Dim buf(CInt(stFile.Length - 1)) As Byte
                    Dim n As Long = stFile.Read(buf, 0, CInt(stFile.Length))
                    Dim fi As New beFileInfo

                    fi.FileInfo.AddFileInfoRow(pfUploadFile.FileName.Trim, pfUploadFile.ContentLength, pfUploadFile.ContentType.Trim, buf)
                    Me.FSCreateContent(pfUploadFile.FileName, fi.FileInfo(0).ContentData) '將檔案依暫存檔案路徑儲存

                    '使用 SQL 語法來讀 Excel。
                    Dim strConnStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileUrl + ";Extended Properties=Excel 8.0"
                    'Dim strConnStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileUrl + ";Extended Properties=Excel 8.0"
                    Dim strSQL As String = "Select * FROM [Sheet1$]"

                    Dim oleCommand As New OleDbCommand(strSQL, New OleDbConnection(strConnStr))
                    Dim oleAdapter As New OleDbDataAdapter(oleCommand)

                    oleAdapter.Fill(dsNew, "Excel")
                Catch ex As Exception
                    Throw New Exception("無法執行匯入的動作，請確認工作表名稱必需為 Sheet1 !!" & vbNewLine &
                                        "暫存檔案路徑：" & FileUrl.Trim & vbNewLine &
                                        ex.ToString)
                Finally
                    vu.Dispose()
                End Try

                '旅遊國家 組合判斷1
                If ddlCountry_A2.SelectedValue = "" Then
                    Throw New Exception("請選擇旅遊國家之旅遊地區。")
                End If

                '旅遊地區整合
                Dim COUNCODE_ALL As String = ""
                COUNCODE_ALL = CHECK_COUNCODE1()
                TRAVEL_A.SelectedValue = COUNCODE_ALL

                Dim ChkIns As String
                ChkIns = CheckExlUP()
                If ChkIns <> String.Empty Then
                    Throw New Exception(ChkIns)
                End If

                Dim ChkTime As String
                ChkTime = CheckTime()
                If ChkTime <> String.Empty Then
                    Throw New Exception(ChkTime)
                End If

                If dsNew.Tables(0).Rows.Count <= 200 Then
#Region "Show executing time information for function [ExlToCMN()]"
                    Dim t As DateTime = DateTime.Now
                    ExlToCMN(dsNew)
                    'Label1.Text = $"Time elapsed : {DateTime.Now - t} sec(s)."
#End Region
                Else
                    Throw New Exception("無法執行匯入的動作，同一保單總保人數不得超過 200人")
                End If

            End If
        Catch ex As Exception
            Me.ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try

    End Sub

#Region "EXCEL 匯入"

    Private Sub ExlToCMN(ByVal dsExl As DataSet)
        Try
            If Not (dsExl.Tables(0).Columns.Count = 22 AndAlso
                dsExl.Tables(0).Columns(6).Caption.Trim = "電話" AndAlso dsExl.Tables(0).Columns(21).Caption.Trim = "被保人是否為孕婦? (Y/N)") Then
                Throw New Exception("匯入excel資料格式版本是否正確！！！")
            End If

            'F1141222301 指定通路代號，IQP依照旅遊目的國內or國外有所不同，國內:原IQP方案；國外：IQP同標準型&豪華型
            Dim IQPSPEChannel As Boolean = False
            If (checkSpecialChannel().Item1 = True AndAlso ddlTargetProject.SelectedValue.Trim = "IQP" AndAlso TRAVEL_A.SelectedValue <> "400") Or
                    (checkChannelO().Item1 = True AndAlso ddlTargetProject.SelectedValue.Trim = "IQP" AndAlso TRAVEL_A.SelectedValue <> "400") Then
                ddlTargetProject.SelectedValue = "ITP"
                txtCase.Text = String.Empty
                IQPSPEChannel = True
            End If

            Dim TargetProject As String = ddlTargetProject.SelectedValue.Trim '20250320

            If Me.StateMain.Tables("IQPISSU").Rows.Count > 0 Then
                Me.StateMain.Tables("IQPISSU").Clear()
            End If
            If Me.StateMain.Tables("IQPOBJ").Rows.Count > 0 Then
                Me.StateMain.Tables("IQPOBJ").Clear()
            End If
            'Me.StateMain.Clear()
            Dim TOTsbErr As String = String.Empty
            Me.IsInputErr = False

            '計算金額
            Dim total As Integer = 0

            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            InsType = ""
            InsDay = ""
            InsCoverage = ""
            InsMdeduct = ""
            CountryType = ""
            '計算天數
            Dim ALLsbErr As String = String.Empty
            GetDay()
            Dim iSeq As Integer = 1
            Dim iii2 As Integer = 0
            For Each drE As DataRow In dsExl.Tables(0).Rows
                Dim ds As beITP101000 = CType(Me.StateMain, beITP101000)
                Dim dr As beITP101000.IQPISSURow
                Dim age As Integer
                dr = ds.IQPISSU.NewIQPISSURow
                Dim sbErr As String = String.Empty

                If drE(0).ToString.Trim.Equals(String.Empty) And drE(1).ToString.Trim.Equals(String.Empty) And drE(2).ToString.Trim.Equals(String.Empty) Then
                    '一堆空白 跳掉
                Else
                    dr("ISEQ") = iSeq
                    '被保險人(NISSU)
                    Dim xNISSU As String = drE(0).ToString.ToUpper.Trim
                    If xNISSU.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "被保險人 不可為空白." & Environment.NewLine
                        dr("NISSU") = ""
                    Else
                        ' 將換行符號剃除 & 不可見符號剃除
                        dr("NISSU") = CommonFunc.StrFormat(xNISSU.Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " "))
                    End If

                    '身份別(1本國,3外國_居留證,5外國_護照號碼)(IISSU_TP)
                    Dim xIISSU_TP As String = drE(1).ToString.Trim
                    If xIISSU_TP.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "身份別 不可為空白." & Environment.NewLine
                        dr("IISSU_TP") = ""
                    Else
                        dr("IISSU_TP") = xIISSU_TP
                    End If

                    Dim xIISSU As String = drE(2).ToString.Trim.ToUpper
                    If xIISSU.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "被保險人ID 不可為空白." & Environment.NewLine
                    Else
                        '全形轉半形
                        xIISSU = StrConv(xIISSU, vbNarrow)
                    End If

                    '先檢查ID是否純英、數，再檢查ID是否符合規則
                    If IsAllENGnINT(xIISSU) = False Then
                        sbErr &= "被保險人ID 格式錯誤." & Environment.NewLine
                    Else
                        '本國人與外國人
                        dr("IISSU") = xIISSU
                        If dr("IISSU_TP") = "1" Then
                            If B2BEntity.HasFunction._CheckCitizenID(xIISSU) Then

                            Else
                                Me.IsInputErr = True
                                sbErr &= "被保險人身分證號 錯誤.請查看被保人名單備註欄描述進行資料修正" & Environment.NewLine
                            End If
                        ElseIf dr("IISSU_TP") = "3" Then
                            If B2BEntity.HasFunction._CheckPassportID(xIISSU) Then

                            Else
                                Me.IsInputErr = True
                                sbErr &= "外國人 被保險人證號 錯誤.請查看被保人名單備註欄描述進行資料修正" & Environment.NewLine
                            End If
                        ElseIf dr("IISSU_TP") = "5" Then
                            '身份別為「5護照號碼」時，旅遊國家必須在「國內」
                            If TRAVEL_A.SelectedValue <> "400" Then
                                Me.IsInputErr = True
                                sbErr &= "身份別為「5護照號碼」 旅遊國家必須在「國內」" & Environment.NewLine
                            End If

                            '檢核被保人姓名必須為半形
                            If ContainFullWidthChar(dr("NISSU")) = True Then
                                Me.IsInputErr = True
                                sbErr &= "身份別為「5護照號碼」 被保險人姓名必須為半形." & Environment.NewLine
                            End If

                            '檢核被保人姓名必須為英文
                            If Not IsAllENG(dr("NISSU")) Then
                                Me.IsInputErr = True
                                sbErr &= "身份別為「5護照號碼」 被保險人姓名必須為英文." & Environment.NewLine
                            End If

                            '檢核被保人ID是否與護照號碼相同
                            If xIISSU = StrConv(drE(20).ToString.Trim.ToUpper, vbNarrow) Then
                                dr("PASSPORT_NUMBER") = StrConv(drE(20).ToString.Trim, vbNarrow)
                            Else
                                Me.IsInputErr = True
                                sbErr &= "身份別為「5護照號碼」 身份證號必須與護照號碼相同." & Environment.NewLine
                            End If
                        Else
                            Me.IsInputErr = True
                            sbErr &= "身分別錯誤 只能輸入1、3或5" & Environment.NewLine
                        End If
                    End If

                    Dim iii As Integer = 0
                    For Each drD As DataRow In dsExl.Tables(0).Select("身分證號 = '" & dr("IISSU") & "' ")
                        If drD("身分證號") = dr("IISSU") Then
                            iii = iii + 1
                        End If
                    Next
                    If iii > 1 Then
                        Me.IsInputErr = True
                        sbErr &= "[被保險人身分證號 重複].請查看被保人名單備註欄描述進行資料修正" & Environment.NewLine
                    End If
                    '塞主被保人
                    If iSeq = 1 Then
                        txtMNIISSU.Text = dr("NISSU")
                        txtMIISSU.Text = dr("IISSU")

                        If dr("IISSU_TP") = "5" Then
                            Me.IsInputErr = True
                            sbErr &= "主被保險人必須為本國人或擁有居留證之外國人（第一筆被保險人預設為主被保險人）" & Environment.NewLine
                        End If

                    End If

                    '性別(IISSU_SEX) 20250326 | 20250505 新增判斷舊式居留證 | 20251618 身份別為5護照號碼，則性別預設為1
                    If dr("IISSU_TP") = "5" Then
                        dr("IISSU_SEX") = "1"
                    Else
                        If xIISSU.Length = 10 Then
                            dr("IISSU_SEX") = If(Regex.IsMatch(xIISSU.Substring(1, 1), "1|8|A|C"), "1", If(Regex.IsMatch(xIISSU.Substring(1, 1), "2|9|B|D"), "2", ""))
                        Else
                            dr("IISSU_SEX") = ""
                        End If
                    End If

                    '生日(NISSUDB)
                    Dim xNISSUDB As String = drE(3).ToString.Trim
                    Try
                        If xNISSUDB.Equals(String.Empty) Then
                            Me.IsInputErr = True
                            sbErr &= dr("NISSU") & " " & xNISSUDB & "生日 不可為空白. 或是請將生日欄位改成[文字格式] " & Environment.NewLine
                            dr("NISSUBD") = ""
                        Else
                            If xNISSUDB.Length = 7 Then

                                Dim strYYY As String = (CInt(xNISSUDB.Substring(0, 3)) + 1911).ToString("000")
                                Dim strMM As String = CInt(xNISSUDB.Substring(3, 2)).ToString("00")
                                Dim strDD As String = CInt(xNISSUDB.Substring(5, 2)).ToString("00")

                                If IsDate(strYYY + "/" + strMM + "/" + strDD) = True Then
                                    If CDate(strYYY + "/" + strMM + "/" + strDD) > Today() Then
                                        Me.IsInputErr = True
                                        sbErr &= dr("NISSU") & " " & xNISSUDB & "出生日期 不得小於輸入日期." & Environment.NewLine
                                    End If
                                    dr("NISSUBD") = CDate(strYYY + "/" + strMM + "/" + strDD)
                                    age = B2BEntity.HasFunction.GetAgeInLaw(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate())

                                Else
                                    Me.IsInputErr = True
                                    sbErr &= dr("NISSU") & " " & xNISSUDB & "生日日期 格式錯誤" & Environment.NewLine
                                End If
                            ElseIf xNISSUDB.Length = 6 Then
                                Dim strYYY As String = (CInt(xNISSUDB.Substring(0, 2)) + 1911).ToString("000")
                                Dim strMM As String = CInt(xNISSUDB.Substring(2, 2)).ToString("00")
                                Dim strDD As String = CInt(xNISSUDB.Substring(4, 2)).ToString("00")

                                If IsDate(strYYY + "/" + strMM + "/" + strDD) = True Then
                                    If CDate(strYYY + "/" + strMM + "/" + strDD) > Today() Then
                                        Me.IsInputErr = True
                                        sbErr &= dr("NISSU") & " " & xNISSUDB & "出生日期 不得小於輸入日期." & Environment.NewLine
                                    End If
                                    dr("NISSUBD") = CDate(strYYY + "/" + strMM + "/" + strDD)
                                    age = B2BEntity.HasFunction.GetAgeInLaw(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate())

                                Else
                                    Me.IsInputErr = True
                                    sbErr &= dr("NISSU") & " " & xNISSUDB & "生日日期 格式錯誤" & Environment.NewLine
                                    dr("NISSUBD") = ""
                                End If

                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("NISSU") & " " & xNISSUDB & "生日日期 格式錯誤" & Environment.NewLine
                                dr("NISSUBD") = ""
                            End If
                        End If
                    Catch ex As Exception
                        Me.IsInputErr = True
                        sbErr &= dr("NISSU") & " " & xNISSUDB & "生日日期 格式錯誤" & Environment.NewLine
                        dr("NISSUBD") = ""
                    End Try

                    '郵遞區號(IISSU_ZIP)
                    Dim xIISSU_ZIP As String = drE(4).ToString.Trim
                    If xIISSU_ZIP.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "郵遞區號不可為空白." & Environment.NewLine
                        dr("IISSU_ZIP") = ""
                    Else
                        If xIISSU_ZIP.Length > 3 Then
                            Me.IsInputErr = True
                            sbErr &= "郵遞區號請輸入3碼." & Environment.NewLine
                            dr("IISSU_ZIP") = xIISSU_ZIP
                        Else
                            dr("IISSU_ZIP") = xIISSU_ZIP
                        End If
                    End If

                    '地址(AISSU)
                    Dim xAISSU As String = drE(5).ToString.Trim
                    If xAISSU.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "地址不可為空白." & Environment.NewLine
                        dr("AISSU") = ""

                    Else
                        ' 將換行符號剃除 & 不可見符號剃除
                        dr("AISSU") = CommonFunc.StrFormat(xAISSU)
                    End If

                    '電話(NISSU_TEL)
                    Dim xNISSU_TEL As String = drE(6).ToString.Trim
                    If xNISSU_TEL.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "電話不可為空白." & Environment.NewLine
                        dr("NISSU_TEL") = ""
                    ElseIf IsMobileNumber(xNISSU_TEL) = False Then
                        Me.IsInputErr = True
                        sbErr &= "電話 請填行動電話." & Environment.NewLine
                    Else
                        dr("NISSU_TEL") = xNISSU_TEL
                    End If


                    'eMail(NISSUMAIL)
                    Dim xNISSUMAIL As String = drE(7).ToString.Trim
                    If xNISSUMAIL.Equals(String.Empty) Then
                        'Me.IsInputErr = True
                        'strErr &= "MAIL不可為空白."
                        dr("NISSUMAIL") = ""
                    Else
                        dr("NISSUMAIL") = xNISSUMAIL
                    End If

                    '與要保人之關係(NINSREL)
                    Dim xNINSREL As String = drE(8).ToString.Trim
                    If xNINSREL.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "與要保人之關係不可為空白." & Environment.NewLine
                        dr("NINSREL") = ""
                    Else
                        dr("NINSREL") = xNINSREL

                        If dr("NINSREL") = "1" Then
                            iii2 = iii2 + 1
                        End If

                        If (iii2 > 1 And dr("NINSREL") = "1") Then
                            Me.IsInputErr = True
                            sbErr &= "[與要保人之關係 [本人] 不可重複]." & Environment.NewLine
                        End If
                        If xNINSREL.Length > 1 Then
                            Me.IsInputErr = True
                            sbErr &= "[與要保人之關係 請輸入代碼." & Environment.NewLine
                        End If

                        If xIISSU = txtIins_IINS2.Text.Trim.ToUpper And dr("NINSREL") <> "1" Then
                            Me.IsInputErr = True
                            sbErr &= "[被保人]為[要保人]本人時，[與要保人之關係]必須填[1] (本人). " & Environment.NewLine
                        End If

                    End If
                    If dr("NINSREL") = "1" Then

                        If txtIins_IINS2.Text.Trim <> dr("IISSU") Then
                            Me.IsInputErr = True
                            sbErr &= "與要保人關係選擇[本人]時，要保人與被保人[身分證字號]要相同. " & Environment.NewLine

                        End If
                        If NAME_IINS.Text.Trim <> dr("NISSU") Then
                            Me.IsInputErr = True
                            sbErr &= "與要保人關係選擇[本人]時，要保人與被保人[姓名]要相同. " & Environment.NewLine
                        End If
                        'Birtdy_IINS
                        If Birtdy_IINS.DateText <> dr("NISSUBD") Then
                            Me.IsInputErr = True
                            sbErr &= "與要保人關係選擇[本人]時，要保人與被保人[生日]要相同. " & Environment.NewLine
                        End If

                    End If
                    CommonFunc.Relationship_exltocmm(dsExl.Tables(0), "".ToString.Trim, dr, txtIins_IINS2, IsInputErr, sbErr)
                    '受益人姓名(NAMEREL)
                    Dim xNAMEREL As String = drE(9).ToString.Trim
                    If xNAMEREL.Equals(String.Empty) Then
                        dr("NAMEREL") = "法定繼承人" + vbCrLf
                    Else
                        If xNAMEREL = "無身故保險金" Then
                            dr("NAMEREL") = "法定繼承人"
                        Else
                            dr("NAMEREL") = xNAMEREL
                        End If
                    End If

                    '受益人身分証號(IDREL)
                    Dim xIDREL As String = drE(10).ToString.Trim
                    If xIDREL.Equals(String.Empty) Then
                        'Me.IsInputErr = True
                        'strErr &= "受益人身分証號不可為空白."
                        dr("IDREL") = ""
                    Else
                        If dr("NAMEREL") = "法定繼承人" Then
                            dr("IDREL") = ""
                        Else
                            If B2BEntity.HasFunction._CheckCitizenID(xIDREL) Then
                                dr("IDREL") = xIDREL
                            ElseIf B2BEntity.HasFunction._CheckPassportID(xIDREL) Then
                                dr("IDREL") = xIDREL
                            Else
                                Me.IsInputErr = True
                                sbErr &= "受益人身分證號 錯誤." & Environment.NewLine
                            End If
                        End If
                    End If

                    '與被保險人之關係(TYPEREL)
                    Dim xTYPEREL As String = drE(11).ToString.Trim
                    If xTYPEREL.Equals(String.Empty) Then
                        'Me.IsInputErr = True
                        'strErr &= "與被保險人之關係不可為空白."
                        dr("TYPEREL") = ""
                    Else
                        '無身故保險金
                        If dr("NAMEREL") = "法定繼承人" Then
                            dr("TYPEREL") = ""
                        Else
                            dr("TYPEREL") = xTYPEREL
                        End If

                    End If

                    Dim xIPLAN As String = drE(17).ToString.ToUpper.Trim

                    'F1141222301 指定通路代號，IQP依照旅遊目的國內or國外有所不同，國內:原IQP方案；國外：IQP同標準型&豪華型
                    'If IQPSPEChannel Then
                    '    xIPLAN = "L"
                    'End If
                    '指定通路代號，選IQP，xIPLAN設"L"，
                    If IQPSPEChannel AndAlso String.IsNullOrWhiteSpace(xIPLAN) Then
                        TargetProject = "IQP"
                        ddlTargetProject.SelectedValue = "IQP"
                        txtCase.Text = "E 不選擇方案"
                    End If

                    If TargetProject = "ITP" AndAlso (xIPLAN = "F" Or xIPLAN = "G") Then
                        Me.IsInputErr = True
                        sbErr &= "方案代號不得為F G ，方案代號錯誤." & Environment.NewLine
                        '20250124 方案不再控制內直接踢掉,往後處理會造成麻煩（由於修改是單一筆但檢核方案確是一整批(要相同方案)）;多筆改到起痟。好幾代前可以選不同方案
                        Throw New Exception("方案代號不得為F G ，方案代號錯誤.")
                    End If

                    If (checkSpecialChannel().Item1 = False AndAlso checkChannelO().Item1 = False) AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "J" Or xIPLAN = "K" Or xIPLAN = "L") Then
                        Me.IsInputErr = True
                        sbErr &= "方案代號不得為J K L ，方案代號錯誤." & Environment.NewLine
                        Throw New Exception("方案代號不得為J K L ，方案代號錯誤.")
                    End If

                    '指定通路代號，選IQP，xIPLAN設"L"，
                    'If (checkSpecialChannel().Item1 = False AndAlso checkChannelO().Item1 = False) AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "J" Or xIPLAN = "K") Then
                    '    Me.IsInputErr = True
                    '    sbErr &= "方案代號不得為J K L ，方案代號錯誤." & Environment.NewLine
                    '    Throw New Exception("方案代號不得為J K L ，方案代號錯誤.")
                    'End If

                    If checkChannelO().Item1 = False AndAlso TargetProject = "ITP" AndAlso xIPLAN = "O" Then
                        Me.IsInputErr = True
                        sbErr &= "方案代號不得為O ，方案代號錯誤." & Environment.NewLine
                        Throw New Exception("方案代號不得為O ，方案代號錯誤.")
                    End If

                    '意外死亡及失能保額(萬) 
                    Dim xISSU_1 As String = drE(12).ToString.Trim
                    If xISSU_1.Equals(String.Empty) Then
                        If age < 15 Then
                            If dr("IISSU_TP") = "5" Then
                                Me.IsInputErr = True
                                sbErr &= "國內地區喪葬費用保額不可為0" & Environment.NewLine
                            Else
                                '15歲以下可以不保意外死亡及失能
                                dr("ISSU_1") = "0"
                            End If
                        Else
                            Me.IsInputErr = True
                            sbErr &= "意外死亡及失能保額不可為空白." & Environment.NewLine
                            dr("ISSU_1") = ""
                        End If
                    Else
                        If GetDec(xISSU_1) > 1500 Then
                            Me.IsInputErr = True
                            sbErr &= "僅承接身故失能保險金最高為1500萬。." & Environment.NewLine
                            dr("ISSU_1") = xISSU_1
                        Else
                            dr("ISSU_1") = xISSU_1
                        End If
                    End If

                    '傷害醫療保額(萬)
                    Dim xISSU_2 As String = drE(13).ToString.Trim
                    If xISSU_2.Equals(String.Empty) Then
                        dr("ISSU_2") = "0"
                    Else
                        dr("ISSU_2") = xISSU_2
                    End If

                    Dim xISSU_3 As String = drE(14).ToString.Trim
                    Dim xISSU_3D As String = drE(15).ToString.Trim
                    Dim xISSU_IT64 As String = drE(15).ToString.Trim
                    Dim xISSU_IT44 As String = drE(16).ToString.Trim
                    If TRAVEL_A.SelectedValue = "400" Then
                        '國內團 沒有海突發
                        '海外突發疾病醫療保額(萬)
                        If GetDec(xISSU_3) <> 0 Or GetDec(xISSU_3D) <> 0 Or GetDec(xISSU_IT64) <> 0 Then
                            Me.IsInputErr = True
                            sbErr &= "出發地點為國內 不得承保海外突發疾病醫療保額." & Environment.NewLine
                        End If

                        dr("ISSU_3") = "0"
                        dr("ISSU_3D") = "0"
                        dr("ISSU_IT64") = "0"
                    Else
                        If (checkSpecialChannel().Item1 = True AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "J" Or xIPLAN = "K" Or xIPLAN = "L")) Or
                           (checkChannelO().Item1 = True AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "J" Or xIPLAN = "K" Or xIPLAN = "L" Or xIPLAN = "O")) Or
                           (checkChannelMN().Item1 AndAlso (xIPLAN = "M" Or xIPLAN = "N")) Then

                            If GetDec(xISSU_3) <> 0 AndAlso (xIPLAN = "J" Or xIPLAN = "K" Or xIPLAN = "L") Then
                                sbErr &= "方案代號為J K L ，不得承保【海外突發疾病醫療保額(萬) 不含傳染】." & Environment.NewLine
                            End If

                            If GetDec(xISSU_3) <> 0 And GetDec(xISSU_3D) <> 0 AndAlso (xIPLAN = "O" Or xIPLAN = "M" Or xIPLAN = "N") Then
                                sbErr &= "海外突發疾病醫療保額 〔含傳染病〕與〔不含傳染病〕僅能擇一." & Environment.NewLine
                            End If

                            '海外突發疾病住院醫療保額(萬)
                            If GetDec(xISSU_IT64) <> 0 Then
                                dr("ISSU_IT64") = xISSU_IT64
                            ElseIf xISSU_IT64.Equals(String.Empty) Then
                                dr("ISSU_IT64") = "0"
                            Else
                                dr("ISSU_IT64") = "0"
                            End If
                            dr("ISSU_3D") = "0"
                        Else
                            If GetDec(xISSU_3) <> 0 And GetDec(xISSU_3D) <> 0 Then
                                sbErr &= "海外突發疾病醫療保額 〔含傳染病〕與〔不含傳染病〕僅能擇一." & Environment.NewLine
                            End If

                            '海外突發疾病醫療保額(萬) 含傳染
                            If GetDec(xISSU_3D) <> 0 Then
                                dr("ISSU_3D") = xISSU_3D
                                dr("ISSU_IT64") = "0"
                            ElseIf xISSU_3D.Equals(String.Empty) Then
                                dr("ISSU_3D") = "0"
                            Else
                                dr("ISSU_3D") = "0"
                            End If
                            dr("ISSU_IT64") = "0"
                        End If

                        '海外突發疾病醫療保額(萬) 不含傳染
                        If GetDec(xISSU_3) <> 0 Then
                            dr("ISSU_3") = xISSU_3
                        ElseIf xISSU_3.Equals(String.Empty) Then
                            dr("ISSU_3") = "0"
                        Else
                            dr("ISSU_3") = "0"
                        End If
                    End If

                    '重大燒燙傷保險金(萬)
                    If GetDec(xISSU_IT44) <> 0 Then
                        '移除鎖控重大燒燙傷最低要保25萬(20170503)
                        dr("ISSU_IT44") = xISSU_IT44
                        If CInt(dr("ISSU_IT44").ToString.Trim) < 1 Then
                            Me.IsInputErr = True
                            sbErr &= "重大燒燙傷保險金 最低保額需為1萬." & Environment.NewLine
                        End If
                    ElseIf xISSU_IT44.Equals(String.Empty) Or xISSU_IT44.Equals("0") Then
                        'Me.IsInputErr = True
                        'strErr &= "重大燒燙傷保險金."
                        dr("ISSU_IT44") = "0"

                    End If

                    '方案
                    If TargetProject = "ITP" Then '20250320 
                        If xIPLAN.Equals(String.Empty) Then
                            Me.IsInputErr = True
                            sbErr &= "旅遊不便險方案不可為空白." & Environment.NewLine
                            dr("iPlanNo") = ""
                        Else
                            dr("iPlanNo") = xIPLAN
                            '主被保人的方案為主
                            If iSeq = 1 Then
                                txtCase.Text = GetiPlanNo(dr("iPlanNo").Trim) ' dr("iPlanNo")
                                lblTargetPlan.Text = txtCase.Text
                                'GetiPlanNo 串出中文 GetiPlanNo(dr("iPlanNo"))
                                Dim dobj As DiseaseFields = New DiseaseFields()

                                If (checkSpecialChannel().Item1 = True AndAlso (dr("iPlanNo").Trim = "J" Or dr("iPlanNo").Trim = "K" Or dr("iPlanNo").Trim = "L")) Or
                                   (checkChannelO().Item1 = True AndAlso (dr("iPlanNo").Trim = "J" Or dr("iPlanNo").Trim = "K" Or dr("iPlanNo").Trim = "L")) Or
                                   (checkChannelMN().Item1 AndAlso (dr("iPlanNo").Trim = "M" Or dr("iPlanNo").Trim = "N")) Then
                                    dobj.IT61 = False
                                    dobj.IT63 = False
                                    dobj.IT64 = True
                                    dobj.IT6E = False
                                    dobj.IT65 = True
                                    dobj.IT66 = True
                                    dobj.IT6F = False
                                    dobj.IT6G = False
                                    dobj.IT6D = False
                                    FieldsSwtich_Disease(dobj)
                                ElseIf (checkChannelO().Item1 = True AndAlso dr("iPlanNo").Trim = "O") Then
                                    dobj.IT61 = True
                                    dobj.IT63 = False
                                    dobj.IT64 = True
                                    dobj.IT6E = False
                                    dobj.IT65 = True
                                    dobj.IT66 = True
                                    dobj.IT6F = False
                                    dobj.IT6G = False
                                    dobj.IT6D = True
                                    FieldsSwtich_Disease(dobj)
                                Else
                                    dobj.IT61 = False
                                    dobj.IT63 = False
                                    dobj.IT64 = True
                                    dobj.IT6E = False
                                    dobj.IT65 = True
                                    dobj.IT66 = True
                                    dobj.IT6F = False
                                    dobj.IT6G = False
                                    dobj.IT6D = False
                                    FieldsSwtich_Disease(dobj)

                                    Me.ddlIT43.Enabled = False
                                    Me.ddlIT43.SelectedValue = ""
                                    Me.txtIT43.Text = ""
                                End If

                            End If
                            If GetiPlanNo(dr("iPlanNo").Trim) <> txtCase.Text.Trim Then
                                Me.IsInputErr = True
                                sbErr &= "旅遊方案需與主被保人相同." & Environment.NewLine
                            End If

                            If TRAVEL_A.SelectedValue = "400" Then
                                '國內
                                If dr("iPlanNo") = "A" Or dr("iPlanNo") = "B" Then
                                Else
                                    Me.IsInputErr = True
                                    sbErr &= "旅遊地區為國內 無法選擇(C)(D)國外的方案." & Environment.NewLine

                                End If
                            ElseIf TRAVEL_A.SelectedValue = "740" Then
                                '740
                                If checkSpecialChannel().Item1 = False And checkChannelO().Item1 = False Then
                                    If checkChannelMN().Item1 Then
                                        If dr("iPlanNo") = "M" Or dr("iPlanNo") = "N" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "歐洲地區 請選擇 (M)(N)的國外方案." & Environment.NewLine
                                        End If
                                    Else
                                        If dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "歐洲地區 請選擇 (C)(D)的國外方案." & Environment.NewLine
                                        End If
                                    End If
                                Else
                                    If checkSpecialChannel().Item1 = True Then
                                        If dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "歐洲地區 請選擇 (C)(D)(J)(K)的國外方案." & Environment.NewLine
                                        End If
                                    ElseIf checkChannelO().Item1 = True Then
                                        If dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Or dr("iPlanNo") = "O" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "歐洲地區 請選擇 (C)(D)(J)(K)(O)的國外方案." & Environment.NewLine
                                        End If
                                    End If
                                End If
                            ElseIf TRAVEL_A.SelectedValue = "750" Then
                                '申根
                                If checkSpecialChannel().Item1 = False And checkChannelO().Item1 = False Then
                                    If checkChannelMN().Item1 Then
                                        If dr("iPlanNo") = "N" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "申根地區 請選擇 (N)海外特選型方案." & Environment.NewLine
                                        End If
                                    Else
                                        If dr("iPlanNo") = "D" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "申根地區 請選擇 (D)海外超值型方案." & Environment.NewLine
                                        End If
                                    End If
                                Else
                                    If checkSpecialChannel().Item1 = True Then
                                        If dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "申根地區 請選擇 (D)(J)(K)的國外方案." & Environment.NewLine
                                        End If
                                    ElseIf checkChannelO().Item1 = True Then
                                        If dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Or dr("iPlanNo") = "O" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "申根地區 請選擇 (D)(J)(K)(O)的國外方案." & Environment.NewLine
                                        End If
                                    End If
                                End If
                            Else
                                '國外
                                If checkSpecialChannel().Item1 = False And checkChannelO().Item1 = False Then
                                    If checkChannelMN().Item1 Then
                                        If dr("iPlanNo") = "A" Or dr("iPlanNo") = "B" Then
                                            Me.IsInputErr = True
                                            sbErr &= "旅遊地區為國外 無法選擇(A)(B)國內的方案." & Environment.NewLine
                                        ElseIf dr("iPlanNo") = "M" Or dr("iPlanNo") = "N" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "旅遊方案 代碼錯誤." & Environment.NewLine
                                        End If
                                    Else
                                        If dr("iPlanNo") = "A" Or dr("iPlanNo") = "B" Then
                                            Me.IsInputErr = True
                                            sbErr &= "旅遊地區為國外 無法選擇(A)(B)國內的方案." & Environment.NewLine
                                        ElseIf dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "旅遊方案 代碼錯誤." & Environment.NewLine
                                        End If
                                    End If
                                Else
                                    If dr("iPlanNo") = "A" Or dr("iPlanNo") = "B" Then
                                        Me.IsInputErr = True
                                        sbErr &= "旅遊地區為國外 無法選擇(A)(B)國內的方案." & Environment.NewLine
                                    ElseIf (checkSpecialChannel().Item1 AndAlso (dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L")) Then
                                    ElseIf (checkChannelO().Item1 AndAlso (dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Or dr("iPlanNo") = "O")) Then
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= "旅遊方案 代碼錯誤." & Environment.NewLine
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Me.titIT6D.Attributes.Add("style", "display:none")
                        Me.lblIT6D.Attributes.Add("style", "display:none")
                        Me.lblUnitIT6D.Attributes.Add("style", "display:none")
                    End If

                    Try
                        '國別
                        Dim xIISSU_COUNTRY As String = drE(18).ToString.Trim
                        If xIISSU_COUNTRY.Equals(String.Empty) Then
                            If dr("IISSU_TP") = "5" Then
                                Me.IsInputErr = True
                                sbErr &= "身份別為「5護照號碼」 國別不得為空." & Environment.NewLine
                            Else
                                dr("IISSU_COUNTRY") = "TW"
                            End If
                        Else
                            If Not Regex.IsMatch(xIISSU_COUNTRY, "^[a-zA-Z]{2}$") Then
                                Me.IsInputErr = True
                                sbErr &= "國別 請輸入兩碼英文代碼." & Environment.NewLine
                            Else
                                If dr("IISSU_TP") = "5" AndAlso xIISSU_COUNTRY = "TW" Then
                                    Me.IsInputErr = True
                                    sbErr &= "身份別為「5護照號碼」 國別不得為「TW」." & Environment.NewLine
                                ElseIf dr("IISSU_TP") = "5" AndAlso dr("NINSREL") = "1" Then

                                    If ddlIINS_COUNTRY.SelectedValue <> xIISSU_COUNTRY Then
                                        Me.IsInputErr = True
                                        sbErr &= "[被保人]為[要保人]本人，且身份別為「5護照號碼」時 國別需相同." & Environment.NewLine
                                    Else
                                        dr("IISSU_COUNTRY") = xIISSU_COUNTRY
                                    End If

                                Else
                                    dr("IISSU_COUNTRY") = xIISSU_COUNTRY
                                End If
                            End If
                        End If

                        ' 檢查 海外旅遊時英文姓名 & 護照號碼，兩者選填，但是不可擇一填寫
                        Dim xPASSPORT_NAME As String = drE(19).ToString.Trim
                        Dim xPASSPORT_NUMBER As String = drE(20).ToString.Trim
                        Dim passportCheck As String = xPASSPORT_NAME & xPASSPORT_NUMBER

                        If Checkbox2.Checked = True And TRAVEL_A.SelectedValue = "750" Then
                            If xPASSPORT_NAME.Equals(String.Empty) Then
                                Me.IsInputErr = True
                                sbErr &= "申根區-英文姓名不可為空白." & Environment.NewLine
                                dr("PASSPORT_NAME") = ""
                            Else
                                dr("PASSPORT_NAME") = xPASSPORT_NAME
                            End If

                            If xPASSPORT_NUMBER.Equals(String.Empty) Then
                                Me.IsInputErr = True
                                sbErr &= "申根區-護照號碼不可為空白。若欄位已填寫，請確認excel的儲存格格式應為〔文字〕。" & Environment.NewLine
                                dr("PASSPORT_NUMBER") = ""
                            Else
                                dr("PASSPORT_NUMBER") = xPASSPORT_NUMBER
                            End If
                        ElseIf TRAVEL_A.SelectedValue <> "400" Then
                            If passportCheck.Equals(String.Empty) = False AndAlso (xPASSPORT_NAME.Equals(String.Empty) = True Or xPASSPORT_NUMBER.Equals(String.Empty) = True) Then
                                Me.IsInputErr = True
                                sbErr &= "英文姓名 與 護照號碼 不可擇一填寫。若欄位皆已填寫，請確認excel的儲存格格式應為〔文字〕。" & Environment.NewLine
                                dr("PASSPORT_NAME") = ""
                                dr("PASSPORT_NUMBER") = ""
                            Else
                                dr("PASSPORT_NAME") = xPASSPORT_NAME
                                dr("PASSPORT_NUMBER") = xPASSPORT_NUMBER
                            End If
                        End If

                        '是否懷孕
                        Dim xEMP As String = drE(21).ToString.Trim
                        If xEMP.Equals(String.Empty) Then
                            dr("EMP") = "N"
                        Else
                            dr("EMP") = xEMP.ToUpper

                            If dr("EMP") = "Y" Or dr("EMP") = "N" Then
                                If dr("IISSU_SEX") = "1" And dr("EMP") = "Y" Then
                                    Me.IsInputErr = True
                                    sbErr &= "被保人是否懷孕 性別錯誤." & Environment.NewLine
                                End If
                            Else
                                Me.IsInputErr = True
                                sbErr &= "被保人是否懷孕 欄位址能輸入 Y跟N." & Environment.NewLine
                            End If
                        End If
                        '懷孕不能保海突 dr("ISSU_3") 
                        If dr("EMP") = "Y" Then
                            If dr("ISSU_3").ToString.Trim <> "" And (dr("ISSU_3").ToString.Trim <> "0" Or dr("ISSU_3D").ToString.Trim <> "0" Or dr("ISSU_IT64").ToString.Trim <> "0") Then
                                Me.IsInputErr = True
                                sbErr &= "被保人為孕婦 建議不投保海外突發保險金." & Environment.NewLine
                            End If
                        End If

                    Catch ex As Exception
                        Throw New Exception("匯入格式錯誤，請於[範例下載]使用新版匯入格式")
                    End Try

                    If sbErr <> "" Then
                        InsertUP_BTN.Text = "X"
                        sbErrAll = "資料有誤, 請修正後重新上傳"
                        Me.UIAppendError(sbErr)
                    End If

                    '檢查年齡到底是幾歲
                    '判斷被保人年齡
                    'Dim age As Integer = B2BEntity.HasFunction.GetAgeInLaw(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate())
                    If age < 15 Then
                        InsType = "IT4Y"
                    Else
                        InsType = "IT41"
                    End If
                    InsDay = txtIQPITP_Days.Text.Trim
                    If dr("ISSU_1") > "0" Then
                        dr("ISSU_1") = CInt(CDec(dr("ISSU_1")) * 10000)
                    End If

                    If sbErr <> "" Then
                        sbErr &= " " & dr("IISSU") & Environment.NewLine
                    End If

                    TOTsbErr += sbErr

                    InsCoverage = dr("ISSU_1")
                    InsMdeduct = "0"

                    '5.CountryType  代表費率別 ：3=個人費率、4=團費率
                    If CHK5.Checked = True Then
                        CountryType = "4"
                    Else
                        CountryType = "3"
                    End If
                    'ITP 不便險取保費的方式 
                    'IT41/IT4Y/IT42/IT62 這四檔 旅平的相關險種
                    '1:InsType(險種)
                    '2:InsDay(天數)
                    '3.InsCoverage 保額 元
                    '4.InsMdeduct 自負額 元
                    '5.CountryType  代表費率別 ：3=個人費率、4=團費率
                    'Dflag(要抓該費率檔內小於系統日的最大日)
                    Dim sumsum1 As Integer = 0
                    Dim sumsum2 As Integer = 0
                    Dim sumsum3 As Integer = 0
                    Dim sumsum4 As Integer = 0
                    Dim sumsum5 As Integer = 0
                    Dim sumsum6 As Integer = 0 'IT43
                    Dim sumsum7 As Integer = 0 'IT64
                    Dim sumsum8 As Integer = 0 'IT65 
                    Dim sumsum9 As Integer = 0 'IT66
                    Dim sumsum10 As Integer = 0 'IT6D
                    Dim sumsum11 As Integer = 0 'IT6F
                    Dim sumsum12 As Integer = 0 'IT6G

                    '算金額1

                    Dim dsMoney1 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                    If dsMoney1.Tables("ITPPrem").Rows.Count > 0 Then
                        sumsum1 = CInt(dsMoney1.Tables("ITPPrem").Rows(0)("PREM"))

                        If checkSpecialChannel().Item1 = True AndAlso TargetProject = "ITP" AndAlso age >= 15 AndAlso (xIPLAN = "J" Or xIPLAN = "K" Or xIPLAN = "L") Then
                            '推算[大眾運輸險]的保額
                            InsType = "IT43"
                            dr("ISSU_IT43") = CInt(CDec(dr("ISSU_1")) * 0.2)
                            InsCoverage = dr("ISSU_IT43")
                            Dim dsMoney6 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                            If dsMoney6.Tables("ITPPrem").Rows.Count > 0 Then
                                sumsum6 = CInt(dsMoney6.Tables("ITPPrem").Rows(0)("PREM"))
                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("IISSU") + "大眾運輸意外保險金(萬) -- 推算發生錯誤." & Environment.NewLine
                            End If
                        End If
                    Else
                        If age < 15 Then
                            If dr("ISSU_1") = "0" Then
                                sumsum1 = 0
                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("IISSU") + "投保喪葬費用保險金最高建議保額為69萬-- 匯入金額錯誤不符." & Environment.NewLine
                            End If
                        Else
                            '不可能沒資料 沒資料就是有問題
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "意外死亡及失能保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                        End If
                        sumsum1 = 0
                    End If

                    InsType = "IT42"
                    If dr("ISSU_2") > "0" Then
                        dr("ISSU_2") = dr("ISSU_2") + "0000"
                    End If
                    InsCoverage = dr("ISSU_2")

                    '判斷金額是否正確
                    If InsCoverage = "0" Or InsCoverage = "" Then
                        sumsum2 = 0
                    Else
                        '算金額2
                        Dim dsMoney2 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                        If dsMoney2.Tables("ITPPrem").Rows.Count > 0 Then
                            sumsum2 = CInt(dsMoney2.Tables("ITPPrem").Rows(0)("PREM"))
                        Else
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "傷害醫療保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                            sumsum2 = 0
                        End If
                    End If

                    If TRAVEL_A.SelectedValue <> "400" Then
                        InsType = "IT62"
                        If dr("ISSU_3") > "0" Then
                            dr("ISSU_3") = dr("ISSU_3") + "0000"
                        End If
                        InsCoverage = dr("ISSU_3")

                        '判斷金額是否正確
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum3 = 0
                        Else
                            '算金額3
                            Dim dsMoney3 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                            If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                                sumsum3 = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))
                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(不含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                sumsum3 = 0
                            End If
                        End If
                    Else
                        If dr("ISSU_3") > "0" Then
                            dr("ISSU_3") = dr("ISSU_3") + "0000"
                        End If
                        InsCoverage = dr("ISSU_3")
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum3 = 0
                        Else
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "國內地區無法投保 海外突發疾病醫療保額(萬)(不含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                            sumsum3 = 0
                        End If
                    End If

                    If TRAVEL_A.SelectedValue <> "400" Then

                        If TargetProject = "ITP" Then
                            InsType = "IT6E"
                        Else
                            InsType = "IT63"
                        End If

                        If dr("ISSU_3D") > "0" Then
                            dr("ISSU_3D") = dr("ISSU_3D") + "0000"
                        End If
                        InsCoverage = dr("ISSU_3D")

                        '判斷金額是否正確
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum5 = 0
                        Else
                            '算金額3
                            Dim dsMoney3 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                            If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                                sumsum5 = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))

                                'F115060600 新增IT6D，方案基本型 & 超值型才有的附加保險金
                                If TargetProject = "ITP" AndAlso (xIPLAN = "C" Or xIPLAN = "D") Then

                                    '算金額IT6F
                                    InsCoverage = (CInt(dr("ISSU_3D")) * 0.01).ToString
                                    InsType = "IT6F"
                                    dsMoney3 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                    If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                                        sumsum11 = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                        sumsum5 = 0
                                        sumsum11 = 0
                                    End If

                                    '算金額IT6G
                                    InsType = "IT6G"
                                    dsMoney3 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                    If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                                        sumsum12 = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                        sumsum5 = 0
                                        sumsum11 = 0
                                        sumsum12 = 0
                                    End If

                                    '算金額IT6D
                                    InsCoverage = (CInt(dr("ISSU_3D")) * 0.05).ToString
                                    InsType = "IT6D"
                                    dsMoney3 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                    If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                                        sumsum10 = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                        sumsum5 = 0
                                        sumsum11 = 0
                                        sumsum12 = 0
                                        sumsum10 = 0
                                    End If
                                End If

                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                sumsum5 = 0
                            End If
                        End If
                    Else
                        If dr("ISSU_3D") > "0" Then
                            dr("ISSU_3D") = dr("ISSU_3D") + "0000"
                        End If
                        InsCoverage = dr("ISSU_3D")
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum5 = 0
                        Else
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "國內地區無法投保 海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                            sumsum5 = 0
                        End If
                    End If

                    If TRAVEL_A.SelectedValue <> "400" Then
                        InsType = "IT64"
                        If dr("ISSU_IT64") > "0" Then
                            dr("ISSU_IT64") = dr("ISSU_IT64") + "0000"
                        End If
                        InsCoverage = dr("ISSU_IT64")

                        '判斷金額是否正確
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum7 = 0
                        Else
                            '算金額3
                            Dim dsMoney64 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                            If dsMoney64.Tables("ITPPrem").Rows.Count > 0 Then
                                sumsum7 = CInt(dsMoney64.Tables("ITPPrem").Rows(0)("PREM"))


                                '算金額IT65
                                InsCoverage = (CInt(InsCoverage) * 0.05).ToString

                                InsType = "IT65"
                                dsMoney64 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                If dsMoney64.Tables("ITPPrem").Rows.Count > 0 Then
                                    sumsum8 = CInt(dsMoney64.Tables("ITPPrem").Rows(0)("PREM"))
                                Else
                                    Me.IsInputErr = True
                                    sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                    sumsum7 = 0
                                    sumsum8 = 0
                                End If

                                '算金額IT66 
                                InsType = "IT66"
                                dsMoney64 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                If dsMoney64.Tables("ITPPrem").Rows.Count > 0 Then
                                    sumsum9 = CInt(dsMoney64.Tables("ITPPrem").Rows(0)("PREM"))
                                Else
                                    Me.IsInputErr = True
                                    sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                    sumsum7 = 0
                                    sumsum8 = 0
                                    sumsum9 = 0
                                End If

                                If checkChannelMN().Item1 Or (checkChannelO().Item1 = True AndAlso xIPLAN = "O") Then
                                    '算金額IT6D
                                    InsType = "IT6D"
                                    dsMoney64 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                    If dsMoney64.Tables("ITPPrem").Rows.Count > 0 Then
                                        sumsum10 = CInt(dsMoney64.Tables("ITPPrem").Rows(0)("PREM"))
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                        sumsum7 = 0
                                        sumsum8 = 0
                                        sumsum9 = 0
                                        sumsum10 = 0
                                    End If
                                End If

                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                sumsum7 = 0
                            End If
                        End If
                    Else
                        If dr("ISSU_IT64") > "0" Then
                            dr("ISSU_IT64") = dr("ISSU_IT64") + "0000"
                        End If
                        InsCoverage = dr("ISSU_IT64")
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum7 = 0
                        Else
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "國內地區無法投保 海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                            sumsum7 = 0
                        End If
                    End If

                    InsType = "IT44"
                    If dr("ISSU_IT44").ToString.Trim > "0" Then
                        dr("ISSU_IT44") = dr("ISSU_IT44") + "0000"
                    End If
                    InsCoverage = dr("ISSU_IT44").ToString.Trim
                    If InsCoverage = "0" Or InsCoverage = "" Then
                        sumsum4 = 0
                    Else
                        '算金額4
                        Dim dsMoney4 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                        If dsMoney4.Tables("ITPPrem").Rows.Count > 0 Then
                            sumsum4 = CInt(dsMoney4.Tables("ITPPrem").Rows(0)("PREM"))
                        Else
                            Me.IsInputErr = True
                            sbErr &= "重大燒燙傷保險金(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                            sumsum4 = 0
                        End If
                    End If

                    dr("ISSU_4") = sumsum1 + sumsum2 + sumsum3 + sumsum4 + sumsum5 + sumsum6 + sumsum7 + sumsum8 + sumsum9 + sumsum11 + sumsum12

                    'ITP 年齡檢核
                    Dim sSolution, IADD1, IADD2, IADD3, IADD4, IADD64, IADD43 As String

                    sSolution = dr("ISSU_1").ToString.Trim
                    IADD1 = dr("ISSU_2").ToString.Trim
                    IADD2 = dr("ISSU_3").ToString.Trim
                    IADD3 = dr("ISSU_IT44").ToString.Trim
                    IADD4 = dr("ISSU_3D").ToString.Trim
                    IADD64 = dr("ISSU_IT64").ToString.Trim
                    IADD43 = dr("ISSU_IT43").ToString.Trim

                    Dim ChkStr As String
                    If IsDate(dr("NISSUBD")) = True Then
                        If (checkSpecialChannel().Item1 = True AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "J" Or xIPLAN = "K" Or xIPLAN = "L")) Or
                           (checkChannelO().Item1 = True AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "J" Or xIPLAN = "K" Or xIPLAN = "L" Or xIPLAN = "O")) Then
                            ChkStr = CheckAge_ITP_SPEChannel(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate(), sSolution, IADD1, IADD2, dr("iPlanNo").ToString, dr("IISSU_TP"), xIPLAN)
                        Else
                            ChkStr = CheckAge_ITP(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate(), sSolution, IADD1, IADD2, dr("iPlanNo").ToString, dr("IISSU_TP"))
                        End If

                        If ChkStr <> String.Empty Then
                            Me.IsInputErr = True

                            sbErr &= ChkStr
                        End If
                    End If

                    Dim ChkStr2 As String = String.Empty
                    If IsDate(dr("NISSUBD")) = True Then
                        If (checkSpecialChannel().Item1 = True AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "J" Or xIPLAN = "K" Or xIPLAN = "L")) Or
                           (checkChannelO().Item1 = True AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "J" Or xIPLAN = "K" Or xIPLAN = "L" Or xIPLAN = "O")) Then
                            ChkStr2 = CheckiPlanNo_SPEChannel(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate(), dr("iPlanNo").ToString, sSolution, IADD1, IADD2, IADD3, IADD64, dr("IISSU_TP"), IADD43)
                        Else
                            'ChkStr2 = CheckiPlanNo(dr("iPlanNo"), sSolution, IADD1, IADD2, IADD3)
                            ChkStr2 = CheckiPlanNo(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate(), dr("iPlanNo").ToString, sSolution, IADD1, IADD2, IADD3, If(checkChannelMN().Item1 AndAlso TargetProject = "ITP", IADD64, IADD4), dr("IISSU_TP"), IADD43)
                        End If

                        If ChkStr2 <> String.Empty Then
                            Me.IsInputErr = True
                            sbErr &= ChkStr2
                        End If
                    End If

                    Dim ChkStr3 As String = CheckPlanNTime(sSolution, IADD1, IADD2, IADD4, IADD64, IADD4)
                    If ChkStr3 <> String.Empty Then
                        Me.IsInputErr = True
                        sbErr &= ChkStr3
                    End If

                    If Not sbErr.Equals(String.Empty) Then dr("ERR") = sbErr

                    ds.IQPISSU.AddIQPISSURow(dr)

                    'BaseSetChildFormData(Me.dgdMain, Me.StateMain)

                    If age < 15 Then
                        'IT4Y
                        Dim dr2 As beITP101000.IQPOBJRow
                        dr2 = ds.IQPOBJ.NewIQPOBJRow
                        dr2("IISSU") = dr("IISSU")
                        dr2("INSTyp") = If(TargetProject = "ITP", "IT4Y", "IQP0Y")
                        dr2("CONTENT") = "旅行平安保險(未滿15歲)"
                        dr2("INSAMT") = dr("ISSU_1")
                        dr2("INSPREM") = sumsum1
                        ds.IQPOBJ.AddIQPOBJRow(dr2)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)

                    Else
                        'IT41
                        Dim dr1 As beITP101000.IQPOBJRow
                        dr1 = ds.IQPOBJ.NewIQPOBJRow
                        dr1("IISSU") = dr("IISSU")
                        dr1("INSTyp") = If(TargetProject = "ITP", "IT41", "IQP01")
                        dr1("CONTENT") = "旅行平安保險[標準型]"
                        dr1("INSAMT") = GetDec(dr("ISSU_1"))
                        dr1("INSPREM") = sumsum1
                        ds.IQPOBJ.AddIQPOBJRow(dr1)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)

                    End If
                    'IT42
                    If sumsum2 <> 0 Then
                        Dim dr3 As beITP101000.IQPOBJRow
                        dr3 = ds.IQPOBJ.NewIQPOBJRow
                        dr3("IISSU") = dr("IISSU")
                        dr3("INSTyp") = If(TargetProject = "ITP", "IT42", "IQP02")
                        dr3("CONTENT") = "傷害醫療保險給付[實支實付型]"
                        dr3("INSAMT") = dr("ISSU_2")
                        dr3("INSPREM") = sumsum2
                        ds.IQPOBJ.AddIQPOBJRow(dr3)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT44
                    If sumsum4 <> 0 Then
                        Dim dr44 As beITP101000.IQPOBJRow
                        dr44 = ds.IQPOBJ.NewIQPOBJRow
                        dr44("IISSU") = dr("IISSU")
                        dr44("INSTyp") = If(TargetProject = "ITP", "IT44", "IQP04")
                        dr44("CONTENT") = "重大燒燙傷保險金"
                        dr44("INSAMT") = dr("ISSU_IT44")
                        dr44("INSPREM") = sumsum4
                        ds.IQPOBJ.AddIQPOBJRow(dr44)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT62
                    If sumsum3 <> 0 Then
                        Dim dr4 As beITP101000.IQPOBJRow
                        dr4 = ds.IQPOBJ.NewIQPOBJRow
                        dr4("IISSU") = dr("IISSU")
                        dr4("INSTyp") = If(TargetProject = "ITP", "IT62", "IQP11")
                        dr4("CONTENT") = "海外突發疾病醫療保險金(不含傳染病)"
                        dr4("INSAMT") = dr("ISSU_3")
                        dr4("INSPREM") = sumsum3
                        ds.IQPOBJ.AddIQPOBJRow(dr4)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT63 or IT6E
                    If sumsum5 <> 0 Then
                        Dim dr4 As beITP101000.IQPOBJRow
                        dr4 = ds.IQPOBJ.NewIQPOBJRow
                        dr4("IISSU") = dr("IISSU")
                        dr4("INSTyp") = If(TargetProject = "ITP", "IT6E", "IQP12")
                        dr4("CONTENT") = "海外突發疾病醫療保險金(含傳染病)"
                        dr4("INSAMT") = dr("ISSU_3D")
                        dr4("INSPREM") = sumsum5
                        ds.IQPOBJ.AddIQPOBJRow(dr4)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT64
                    If sumsum7 <> 0 Then
                        Dim dr64 As beITP101000.IQPOBJRow
                        dr64 = ds.IQPOBJ.NewIQPOBJRow
                        dr64("IISSU") = dr("IISSU")
                        dr64("INSTyp") = If(TargetProject = "ITP", "IT64", "IQP13")
                        dr64("CONTENT") = "海外突發疾病住院醫療保險金"
                        dr64("INSAMT") = dr("ISSU_IT64")
                        dr64("INSPREM") = sumsum7
                        ds.IQPOBJ.AddIQPOBJRow(dr64)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT65 
                    If sumsum8 <> 0 Then
                        Dim dr65 As beITP101000.IQPOBJRow
                        dr65 = ds.IQPOBJ.NewIQPOBJRow
                        dr65("IISSU") = dr("IISSU")
                        dr65("INSTyp") = "IT65"
                        dr65("CONTENT") = "海外突發疾病門診醫療保險金"
                        dr65("INSAMT") = (CInt(dr("ISSU_IT64")) * 0.05).ToString
                        dr65("INSPREM") = sumsum8
                        ds.IQPOBJ.AddIQPOBJRow(dr65)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT66
                    If sumsum9 <> 0 Then
                        Dim dr66 As beITP101000.IQPOBJRow
                        dr66 = ds.IQPOBJ.NewIQPOBJRow
                        dr66("IISSU") = dr("IISSU")
                        dr66("INSTyp") = "IT66"
                        dr66("CONTENT") = "海外突發疾病急診醫療保險金"
                        dr66("INSAMT") = (CInt(dr("ISSU_IT64")) * 0.05).ToString
                        dr66("INSPREM") = sumsum9
                        ds.IQPOBJ.AddIQPOBJRow(dr66)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT6F 
                    If sumsum11 <> 0 AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "C" Or xIPLAN = "D") Then
                        Dim dr6F As beITP101000.IQPOBJRow
                        dr6F = ds.IQPOBJ.NewIQPOBJRow
                        dr6F("IISSU") = dr("IISSU")
                        dr6F("INSTyp") = "IT6F"
                        dr6F("CONTENT") = "海外突發疾病急診醫療保險金"
                        dr6F("INSAMT") = (CInt(dr("ISSU_3D")) * 0.01).ToString
                        dr6F("INSPREM") = sumsum11
                        ds.IQPOBJ.AddIQPOBJRow(dr6F)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT6G
                    If sumsum12 <> 0 AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "C" Or xIPLAN = "D") Then
                        Dim dr6G As beITP101000.IQPOBJRow
                        dr6G = ds.IQPOBJ.NewIQPOBJRow
                        dr6G("IISSU") = dr("IISSU")
                        dr6G("INSTyp") = "IT6G"
                        dr6G("CONTENT") = "海外突發疾病門診醫療保險金"
                        dr6G("INSAMT") = (CInt(dr("ISSU_3D")) * 0.01).ToString
                        dr6G("INSPREM") = sumsum12
                        ds.IQPOBJ.AddIQPOBJRow(dr6G)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT6D
                    If sumsum10 <> 0 AndAlso TargetProject = "ITP" AndAlso (xIPLAN = "C" Or xIPLAN = "D" Or xIPLAN = "M" Or xIPLAN = "N" Or xIPLAN = "O") Then
                        Dim dr6D As beITP101000.IQPOBJRow
                        dr6D = ds.IQPOBJ.NewIQPOBJRow
                        dr6D("IISSU") = dr("IISSU")
                        dr6D("INSTyp") = "IT6D"
                        dr6D("CONTENT") = "返國繼續住院醫療保險金"
                        If checkChannelMN().Item1 Or (checkChannelO().Item1 = True AndAlso xIPLAN = "O") Then
                            dr6D("INSAMT") = (CInt(dr("ISSU_IT64")) * 0.05).ToString
                        Else
                            dr6D("INSAMT") = (CInt(dr("ISSU_3D")) * 0.05).ToString
                        End If
                        dr6D("INSPREM") = sumsum10
                        ds.IQPOBJ.AddIQPOBJRow(dr6D)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT43
                    If sumsum6 <> 0 Then
                        Dim dr5 As beITP101000.IQPOBJRow
                        dr5 = ds.IQPOBJ.NewIQPOBJRow
                        dr5("IISSU") = dr("IISSU")
                        dr5("INSTyp") = If(TargetProject = "ITP", "IT43", "IQP03")
                        dr5("CONTENT") = "大眾運輸意外保險金"
                        dr5("INSAMT") = dr("ISSU_IT43")
                        dr5("INSPREM") = sumsum6
                        ds.IQPOBJ.AddIQPOBJRow(dr5)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If

                    If TargetProject = "ITP" AndAlso xIPLAN <> "L" Then
#Region "不便險方案"
                        '不便險方案
                        Dim iPlanNo, IISSU As String
                        iPlanNo = dr("iPlanNo").ToString.Trim
                        IISSU = dr("IISSU").ToString.Trim
                        InsType = ""
                        InsDay = ""
                        InsCoverage = ""
                        InsMdeduct = ""
                        CountryType = ""

                        'ITP 不便險取保費的方式
                        '0 IISSU
                        '1:InsType(險種)
                        '2:InsDay(天數)
                        '3.InsCoverage 保額 元
                        '4.InsMdeduct 自負額 元
                        '5.CountryType 1=國內 2=國外
                        '=====IT11=====
                        InsType = "IT11"
                        InsDay = txtIQPITP_Days.Text
                        If iPlanNo = "A" Then
                            InsCoverage = "50000"
                            'ElseIf iPlanNo = "B" Then
                            '    InsCoverage = "50000"
                        ElseIf iPlanNo = "C" Or iPlanNo = "M" Or iPlanNo = "O" Then
                            InsCoverage = "200000"
                        ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                            InsCoverage = "200000"
                        ElseIf iPlanNo = "J" Then
                            InsCoverage = "3000000"
                        ElseIf iPlanNo = "K" Then
                            InsCoverage = "5000000"
                        Else
                            InsCoverage = "0"
                            '20250124 方案不再控制內直接踢掉,往後處理會造成麻煩（由於修改是單一筆但檢核方案確是一整批(要相同方案)）;多筆改到起痟。好幾代前可以選不同方案
                            Throw New Exception($"[{iPlanNo}]方案非範圍內選項。請重新修改。")
                        End If
                        InsMdeduct = "2500"
                        '2016.2.16 (5集體國內 6集體國外) CHK5
                        'ITP費率檔(個人集體直接) 1啟用 0停用 
                        If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                            If CHK5.Checked = True Then
                                If TRAVEL_A.SelectedValue = "400" Then
                                    CountryType = "5" '國外5 外6 別
                                Else
                                    CountryType = "6" '國外5 外6 別
                                End If
                            Else
                                If TRAVEL_A.SelectedValue = "400" Then
                                    CountryType = "1" '國內1 外2 別
                                Else
                                    CountryType = "2" '國內1 外2 別
                                End If
                            End If
                        Else
                            If TRAVEL_A.SelectedValue = "400" Then
                                CountryType = "1" '國內1 外2 別
                            Else
                                CountryType = "2" '國內1 外2 別
                            End If
                        End If
                        Dim sumsumIT11 As Integer
                        Dim sumNewIT11 As Integer
                        sumsumIT11 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT11, "", iPlanNo)
                        sumNewIT11 = sumsumIT11

                        '=====IT2R=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Then
                            '20251218
                            '	身故500萬以上，方案J K優惠IT2R不算錢
                            Dim IT2RFree = False
                            If iPlanNo = "J" Or iPlanNo = "K" Then
                                If CType(sSolution, Int32) >= 5000000 Then IT2RFree = True
                            End If

                            InsType = "IT2R"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "1000000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "3000000"
                            ElseIf iPlanNo = "J" Or iPlanNo = "K" Then
                                InsCoverage = "2500000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT2R As Integer
                            Dim sumNewIT2R As Integer
                            If IT2RFree Then
                                sumsumIT2R = InsertIQPOBJ_EXL_Free(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT2R, "")
                            Else
                                sumsumIT2R = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT2R, "", iPlanNo)
                            End If
                            sumNewIT2R = sumsumIT2R
                        End If
                        '=====IT37=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT37"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "10000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Or iPlanNo = "O" Then
                                InsCoverage = "30000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "120000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "150000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT37 As Integer
                            Dim sumNewIT37 As Integer
                            sumsumIT37 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT37, "", iPlanNo)
                            sumNewIT37 = sumsumIT37
                        End If
                        '=====IT36=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT36"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "20000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Or iPlanNo = "O" Then
                                InsCoverage = "60000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "120000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "150000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT36 As Integer
                            Dim sumNewIT36 As Integer
                            sumsumIT36 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT36, "", iPlanNo)
                            sumNewIT36 = sumsumIT36
                        End If
                        '=====IT3A=====
                        If iPlanNo <> "A" Then
                            InsType = "IT3A"
                            If iPlanNo = "B" Then
                                InsCoverage = "1000"
                            ElseIf iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "O" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT3A As Integer
                            Dim sumNewIT3A As Integer
                            sumsumIT3A = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT3A, "", iPlanNo)
                            sumNewIT3A = sumsumIT3A
                        End If
                        '=====IT3B=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT3B"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Or iPlanNo = "O" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "3000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT3B As Integer
                            Dim sumNewIT3B As Integer
                            sumsumIT3B = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT3B, "", iPlanNo)
                            sumNewIT3B = sumsumIT3B
                        End If
                        '=====IT3C=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "M" Or iPlanNo = "N" Then
                            InsType = "IT3C"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "10000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "30000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT3C As Integer
                            Dim sumNewIT3C As Integer
                            sumsumIT3C = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT3C, "", iPlanNo)
                            sumNewIT3C = sumsumIT3C
                        End If
                        '=====IT3D=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT3D"
                            If iPlanNo = "C" Or iPlanNo = "M" Or iPlanNo = "O" Then
                                InsCoverage = "1000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "2000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT3D As Integer
                            Dim sumNewIT3D As Integer
                            sumsumIT3D = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT3D, "", iPlanNo)
                            sumNewIT3D = sumsumIT3D
                        End If
                        '=====IT3E=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "M" Or iPlanNo = "N" Then
                            InsType = "IT3E"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "1000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "2000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT3E As Integer
                            Dim sumNewIT3E As Integer
                            sumsumIT3E = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT3E, "", iPlanNo)
                            sumNewIT3E = sumsumIT3E
                        End If
                        '=====IT39=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT39"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "O" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT39 As Integer
                            Dim sumNewIT39 As Integer
                            sumsumIT39 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT39, "", iPlanNo)
                            sumNewIT39 = sumsumIT39

                        End If
                        '=====IT38=====
                        If iPlanNo <> "A" Then
                            InsType = "IT38"
                            If iPlanNo = "B" Then
                                InsCoverage = "1500"
                            ElseIf iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "O" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT38 As Integer
                            Dim sumNewIT38 As Integer
                            sumsumIT38 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT38, "", iPlanNo)
                            sumNewIT38 = sumsumIT38
                        End If
                        '=====IT2M=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT2M"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Or iPlanNo = "O" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT2M As Integer
                            Dim sumNewIT2M As Integer
                            sumsumIT2M = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT2M, "", iPlanNo)
                            sumNewIT2M = sumsumIT2M
                        End If
                        '=====IT27=====
                        If iPlanNo <> "A" And iPlanNo <> "O" Then
                            InsType = "IT27"
                            If iPlanNo = "B" Then
                                InsCoverage = "2000"
                            ElseIf iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "5000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "6000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT27 As Integer
                            Dim sumNewIT27 As Integer
                            sumsumIT27 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT27, "", iPlanNo)
                            sumNewIT27 = sumsumIT27
                        End If
                        '=====IT2D=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT2D"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "20000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "30000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "30000"
                            ElseIf iPlanNo = "K" Or iPlanNo = "O" Then
                                InsCoverage = "50000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT2D As Integer
                            Dim sumNewIT2D As Integer
                            sumsumIT2D = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT2D, "", iPlanNo)
                            sumNewIT2D = sumsumIT2D
                        End If
                        '=====IT28=====
                        InsType = "IT28"
                        If iPlanNo = "A" Or iPlanNo = "B" Or iPlanNo = "C" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "O" Then
                            InsCoverage = "3000"
                        ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                            InsCoverage = "5000"
                        ElseIf iPlanNo = "J" Then
                            InsCoverage = "3000"
                        ElseIf iPlanNo = "K" Then
                            InsCoverage = "3000"
                        End If
                        InsMdeduct = "0"

                        Dim sumsumIT28 As Integer
                        Dim sumNewIT28 As Integer
                        sumsumIT28 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT28, "", iPlanNo)
                        sumNewIT28 = sumsumIT28

                        '=====IT2G=====
                        If iPlanNo = "A" Or iPlanNo = "B" Then
                            InsType = "IT2G"
                            InsCoverage = "3000"
                            InsMdeduct = "0"

                            Dim sumsumIT2G As Integer
                            Dim sumNewIT2G As Integer
                            sumsumIT2G = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT2G, "", iPlanNo)
                            sumNewIT2G = sumsumIT2G
                        End If
                        '=====IT31=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT31"
                            If iPlanNo = "C" Or iPlanNo = "M" Or iPlanNo = "O" Then
                                InsCoverage = "50000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "50000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "30000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "50000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT31 As Integer
                            Dim sumNewIT31 As Integer
                            sumsumIT31 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT31, "", iPlanNo)
                            sumNewIT31 = sumsumIT31

                        End If
                        '=====IT34=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT34"
                            If iPlanNo = "J" Then
                                InsCoverage = "120000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "150000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT34 As Integer
                            Dim sumNewIT34 As Integer
                            sumsumIT34 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT34, "", iPlanNo)
                            sumNewIT34 = sumsumIT34

                        End If
                        '=====IT35=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT35"
                            If iPlanNo = "J" Then
                                InsCoverage = "120000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "150000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT35 As Integer
                            Dim sumNewIT35 As Integer
                            sumsumIT35 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT35, "", iPlanNo)
                            sumNewIT35 = sumsumIT35

                        End If
                        '=====IT2P=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT2P"
                            If iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT2P As Integer
                            Dim sumNewIT2P As Integer
                            sumsumIT2P = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT2P, "", iPlanNo)
                            sumNewIT2P = sumsumIT2P

                        End If
                        '=====IT2U=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT2U"
                            If iPlanNo = "J" Then
                                InsCoverage = "10000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "10000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT2U As Integer
                            Dim sumNewIT2U As Integer
                            sumsumIT2U = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT2U, "", iPlanNo)
                            sumNewIT2U = sumsumIT2U

                        End If
                        '=====IT2T=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT2T"
                            If iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT2T As Integer
                            Dim sumNewIT2T As Integer
                            sumsumIT2T = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT2T, "", iPlanNo)
                            sumNewIT2T = sumsumIT2T

                        End If
                        '=====IT2S=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT2S"
                            If iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT2S As Integer
                            Dim sumNewIT2S As Integer
                            sumsumIT2S = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, iSeq, sumNewIT2S, "", iPlanNo)
                            sumNewIT2S = sumsumIT2S

                        End If
                        ''=====IT2N=====
                        'If iPlanNo = "D" Then
                        '    InsType = "IT2N"
                        '    InsCoverage = "3000"
                        '    InsMdeduct = "0"
                        '    InsertIQPOBJ_EXL(dr("IISSU"), InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2N)
                        'End If
#End Region
                    End If

                    Dim TOTPREM As Integer = 0
                    For Each dr9 As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & dr("IISSU") & "' ")
                        TOTPREM += CInt(dr9("INSPREM"))
                    Next
                    dr("ISSU_4") = TOTPREM
                    iSeq = iSeq + 1
                    ALLsbErr &= sbErr
                End If
            Next
            BaseSetChildFormData(Me.dgdMain, Me.StateMain)
            If Me.IsInputErr Then

                InsertUP_BTN.Text = "X"
                dgdMain.Columns(0).Visible = False
                dgdMain.Columns(1).Visible = False
                dgdMain.Columns(2).Visible = False
                sbErrAll = "資料有誤, 請修正後重新上傳"
                If TargetProject = "ITP" Then txtCase.Text = GetiPlanNo(txtCase.Text)
                Me.UIAppendError(ALLsbErr)

            Else
                If StateMain.Tables("IQPISSU").Rows.Count = 0 Then
                    Me.UIAppendMessage("無資料可供轉換!")
                Else
                    'LockPage()
                    InsertUP_BTN.Text = ""
                    'F115031101'Me.StateCommandBar.CommandState.FindByNameArgument("SaveTemp", "").Enabled = True
                    Me.StateCommandBar.CommandState.FindByNameArgument("SaveTemp_New", "").Enabled = True
                    sbErrAll = ""
                    dgdMain.Columns(0).Visible = True
                    dgdMain.Columns(1).Visible = True
                    dgdMain.Columns(2).Visible = True
                    Me.UIAppendMessage("匯入資料正確!")

                    '20250522 上方確認被保人投保方案是否相同時，txtCase已經給值
                    ''顯示中文
                    'If TargetProject = "ITP" Then txtCase.Text = GetiPlanNo(txtCase.Text)
                End If
            End If

            'StateMain.AcceptChanges()
        Catch ex As Exception
            InsertUP_BTN.Text = "X"
            'Me.UIAppendError(ex.Message.ToString)
            ProcessException(ex)
        End Try
    End Sub
    '判斷匯入的Excel是否有錯
    Property IsInputErr() As Boolean
        Get
            If IsNothing(Me.ViewState("isErr")) Then
                Return False
            Else
                Return CBool(Me.ViewState("isErr"))
            End If
        End Get
        Set(ByVal Value As Boolean)
            Me.ViewState("isErr") = Value
        End Set
    End Property
    '檔案匯入路徑
    Private Property FileUrl() As String
        Get
            If Me.ViewState("FileUrl") Is Nothing Then
                Return String.Empty
            Else
                Return CType(Me.ViewState("FileUrl"), String)
            End If
        End Get

        Set(ByVal Value As String)
            If Value = String.Empty Then
                Me.ViewState.Remove("FileUrl")
            Else
                Me.ViewState("FileUrl") = Value
            End If
        End Set
    End Property
    Public Function InsertIQPOBJ_EXL(ByVal IISSU_EX As String, ByVal InsType_EX As String, ByVal InsDay_EX As String, ByVal InsCoverage_EX As String, ByVal InsMdeduct_EX As String, ByVal CountryType_EX As String, ByVal i As Integer, ByVal sumNew As Integer, ByVal OID As String, ByVal iPlanNo As String) As Integer
        Dim ds As beITP101000 = CType(Me.StateMain, beITP101000)

        'ITP 不便險取保費的方式
        '0:IISSU_EX 被保人
        '1:InsType(險種)
        '2:InsDay(天數)
        '3.InsCoverage 保額 元
        '4.InsMdeduct 自負額 元
        '5.CountryType 1=國內 2=國外

        Dim sumIT As Integer

        If iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "L" Or iPlanNo = "O" Then
            sumIT = 0
            '算金額
            Dim dsIT As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType_EX, InsDay_EX, InsCoverage_EX, InsMdeduct_EX, CountryType_EX)
            If dsIT.Tables("ITPPrem").Rows.Count > 0 Then
                sumIT = CInt(dsIT.Tables("ITPPrem").Rows(0)("PREM"))
            Else
                sumIT = 0
            End If
        Else
            If i = 1 Then
                sumIT = 0
                '算金額
                Dim dsIT As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType_EX, InsDay_EX, InsCoverage_EX, InsMdeduct_EX, CountryType_EX)
                If dsIT.Tables("ITPPrem").Rows.Count > 0 Then
                    sumIT = CInt(dsIT.Tables("ITPPrem").Rows(0)("PREM"))
                Else
                    sumIT = 0
                End If
            Else
                sumIT = sumNew
            End If
        End If


        Dim drIT As beITP101000.IQPOBJRow
        drIT = ds.IQPOBJ.NewIQPOBJRow
        drIT("IQPMainOID") = OID
        drIT("IISSU") = IISSU_EX
        drIT("INSTyp") = InsType_EX
        drIT("INSAMT") = InsCoverage_EX
        drIT("INSPREM") = sumIT
        drIT("MDEDUCT") = InsMdeduct_EX 'IT11 MDEDUCT 自負額
        ds.IQPOBJ.AddIQPOBJRow(drIT)
        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)

        Return sumIT

    End Function
    Public Function InsertIQPOBJ_EXL_Free(ByVal IISSU_EX As String, ByVal InsType_EX As String, ByVal InsDay_EX As String, ByVal InsCoverage_EX As String, ByVal InsMdeduct_EX As String, ByVal CountryType_EX As String, ByVal i As Integer, ByVal sumNew As Integer, ByVal OID As String) As Integer
        Dim ds As beITP101000 = CType(Me.StateMain, beITP101000)

        Dim sumIT As Integer = 0

        Dim drIT As beITP101000.IQPOBJRow
        drIT = ds.IQPOBJ.NewIQPOBJRow
        drIT("IQPMainOID") = OID
        drIT("IISSU") = IISSU_EX
        drIT("INSTyp") = InsType_EX
        drIT("INSAMT") = InsCoverage_EX
        drIT("INSPREM") = sumIT
        drIT("MDEDUCT") = InsMdeduct_EX 'IT11 MDEDUCT 自負額
        ds.IQPOBJ.AddIQPOBJRow(drIT)
        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)

        Return sumIT

    End Function
    '資料存檔(若存在則先刪再存)
    Public Sub FSCreateContent(ByVal fileName As String, ByVal binaryData() As Byte)
        Dim fs As FileStream
        Dim bw As BinaryWriter

        Dim vu As VirtualUser = VirtualUser.CreateVirtualUserFromConfig(VirtualUser.ConfigKeyFileServiceUser)
        vu.Impersonate()
        If File.Exists(FileUrl) Then
            File.Delete(FileUrl)
        End If

        Try
            fs = New FileStream(FileUrl, FileMode.CreateNew)
            bw = New BinaryWriter(fs)
            bw.Write(binaryData)
            bw.Close()
            fs.Close()
        Catch ex As Exception
            If Not bw Is Nothing Then bw.Close()
            If Not fs Is Nothing Then fs.Close()
        Finally
            vu.Dispose()
        End Try
    End Sub
#End Region

    '公會查詢service As System.Data.DataSet
    Function CallService(ByVal ds As beHASB2CTransQuery, ByVal myForm As Web.UI.Page)

        Dim dsTmp As New DataSet
        Try
            Dim ws As B2BWebReference.BusinessService = InitsWS_IQP(GetWsURL_IQP(CBool(ConfigurationManager.AppSettings("bHASB2CRequest"))), CDbl(600000))

            dsTmp = ws.InvokeBusinessService("HASService.bsHASB2CTrans", ds)

            ws.Dispose()
        Catch ex As Exception
            'Throw New Exception("ws系統錯誤")
            Throw ex
            'PopupMSG("系統錯誤", myForm)
        End Try
        Return dsTmp
    End Function
    Public Function InitsWS_IQP(ByVal sWsUrl As String, ByVal iTimerOut As Double)
        Dim ws As New B2BWebReference.BusinessService
        '若有傳遞URL就用傳入的URL當來源WS
        If sWsUrl <> "" Then
            ws.Url = sWsUrl
        End If

        If iTimerOut = 0 Then
            ws.Timeout = 60000
        Else
            ws.Timeout = iTimerOut
        End If
        Return ws
    End Function
    'Call Service路徑_MIS報價service
    Function GetWsURL_IQP(ByVal bFun As Boolean)
        Dim sUrl As String = ""
        Dim sWsURLT As String = ConfigurationManager.AppSettings("sWsURLT_IQP")
        Dim sWsURLP As String = ConfigurationManager.AppSettings("sWsURLP_IQP")
        sUrl = sWsURLT
        If bFun Then
            sUrl = sWsURLP
        Else
            sUrl = sWsURLT
        End If
        WebServiceURL = sUrl
        Return sUrl
    End Function

    Protected Overrides Sub BaseSetChildFormData(ByVal childDataGrid As DataGrid, ByVal mainDataSet As DataSet)
        '此功能以遞迴方式設定所有明細項目的資料到畫面上
        If Not childDataGrid Is Nothing Then
            Dim dv As DataView
            If mainDataSet Is Nothing Then
                dv = Nothing
            Else
                dv = AppGetDataGridBoundDataView(childDataGrid, mainDataSet)
            End If
            '依照ISEQ排序
            If childDataGrid.ID = "dgdMain" Then
                dv.Sort = "ISEQ"
            End If
            BaseEnsureDataGridState(childDataGrid, dv)
            childDataGrid.DataSource = dv
            childDataGrid.DataBind()
            If txtPESTATUS.Text.ToUpper = "S" And String.IsNullOrEmpty(txtPESTATUS.Text) Then
                ShowPremInfoOrNot(False)
            End If
        End If
        Dim dgds() As DataGrid = AppGetChildDataGrids(childDataGrid)
        If Not dgds Is Nothing Then
            For Each dgd As DataGrid In dgds
                BaseSetChildFormData(dgd, mainDataSet)
            Next
        End If
    End Sub
    Function GetDisplay() As String


        Dim ShowS As String = ConfigurationManager.AppSettings("bSHOW")
        Dim GetShow As String

        If ShowS = "0" Then
            GetShow = "DISPLAY: none"
        Else
            GetShow = ""
        End If

        Return GetShow
    End Function

    Protected Overrides Sub BaseOnPageLaunch(ByVal ti As IntelliSys.NetExpress.Web.TransferInfo)

        Dim basic As String = CType(StateTransfer.Args(0), String)
        If basic.Length = 14 Then
            Dim strJobID As String = basic
            Me.txtIQPMainOid.Text = strJobID
            Me.BaseDoQuery()
            If Me.txtIQPMainOid.Text.Trim <> "" Then

                Dim Rtype As String = CType(StateTransfer.Args(1), String)
                Me.txtRtype.Text = Rtype
                '資料顯示類別    E=編輯  V=預覽
                If Me.txtRtype.Text = "V" Then

                    InsertUP_BTN.Text = "X"
                    'IISSU_N.Enabled = False
                    CK2.Checked = True

                ElseIf Me.txtRtype.Text = "E" Then
                    'CheckS.Text = "1"
                    'CheckSNNN.Text = "請[重新計算保費]"
                    'ReCaculate.Enabled = True
                    CK1.Enabled = True
                    UP_BTN.Enabled = False
                Else
                    InsertUP_BTN.Text = "X"
                    'IISSU_N.Enabled = False
                    CK2.Checked = True


                End If
            End If


        End If

    End Sub
    Protected Overrides Sub BaseOnPageCall()
        Try
            Dim ti As TransferInfo = Me.StateTransfer
            Select Case ti.CommandName
                Case "SELECT"
                    Dim basic As String = CType(StateTransfer.Args(0), String)
                    If basic.Length = 14 Then
                        Dim strJobID As String = basic
                        Me.txtIQPMainOid.Text = strJobID
                        Me.BaseDoQuery()
                        If Me.txtIQPMainOid.Text.Trim <> "" Then

                            Dim Rtype As String = CType(StateTransfer.Args(1), String)
                            Me.txtRtype.Text = Rtype
                            '資料顯示類別    E=編輯  V=預覽
                            If Me.txtRtype.Text = "V" Then

                                InsertUP_BTN.Text = "X"
                                'IISSU_N.Enabled = False
                                CK2.Checked = True

                            ElseIf Me.txtRtype.Text = "E" Then
                                'CheckS.Text = "1"
                                'CheckSNNN.Text = "請[重新計算保費]"
                                'ReCaculate.Enabled = True
                                CK1.Enabled = True
                                'UP_BTN.Enabled = False
                            Else
                                InsertUP_BTN.Text = "X"
                                'IISSU_N.Enabled = False
                                CK2.Checked = True


                            End If
                        End If


                    End If

                Case Else
                    MyBase.BaseOnPageCall()
            End Select
        Catch ex As Exception
            Me.ProcessException(ex)
        End Try
    End Sub
    Protected Overrides Sub BaseOnPageReturn(ByVal ti As IntelliSys.NetExpress.Web.TransferInfo)
        Select Case ti.CommandName
            Case "SELECT"

            Case "WaitRSReport"
                If ti.Committed Then
                    Dim jobID As String = ti.Args(0).ToString
                    RSReportViewer.ViewReport(Me, jobID)
                End If
            Case "WaitFSReport"
                If ti.Committed Then
                    Dim jobID As String = ti.Args(0).ToString
                    FSReportViewer.ViewReport(Me, jobID)
                End If
            Case "NewKYC"
                '是否使用EKYC
                If CommonFunc.EKYCOrNot(Me, "BHAS101003", Me.txtCompanyid.Text) = False Then
                    Exit Select
                End If

                If ti.Committed Then
                    Dim tr As Object() = CType(ti.Result, Object())
                    If Not IsNothing(tr) Then
                        Dim TempXML = CType(tr(0), String)
                        Try
                            Dim KYCDs As DataSet = EntityHelper.DeserializeTypedXML(TempXML)

                            If IsDBNull(KYCDs) Then
                                Throw New Exception
                            End If

                            '先將KYC資料清空，避免囤積前次要保試算後的KYC結果
                            StateMain.Tables("KYCData").Clear()

                            Dim dskyc As beITP101000 = CType(Me.StateMain, beITP101000)
                            Dim drkyc As beITP101000.KYCDataRow
                            drkyc = dskyc.KYCData.NewKYCDataRow

                            drkyc("KYCXML") = TempXML
                            drkyc("HaveEKYC") = "true"
                            dskyc.KYCData.AddKYCDataRow(drkyc)
                            CheckSNNN.Text = "檢核資料正確"

                            Me.InvokeCommonService("B2BService.bsB2BSysApLog", "BHAS101003", "GetEKYC", "自EKYC頁返回並順利取KYC資料 " & txtIQPMainOid.Text)
                        Catch ex As Exception
                            Throw New Exception("EKYC存取異常，請重新執行[保費試算]，謝謝。")
                        End Try
                    Else
                        Throw New Exception("EKYC存取異常，請重新執行[保費試算]，謝謝。")
                    End If
                End If
            Case "GetFlight"
                If ti.Committed Then
                    Dim trtd As DataTable = ti.Result

                    If trtd IsNot Nothing AndAlso trtd.Rows.Count > 0 Then
                        Dim currentMaxIseq As Integer = StateMain.Tables("HASM_FLIGHT_DETAIL").AsEnumerable() _
                            .Where(Function(r) Not IsDBNull(r("ISEQ")) AndAlso r("ISEQ").ToString() <> "") _
                            .Select(Function(r) Convert.ToInt32(r("ISEQ"))) _
                            .DefaultIfEmpty(0) _
                            .Max()

                        For Each row2 As DataRow In trtd.Rows
                            currentMaxIseq += 1
                            Dim newRow As DataRow = StateMain.Tables("HASM_FLIGHT_DETAIL").NewRow()

                            For Each col As DataColumn In trtd.Columns
                                newRow(col.ColumnName) = row2(col.ColumnName)
                            Next

                            newRow("ISEQ") = currentMaxIseq
                            StateMain.Tables("HASM_FLIGHT_DETAIL").Rows.Add(newRow)
                        Next

                        BaseSetChildFormData(dgdFlight, StateMain)
                    End If

                End If
            Case "ModifyWhenY"
                If ti.Committed Then
                    Me.BaseDoQuery()
                    UIAlertMessage("請務必 重 新 套 印 要 保 文 件 !")
                End If
            Case Else
                MyBase.BaseOnPageReturn(ti)

        End Select
    End Sub
    '<summary>
    '非車系報價、出單、批改呼叫「制裁」SERVICE
    '</summary>
    '<param name="iinscls">使用的險別；  I (人身保險 ITP IRP IQP) F(火險) A(家電 BOP)</param>
    '<param name="ICUS">要/被保人id</param>
    '<param name="NCUS">要/被保人中文名稱</param>
    '<param name="FCLS">要/被保人編號別，主被保人編號別|1身分證號 2統編 3客戶代號</param>
    '<param name="DBIRTH">要/被保人生日； (yyyy/MM/dd)</param>
    '<param name="icusseq">預設00</param>
    '<param name="ICOUNTRY">要/被保人國別 (TW)</param>
    '<returns>EX:  True|''  ,  False|msg</returns>
    '<remarks></remarks>
    Function CallwsSanction(ByVal iinscls As String, ByVal ICUS As String, ByVal NCUS As String, ByVal FCLS As String, ByVal DBIRTH As DateTime, Optional ByVal icusseq As String = "00", Optional ByVal ICOUNTRY As String = "TW") As String
        iinscls = iinscls.Trim.ToUpper
        If iinscls <> "I" AndAlso iinscls <> "F" AndAlso iinscls <> "A" Then
            Return "False|險別錯誤"
        End If

        'If CheckIDNO(ICUS.Trim) = False Then
        '    Return "False|ID錯誤"
        'End If

        'If FCLS.Trim <> "1" AndAlso FCLS.Trim <> "2" AndAlso FCLS.Trim <> "3" Then
        '    Return "False|編號別錯誤"
        'End If

        Dim dsTmp As DataSet
        Dim ws As B2BWebReference.BusinessService = InitsWS_IQP(GetWsURL_IQP(CBool(ConfigurationManager.AppSettings("bHASB2CRequest"))), CDbl(ConfigurationManager.AppSettings("iWSTimerOut")))
        Dim msg As String = ""
        Try
            msg = "True|"
            dsTmp = ws.InvokeCommonService("PUBService.bsCMNgetNotCarCustCtrl", New String() {iinscls, ICUS, icusseq, NCUS, FCLS, DBIRTH.ToString("yyyy/MM/dd"), ICOUNTRY, iinscls, "ITP,IQP"})
            For i As Integer = 0 To dsTmp.Tables("NameValues").Rows.Count - 1
                If dsTmp.Tables("NameValues").Rows(i)("Name").ToString.Trim.ToUpper = "PASS" Then
                    If dsTmp.Tables("NameValues").Rows(i)("Value").ToString.Trim = "N" Then msg = "False|PASS is equal N"
                End If
            Next
        Catch ex As Exception
            msg = "False|" + ex.Message.ToString
        End Try
        Return msg

    End Function

#Region "Download Function"

    Private Function createHttpTempFile(ByVal drReportItem As beFSReports.ReportItemsRow) As String
        Dim strHashKey As String = computeHash(drReportItem)
        Dim strFileExtension As String = System.IO.Path.GetExtension(drReportItem.FileName)
        Dim strTempFileName As String = strHashKey & strFileExtension
        Dim strTempFileUrl As String = Me.FSGetTempFileUrl(strTempFileName)

        Dim vu As VirtualUser = VirtualUser.CreateVirtualUserFromConfig(VirtualUser.ConfigKeyFileServiceUser)
        vu.Impersonate()
        Dim strTempFileFullName As String = Server.MapPath(strTempFileUrl)
        vu.Dispose()

        If Not FSIsTempFileExist(strTempFileFullName) Then
            Dim binaryData As Byte()
            If drReportItem.IsFileDataNull Then
                binaryData = Me.FSReadFile(drReportItem.PathID, drReportItem.FileName)
            Else
                binaryData = drReportItem.FileData
            End If
            Dim strUrl As String = Me.FSCreateTempContent(strTempFileFullName, binaryData)
            System.Diagnostics.Debug.Assert(strUrl = strTempFileUrl)
        End If
        Return strTempFileUrl
    End Function

    Private Function computeHash(ByVal drReportItem As beFSReports.ReportItemsRow) As String
        '計算唯一值，用來防止重複建立暫存檔案
        Dim strHashKey As String
        If drReportItem.IsFileDataNull Then
            Dim dt As DateTime = FSGetFileLastWriteTime(drReportItem.PathID, drReportItem.FileName)
            strHashKey = UTF.GetMD5HexHash(drReportItem.PathID.ToString() & drReportItem.FileName & dt.ToString())    'compute unique cache id
        Else
            strHashKey = UTF.GetMD5HexHash(UTF.GetMD5HexHash(drReportItem.FileData) & drReportItem.FileName)   'compute unique cache id
        End If
        Return strHashKey
    End Function
    Private Sub ForceDownload(ByVal strPath As String, ByVal strFileName As String)
        If Request.Browser.Browser = "IE" Then 'ie瀏灠器
            '跳出視窗開啟檔案 
            FSReportViewer.ViewReport(Me, "報表", FSReadContent(strPath), FSReportViewer.ViewType.TempUrl)
        Else '非ie瀏灠器
            '提示開啟另存新檔
            Dim binaryData() As Byte = FSReadContent(strPath)
            Dim headerValue As String = "attachment; filename=" & HttpUtility.UrlEncode(System.IO.Path.GetFileName(strFileName))   '避免中文檔名問題
            HttpContext.Current.Response.ClearHeaders()
            HttpContext.Current.Response.Clear()
            HttpContext.Current.Response.AppendHeader("content-disposition", headerValue)
            HttpContext.Current.Response.BinaryWrite(binaryData)
            HttpContext.Current.Response.End()
        End If
    End Sub
#End Region

    Private Sub ddlCountry1_SelectedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlCountry_A1.SelectedIndexChanged,
        ddlCountry_B1.SelectedIndexChanged, ddlCountry_C1.SelectedIndexChanged,
        ddlCountry_D1.SelectedIndexChanged, ddlCountry_E1.SelectedIndexChanged,
        ddlCountry_F1.SelectedIndexChanged, ddlCountry_G1.SelectedIndexChanged,
        ddlCountry_H1.SelectedIndexChanged

        Dim ddl_1 As DropDownList = TryCast(sender, DropDownList)

        CheckSNNN.Text = ""
        SetCountriesToDropDownList(ddl_1)
    End Sub

    Private Sub SetCountriesToDropDownList(ddl_1 As DropDownList)
        '2025-05-12 user提IQP & ITP方案選擇不使用下拉，改透過方案選擇按鈕設定，故將可選國家最大化，待﹝保費試算﹞時檢核該方案能否選擇該國家。
        Dim Load_asia As String = "b2b.load_asia2notw_h"
        Dim ddl_2 As DropDownList = New DropDownList
        Dim dicCounItems As Dictionary(Of String, String) = New Dictionary(Of String, String)

        If ddl_1.ID = "ddlCountry_A1" Then
            Load_asia = "b2b.load_asia2h"
        End If

        dicCounItems.Add("ddlCountry_A1", "ddlCountry_A2")
        dicCounItems.Add("ddlCountry_B1", "ddlCountry_B2")
        dicCounItems.Add("ddlCountry_C1", "ddlCountry_C2")
        dicCounItems.Add("ddlCountry_D1", "ddlCountry_D2")
        dicCounItems.Add("ddlCountry_E1", "ddlCountry_E2")
        dicCounItems.Add("ddlCountry_F1", "ddlCountry_F2")
        dicCounItems.Add("ddlCountry_G1", "ddlCountry_G2")
        dicCounItems.Add("ddlCountry_H1", "ddlCountry_H2")


        If ddl_1 IsNot Nothing AndAlso dicCounItems.ContainsKey(ddl_1.ID) Then
            ddl_2 = CType(FindControl(dicCounItems(ddl_1.ID)), DropDownList)

            If ddl_1.SelectedValue = "" Then
                ddl_1.SelectedValue = ""
                ddl_2.Items.Clear()
            Else
                ddl_2.Items.Clear()
                Dim ds As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", Load_asia, ddl_1.SelectedValue.Trim)
                Dim C As New ListItem
                C.Text = ""
                C.Value = ""
                ddl_2.Items.Add(C)
                For Each dr As DataRow In ds.Tables("ITP_COUNTRY").Rows
                    Dim b As New ListItem
                    b.Value = dr("COUNCODE") + dr("TRAVEL_CODE")
                    If dr("COUNCODE") = "T1" Or dr("COUNCODE") = "T2" Or dr("COUNCODE") = "T3" Then
                        dr("COUNCODE") = "TW"
                    End If
                    b.Text = dr("COUNTRYCHA") + dr("COUNCODE")
                    ddl_2.Items.Add(b)
                Next
            End If
        End If
    End Sub

    Private Sub btMore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btMore.Click
        MoreCountryVisible(True)
        CheckSNNN.Text = ""
    End Sub

    Private Sub btMore2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btMore2.Click
        Me.ShowS22.Visible = True
        Me.ShowS23.Visible = True
        CheckSNNN.Text = ""
    End Sub

    Private Sub MoreCountryVisible(ShowOrNot As Boolean)
        Me.ShowS2.Visible = ShowOrNot

        Dim td_List As HtmlTableCell() = {Me.td_CountryB, Me.td_CountryC, Me.td_CountryD}

        For Each td As HtmlTableCell In td_List
            For Each ctrl As Control In td.Controls
                ctrl.Visible = ShowOrNot
            Next
        Next

    End Sub
    Private Sub CheckMain()
        Try
            'F115031101 改依Statemain IQPISSU的IPlanNo當作目標方案，詳述參考GetTargetProject()的說明
            Dim TargetProject As String = GetTargetProject()
            'Dim TargetProject As String = ddlTargetProject.SelectedValue.Trim() '20250320

            Dim OID As String = Me.StateMain.Tables("IQPMAIN").Rows(0)("OID").ToString.Trim '20250320 

            Dim dsCopy As DataSet
            dsCopy = StateMain.Copy

            If Me.StateMain.Tables("IQPISSU").Rows.Count > 0 Then
                Me.StateMain.Tables("IQPISSU").Clear()
            End If
            If Me.StateMain.Tables("IQPOBJ").Rows.Count > 0 Then
                Me.StateMain.Tables("IQPOBJ").Clear()
            End If

            Dim TOTsbErr As String = String.Empty
            Me.IsInputErr = False

            '計算金額
            Dim total As Integer = 0
            Dim IQPMain_totPrem As Integer = 0
            Dim age As Integer
            Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String
            InsType = ""
            InsDay = ""
            InsCoverage = ""
            InsMdeduct = ""
            CountryType = ""

            '計算天數
            Dim ALLsbErr As String = String.Empty
            GetDay()

            '若要保人為外國人，且用護照號碼投保
            If StateMain.Tables("IQPMAIN").Rows(0)("FinsCls").ToString.Trim = "5" Then
                StateMain.Tables("IQPMAIN").Rows(0)("FinsTyp") = "7"
            End If

            Dim i As Integer = 1
            Dim iii2 As Integer = 0
            Dim IISSU_Count As Integer = dsCopy.Tables("IQPISSU").Rows.Count

            For Each drE As DataRow In dsCopy.Tables("IQPISSU").Rows
                If drE.RowState <> DataRowState.Deleted Then
                    '20250320 由於之前都是參考看看無意義,往後才是真正結果★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
                    '需求 按[保費試算]依TargetProject那組狀況,產生金額
                    Dim ds As beITP101000 = CType(Me.StateMain, beITP101000)
                    Dim dr As beITP101000.IQPISSURow
                    dr = ds.IQPISSU.NewIQPISSURow
                    Dim sbErr As String = String.Empty

                    dr("ISEQ") = drE("ISEQ")
#Region "檢證============================================================================================="
                    '被保險人(NISSU)
                    If drE("NISSU").ToString.Trim.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "被保險人 不可為空白." & Environment.NewLine
                        dr("NISSU") = ""
                    Else
                        dr("NISSU") = drE("NISSU").ToString.Trim
                    End If

                    '身份別(1本國,3外國,5)(IISSU_TP)
                    If drE("IISSU_TP").ToString.Trim.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "身份別 不可為空白." & Environment.NewLine
                        dr("IISSU_TP") = ""
                    Else
                        dr("IISSU_TP") = drE("IISSU_TP").ToString.Trim
                    End If

                    '本國人與外國人
                    If dr("IISSU_TP") = "1" Then
                        If B2BEntity.HasFunction._CheckCitizenID(drE("IISSU").ToString.Trim) Then
                            dr("IISSU") = drE("IISSU").ToString.Trim.ToUpper
                        Else
                            Me.IsInputErr = True
                            dr("IISSU") = drE("IISSU").ToString.Trim
                            sbErr &= "被保險人身分證號 錯誤." & Environment.NewLine
                            'dr("IISSU") = ""
                        End If
                    ElseIf dr("IISSU_TP") = "3" Then
                        If B2BEntity.HasFunction._CheckPassportID(drE("IISSU").ToString.Trim) Then
                            dr("IISSU") = drE("IISSU").ToString.Trim
                        Else
                            Me.IsInputErr = True
                            dr("IISSU") = drE("IISSU").ToString.Trim
                            sbErr &= "外國人 被保險人證號 錯誤." & Environment.NewLine
                            'dr("IISSU") = ""
                        End If
                    ElseIf dr("IISSU_TP") = "5" Then
                        If String.IsNullOrWhiteSpace(drE("IISSU").ToString.Trim) Then
                            Me.IsInputErr = True
                            sbErr &= "外國人 被保險人護照號碼 不可為空." & Environment.NewLine
                        Else
                            dr("IISSU") = drE("IISSU").ToString.Trim
                        End If
                        If TRAVEL_A.SelectedValue <> "400" Then
                            Me.IsInputErr = True
                            sbErr &= "被保人勾選「持外國人護照投保」，旅遊國家必須在「國內」." & Environment.NewLine
                        Else
                            If ContainFullWidthChar(dr("NISSU").ToString) = True Then
                                Me.IsInputErr = True
                                sbErr &= "被保人勾選「持外國人護照投保」，被保險人姓名必須為半形." & Environment.NewLine
                            End If

                            If Not IsAllENG(dr("NISSU").ToString) Then
                                Me.IsInputErr = True
                                sbErr &= "被保人勾選「持外國人護照投保」，被保人姓名必須為英文." & Environment.NewLine
                            End If

                            If drE("IISSU_COUNTRY").ToString = "TW" Or String.IsNullOrEmpty(drE("IISSU_COUNTRY").ToString) Then
                                Me.IsInputErr = True
                                sbErr &= "被保人勾選「持外國人護照投保」，被保人國別不可為空、或為「TW」." & Environment.NewLine
                            Else
                                dr("IISSU_COUNTRY") = drE("IISSU_COUNTRY").ToString
                            End If
                        End If

                        '用護照號碼投保國內旅遊，旅遊天數至少4天
                        If GetDec(txtIQPITP_Days.Text) < 4 Then
                            Me.IsInputErr = True
                            sbErr &= "被保人勾選「持外國人護照投保」，旅遊天數至少4天" & Environment.NewLine
                        End If

                        '用護照號碼投保國內旅遊，被保險人數至少5人
                        If dsCopy.Tables("IQPISSU").Rows.Count < 5 Then
                            Me.IsInputErr = True
                            sbErr &= "被保人用護照號碼投保國內旅遊，被保險人總數至少5人" & Environment.NewLine
                        End If
                    End If

                    Dim iii As Integer = 0
                    For Each drD As DataRow In dsCopy.Tables("IQPISSU").Select("IISSU = '" & dr("IISSU") & "' ")
                        If drD("IISSU") = dr("IISSU") Then
                            iii = iii + 1
                        End If
                    Next
                    If iii > 1 Then
                        Me.IsInputErr = True
                        sbErr &= "[被保險人身分證號 重複]." & Environment.NewLine
                    End If
                    '塞主被保人
                    If drE("ISEQ") = 1 Then
                        txtMNIISSU.Text = dr("NISSU")
                        txtMIISSU.Text = dr("IISSU")
                    End If

                    '性別(IISSU_SEX)
                    If drE("IISSU_SEX").ToString.Trim.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "性別 不可為空白."
                        dr("IISSU_SEX") = ""
                    Else
                        dr("IISSU_SEX") = drE("IISSU_SEX").ToString.Trim

                        If dr("IISSU_SEX") = "1" Or dr("IISSU_SEX") = "2" Then
                            dr("IISSU_SEX") = drE("IISSU_SEX").ToString.Trim
                        Else
                            dr("IISSU_SEX") = "3" '性別其他
                        End If
                    End If

                    '生日(NISSUDB)
                    If drE("NISSUBD").ToString.Trim.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "生日 不可為空白." & Environment.NewLine
                        dr("NISSUBD") = ""
                    Else
                        dr("NISSUBD") = drE("NISSUBD")
                    End If

                    '郵遞區號(IISSU_ZIP)
                    If drE("IISSU_ZIP").ToString.Trim.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "郵遞區號不可為空白." & Environment.NewLine
                        dr("IISSU_ZIP") = ""
                    Else
                        If drE("IISSU_ZIP").ToString.Trim.Length > 3 Then
                            Me.IsInputErr = True
                            sbErr &= "郵遞區號請輸入3碼." & Environment.NewLine
                            dr("IISSU_ZIP") = drE("IISSU_ZIP").ToString.Trim
                        Else
                            dr("IISSU_ZIP") = drE("IISSU_ZIP").ToString.Trim
                        End If
                    End If

                    '地址(AISSU)
                    If drE("AISSU").ToString.Trim.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "地址不可為空白." & Environment.NewLine
                        dr("AISSU") = ""
                    Else
                        dr("AISSU") = drE("AISSU").ToString.Trim
                    End If

                    '電話(NISSU_TEL)
                    If drE("NISSU_TEL").ToString.Trim.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "電話不可為空白." & Environment.NewLine
                        dr("NISSU_TEL") = ""
                    Else
                        dr("NISSU_TEL") = drE("NISSU_TEL").ToString.Trim
                    End If

                    '30人以下才跑檢核，30以上則交由bsBHASCheckPEstatus2檢核
                    If IISSU_Count <= 30 Then
                        'A114123106)全險「通訊地址、電話」管控需求－PUB地址卡控程式呼叫說明------------------------------------------------------------------------
                        '叫用  PUBService.bsPubAddrCHK  範例
                        '(00) 0檢查用/1送審用(必)        (06)要保人郵遞區號     (12)被保人地址   
                        '(01)險別(C,M,F,A,I,E,H) (必)    (07)要保人地址      　 (13)被保人電話區碼  
                        '(02)資料來源(B2B/B2C/MIS) (必)  (08)要保人電話區碼     (14)被保人電話         　
                        '(03)報價/出單/續保/批單(必)     (09)要保人電話         　
                        '(04)PROGRANID(必)               (10)被保人統編       　
                        '(05)要保人統編                  (11)被保人郵遞區號
                        'Checkmain 
                        Dim ds99 = CallCommonService("PUBService.bsPubAddrCHK", New String() {"0", "H", "B2B", "1", "BHAS101003",
                        txtIins_IINS2.Text.Trim, IINS_AREA.Text.Trim, IINS_ADDR.Text.Trim, "", IINS_TEL.Text.Trim,
                        dr("NISSU"), dr("IISSU_ZIP"), dr("AISSU"), "", dr("NISSU_TEL")})
                        If ds99.Tables("RESULT").Rows(0)("PASS") = "N" AndAlso ds99.Tables("ERROR").Rows.Count > 0 Then
                            'UIAppendMessage(ds99.Tables("ERROR").Rows(0)("NAME"))
                            Me.IsInputErr = True
                            sbErr &= ds99.Tables("ERROR").Rows(0)("NAME")
                        End If
                        'A114123106)全險「通訊地址、電話」管控需求－PUB地址卡控程式呼叫說明------------------------------------------------------------------------
                    End If

                    'eMail(NISSUMAIL)
                    If drE("NISSUMAIL").ToString.Trim.Equals(String.Empty) Then
                        'Me.IsInputErr = True
                        'strErr &= "MAIL不可為空白."
                        dr("NISSUMAIL") = ""
                    Else
                        dr("NISSUMAIL") = drE("NISSUMAIL").ToString.Trim
                    End If

                    '與要保人之關係(NINSREL)
                    If drE("NINSREL").ToString.Trim.Equals(String.Empty) Then
                        Me.IsInputErr = True
                        sbErr &= "與要保人之關係不可為空白." & Environment.NewLine
                        dr("NINSREL") = ""
                    Else
                        dr("NINSREL") = drE("NINSREL").ToString.Trim

                        If dr("NINSREL") = "1" Then
                            iii2 = iii2 + 1
                        End If

                        If (iii2 > 1 And dr("NINSREL") = "1") Then
                            Me.IsInputErr = True
                            sbErr &= "[與要保人之關係 [本人] 不可重複]." & Environment.NewLine
                        End If
                        If drE("NINSREL").ToString.Trim.Length > 1 Then
                            Me.IsInputErr = True
                            sbErr &= "[與要保人之關係 請輸入代碼." & Environment.NewLine
                        End If

                    End If


                    If dr("NINSREL") = "1" Then

                        If txtIins_IINS2.Text.Trim <> dr("IISSU") Then
                            Me.IsInputErr = True
                            sbErr &= "與要保人關係選擇[本人]時，要保人與被保人[身分證字號]要相同. " & Environment.NewLine

                        End If
                        If NAME_IINS.Text.Trim <> dr("NISSU") Then
                            Me.IsInputErr = True
                            sbErr &= "與要保人關係選擇[本人]時，要保人與被保人[姓名]要相同. " & Environment.NewLine
                        End If
                        'Birtdy_IINS
                        If Birtdy_IINS.DateText <> dr("NISSUBD") Then
                            Me.IsInputErr = True
                            sbErr &= "與要保人關係選擇[本人]時，要保人與被保人[生日]要相同. " & Environment.NewLine
                        End If
                    End If

                    '受益人姓名(NAMEREL)
                    If drE("NAMEREL").ToString.Trim.Equals(String.Empty) Then
                        dr("NAMEREL") = "法定繼承人" + vbCrLf
                    Else
                        dr("NAMEREL") = drE("NAMEREL").ToString.Trim
                    End If
                    'TELREL
                    If drE("TELREL").ToString.Trim.Equals(String.Empty) Then
                        dr("TELREL") = ""
                    Else
                        dr("TELREL") = drE("TELREL").ToString.Trim
                    End If
                    '受益人地址ZIP(ZIPREL)
                    If drE("ZIPREL").ToString.Trim.Equals(String.Empty) Then
                        dr("ZIPREL") = ""
                    Else
                        dr("ZIPREL") = drE("ZIPREL").ToString.Trim
                    End If
                    '受益人地址(AREL)
                    If drE("AREL").ToString.Trim.Equals(String.Empty) Then
                        dr("AREL") = ""
                    Else
                        dr("AREL") = drE("AREL").ToString.Trim
                    End If

                    '受益人身分証號(IDREL)
                    If drE("IDREL").ToString.Trim.Equals(String.Empty) Then
                        dr("IDREL") = ""
                    Else
                        If dr("NAMEREL") = "法定繼承人" Then
                            dr("IDREL") = ""
                        Else
                            If B2BEntity.HasFunction._CheckCitizenID(drE("IDREL").ToString.Trim) Then
                                dr("IDREL") = drE("IDREL").ToString.Trim
                            ElseIf B2BEntity.HasFunction._CheckPassportID(drE("IDREL").ToString.Trim) Then
                                dr("IDREL") = drE("IDREL").ToString.Trim
                            Else
                                Me.IsInputErr = True
                                sbErr &= "受益人身分證號 錯誤." & Environment.NewLine
                            End If
                        End If

                    End If

                    '與被保險人之關係(TYPEREL)

                    If drE("TYPEREL").ToString.Trim.Equals(String.Empty) Then
                        'Me.IsInputErr = True
                        'strErr &= "與被保險人之關係不可為空白."
                        dr("TYPEREL") = ""
                    Else
                        '無身故保險金
                        If dr("NAMEREL") = "法定繼承人" Then
                            dr("TYPEREL") = ""
                        Else
                            dr("TYPEREL") = drE("TYPEREL").ToString.Trim
                        End If

                    End If
                    '--------------------受益人2--------------------
                    '受益人姓名2(NAMEREL2)
                    If drE("NAMEREL2").ToString.Trim.Equals(String.Empty) Then
                        dr("NAMEREL2") = ""
                    Else
                        dr("NAMEREL2") = drE("NAMEREL2").ToString.Trim
                    End If

                    '受益人2身分証號(IDREL2)
                    If drE("IDREL2").ToString.Trim.Equals(String.Empty) Then
                        dr("IDREL2") = ""
                    Else
                        If B2BEntity.HasFunction._CheckCitizenID(drE("IDREL2").ToString.Trim) Then
                            dr("IDREL2") = drE("IDREL2").ToString.Trim
                        Else
                            Me.IsInputErr = True
                            sbErr &= "受益人2身分證號 錯誤." & Environment.NewLine

                        End If

                    End If
                    '與被保險人之關係2(TYPEREL2)
                    If drE("TYPEREL2").ToString.Trim.Equals(String.Empty) Then

                        dr("TYPEREL2") = ""
                    Else
                        dr("TYPEREL2") = drE("TYPEREL2").ToString.Trim
                    End If

                    'RELMOTHED 給付方式
                    If drE("RELMOTHED").ToString.Trim.Equals(String.Empty) Then
                        dr("RELMOTHED") = ""
                    Else
                        dr("RELMOTHED") = drE("RELMOTHED").ToString.Trim
                    End If
                    '--------------------受益人2--------------------
                    '--------------------受益人3--------------------
                    '受益人姓名3(NAMEREL3)
                    If drE("NAMEREL3").ToString.Trim.Equals(String.Empty) Then
                        dr("NAMEREL3") = ""
                    Else
                        dr("NAMEREL3") = drE("NAMEREL3").ToString.Trim
                    End If

                    '受益人3身分証號(IDREL3)
                    If drE("IDREL3").ToString.Trim.Equals(String.Empty) Then
                        dr("IDREL3") = ""
                    Else
                        If B2BEntity.HasFunction._CheckCitizenID(drE("IDREL3").ToString.Trim) Then
                            dr("IDREL3") = drE("IDREL3").ToString.Trim
                        Else
                            Me.IsInputErr = True
                            sbErr &= "受益人3身分證號 錯誤." & Environment.NewLine

                        End If

                    End If
                    '與被保險人之關係3(TYPEREL3)
                    If drE("TYPEREL3").ToString.Trim.Equals(String.Empty) Then

                        dr("TYPEREL3") = ""
                    Else
                        dr("TYPEREL3") = drE("TYPEREL3").ToString.Trim
                    End If
                    '--------------------受益人3--------------------

                    age = B2BEntity.HasFunction.GetAgeInLaw(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate())
                    '意外死亡及失能保額(萬)
                    If drE("ISSU_1").ToString.Trim.Equals(String.Empty) Then
                        If age < 15 Then
                            If dr("IISSU_TP") = "5" Then
                                Me.IsInputErr = True
                                sbErr &= "國內地區喪葬費用保額不可為0." & Environment.NewLine
                            Else
                                '15歲以下可以不保意外死亡及失能
                                dr("ISSU_1") = "0"
                            End If
                        Else
                            Me.IsInputErr = True
                            sbErr &= "意外死亡及失能保額不可為空白." & Environment.NewLine
                            dr("ISSU_1") = ""
                        End If
                    ElseIf (drE("iPlanNo").ToString.Trim = "D") Or (drE("iPlanNo").ToString.Trim = "N") Then
                        If GetDec(drE("ISSU_1").ToString.Trim) < 5000000 Then
                            If Checkbox2.Checked = True And age >= 15 Then
                                Me.IsInputErr = True
                                If (drE("iPlanNo").ToString.Trim = "D") Then
                                    sbErr &= "選擇申根證明文件，海外超值型建議搭配500萬(含)以上保額." & Environment.NewLine
                                ElseIf (drE("iPlanNo").ToString.Trim = "N") Then
                                    sbErr &= "選擇申根證明文件，海外特選型建議搭配500萬(含)以上保額." & Environment.NewLine
                                End If
                                dr("ISSU_1") = drE("ISSU_1").ToString.Trim
                            Else
                                dr("ISSU_1") = drE("ISSU_1").ToString.Trim
                            End If
                        Else
                            dr("ISSU_1") = drE("ISSU_1").ToString.Trim
                        End If
                    Else
                        If GetDec(drE("ISSU_1").ToString.Trim) < 100000 AndAlso drE("ISSU_1").ToString.Trim <> 0 Then
                            Me.IsInputErr = True
                            sbErr &= "意外死亡及失能保額不得低於10萬." & Environment.NewLine
                        End If
                        dr("ISSU_1") = drE("ISSU_1").ToString.Trim
                    End If
                    '要判斷金額 是否符合

                    '傷害醫療保額(萬)
                    If drE("ISSU_2").ToString.Trim.Equals(String.Empty) Then
                        dr("ISSU_2") = "0"
                    Else
                        dr("ISSU_2") = drE("ISSU_2").ToString.Trim
                    End If

                    If TRAVEL_A.SelectedValue = "400" Then
                        '國內團 沒有海突發
                        '海外突發疾病醫療保額(萬)(不含傳染病)
                        If drE("ISSU_3").ToString.Trim.Equals(String.Empty) Or drE("ISSU_3").ToString.Trim.Equals("0") Then
                            dr("ISSU_3") = "0"
                        Else
                            Me.IsInputErr = True
                            sbErr &= "出發地點為國內 不得承保海外突發疾病醫療保額." & Environment.NewLine
                            dr("ISSU_3") = ""
                        End If

                        '海外突發疾病保險金(含傳染病)						
                        If drE("ISSU_3D").ToString.Trim.Equals(String.Empty) Or drE("ISSU_3D").ToString.Trim.Equals("0") Then
                            dr("ISSU_3D") = "0"
                        Else
                            Me.IsInputErr = True
                            sbErr &= "出發地點為國內 不得承保海外突發疾病醫療保額." & Environment.NewLine
                            dr("ISSU_3D") = ""
                        End If

                        '海外突發疾病住院醫療保險金					
                        If drE("ISSU_IT64").ToString.Trim.Equals(String.Empty) Or drE("ISSU_IT64").ToString.Trim.Equals("0") Then
                            dr("ISSU_IT64") = "0"
                        Else
                            Me.IsInputErr = True
                            sbErr &= "出發地點為國內 不得承保海外突發疾病醫療保額." & Environment.NewLine
                            dr("ISSU_IT64") = ""
                        End If
                    Else
                        ''海外突發疾病醫療保額(萬)(不含傳染病)
                        'If drE("ISSU_3").ToString.Trim.Equals(String.Empty) Then
                        '	'Me.IsInputErr = True
                        '	'strErr &= "海外突發疾病醫療保額不可為空白."
                        '	dr("ISSU_3") = "0"
                        'Else
                        '	dr("ISSU_3") = drE("ISSU_3").ToString.Trim
                        'End If

                        'If rad_Disease_N.Checked = True Then
                        '海外突發疾病醫療保額(萬)(不含傳染病)
                        If drE("ISSU_3").ToString.Trim.Equals(String.Empty) Then
                            'Me.IsInputErr = True
                            'strErr &= "海外突發疾病醫療保額不可為空白."
                            dr("ISSU_3") = "0"
                        Else
                            dr("ISSU_3") = drE("ISSU_3").ToString.Trim
                        End If
                        'End If

                        'If rad_Disease_Y.Checked = True Then
                        '海外突發疾病保險金(含傳染病)
                        If drE("ISSU_3D").ToString.Trim.Equals(String.Empty) Then
                            'Me.IsInputErr = True
                            'strErr &= "海外突發疾病醫療保額不可為空白."
                            dr("ISSU_3D") = "0"
                        Else
                            dr("ISSU_3D") = drE("ISSU_3D").ToString.Trim
                        End If
                        'End If

                        '海外突發疾病住院醫療保險金
                        If drE("ISSU_IT64").ToString.Trim.Equals(String.Empty) Then
                            dr("ISSU_IT64") = "0"
                        Else
                            dr("ISSU_IT64") = drE("ISSU_IT64").ToString.Trim
                        End If

                        '三種海外突發疾病只能有一個
                        Dim tempcount As Int32 = 0

                        If dr("ISSU_3").ToString <> "0" Then tempcount += 1
                        If dr("ISSU_3D").ToString <> "0" Then tempcount += 1
                        If dr("ISSU_IT64").ToString <> "0" Then tempcount += 1

                        If tempcount > 1 Then
                            Me.IsInputErr = True
                            sbErr &= "海外突發疾病項目不可重複投保."
                        End If

                        If drE("iPlanNo").ToString = "J" Or drE("iPlanNo").ToString = "K" Or drE("iPlanNo").ToString = "L" Then
                            If dr("ISSU_3").ToString <> "0" Or dr("ISSU_3D").ToString <> "0" Then
                                Me.IsInputErr = True
                                sbErr &= "您選擇的方案，僅可投保 海外突發疾病住院醫療保險金(B)."
                            End If
                        ElseIf drE("iPlanNo").ToString = "O" Then
                            If dr("ISSU_3D").ToString <> "0" Then
                                Me.IsInputErr = True
                                sbErr &= "您選擇的方案，建議不投保 海外突發疾病醫療保險(含傳染病)(A)."
                            End If

                        ElseIf drE("iPlanNo").ToString = "M" Or drE("iPlanNo").ToString = "N" Then
                            ' M、N可保不含傳染病
                            If dr("ISSU_3D").ToString <> "0" Then
                                Me.IsInputErr = True
                                sbErr &= "您選擇的方案，僅可投保 海外突發疾病住院醫療保險金(B)."
                            End If
                        Else
                            If dr("ISSU_IT64").ToString <> "0" Then
                                Me.IsInputErr = True
                                sbErr &= "您選擇的方案，建議不投保 海外突發疾病住院醫療保險金(B)."
                            End If
                        End If

                    End If

                    '重大燒燙傷保險金(萬)
                    If drE("ISSU_IT44").ToString.Trim.Equals(String.Empty) Or drE("ISSU_IT44").ToString.Trim.Equals("0") Then
                        'Me.IsInputErr = True
                        'strErr &= "重大燒燙傷保險金."
                        dr("ISSU_IT44") = "0"
                    Else
                        '移除鎖控重大燒燙傷最低要保25萬(20170503)
                        dr("ISSU_IT44") = drE("ISSU_IT44").ToString.Trim
                        If CInt(dr("ISSU_IT44")) < 1 Then
                            Me.IsInputErr = True
                            sbErr &= "重大燒燙傷保險金 最低保額需為1萬." & Environment.NewLine
                        End If
                    End If

                    '懷孕否
                    If drE("EMP").ToString.Trim.Equals(String.Empty) Then
                        dr("EMP") = "N"
                    Else
                        dr("EMP") = drE("EMP").ToString.Trim
                        If dr("EMP") = "Y" Then
                            If (dr("ISSU_3").ToString.Trim <> "0" And String.IsNullOrEmpty(dr("ISSU_3").ToString.Trim) = False) Or (dr("ISSU_3D").ToString.Trim <> "0" And String.IsNullOrEmpty(dr("ISSU_3D").ToString.Trim) = False) Or (dr("ISSU_IT64").ToString.Trim <> "0" And String.IsNullOrEmpty(dr("ISSU_IT64").ToString.Trim) = False) Then
                                Me.IsInputErr = True
                                sbErr &= "被保險人為孕婦，建議不投保海外突發疾病保險金." & Environment.NewLine
                            End If
                        End If
                        If dr("EMP") = "Y" And dr("IISSU_SEX") = "1" Then
                            Me.IsInputErr = True
                            sbErr &= "被保險人是否為孕婦? 選擇性別錯誤." & Environment.NewLine
                        End If
                    End If
                    If IsDBNull(drE("iPlanNo")) = False Then
                        '舊版IQP不要20260109

                        '大眾運輸意外保險金(萬)
                        If (checkSpecialChannel().Item1 = True Or checkChannelO().Item1 = True) AndAlso (drE("iPlanNo") = "J" Or drE("iPlanNo") = "K" Or drE("iPlanNo") = "L") Then
                            If drE("ISSU_IT43").ToString.Trim.Equals(String.Empty) Then
                                If age >= 15 Then
                                    Me.IsInputErr = True
                                    sbErr &= "大眾運輸意外保險金 保額應為身故的20%." & Environment.NewLine
                                End If
                                dr("ISSU_IT43") = "0"
                            Else
                                If age < 15 AndAlso Not drE("ISSU_IT43").ToString.Trim.Equals("0") Then
                                    Me.IsInputErr = True
                                    sbErr &= "未滿足15歲，建議不投保大眾運輸意外保險金." & Environment.NewLine
                                    dr("ISSU_IT43") = "0"
                                Else
                                    dr("ISSU_IT43") = drE("ISSU_IT43").ToString.Trim
                                End If
                            End If
                        Else
                            If Not drE("ISSU_IT43").ToString.Trim.Equals(String.Empty) And Not drE("ISSU_IT43").ToString.Trim.Equals("0") Then
                                Me.IsInputErr = True
                                sbErr &= "您選擇的方案，建議不投保大眾運輸意外保險金." & Environment.NewLine
                            End If
                            dr("ISSU_IT43") = "0"
                        End If
                    End If
                    If TargetProject = "ITP" Then
                        '方案
                        If drE("iPlanNo").ToString.Trim.Equals(String.Empty) Then
                            Me.IsInputErr = True
                            sbErr &= "旅遊不便險方案不可為空白." & Environment.NewLine

                            dr("iPlanNo") = ""
                        Else
                            dr("iPlanNo") = drE("iPlanNo").ToString.Trim
                            '主被保人的方案為主
                            If i = 1 Then
                                txtCase.Text = GetiPlanNo(dr("iPlanNo").Trim)
                            End If
                            If GetiPlanNo(dr("iPlanNo").Trim) <> txtCase.Text.Trim Then
                                Me.IsInputErr = True
                                sbErr &= "旅遊方案需與主被保人相同." & Environment.NewLine
                            End If

                            If TRAVEL_A.SelectedValue = "400" Then
                                '國內
                                If dr("iPlanNo") = "A" Or dr("iPlanNo") = "B" Then
                                Else
                                    Me.IsInputErr = True
                                    If checkSpecialChannel().Item1 = True Then
                                        sbErr &= "旅遊地區為國內 無法選擇(C)(D)(J)(K)國外的方案." & Environment.NewLine
                                    ElseIf checkChannelO().Item1 = True Then
                                        sbErr &= "旅遊地區為國內 無法選擇(C)(D)(J)(K)(O)國外的方案." & Environment.NewLine
                                    ElseIf checkChannelMN().Item1 Then
                                        sbErr &= "旅遊地區為國內 無法選擇(M)(N)國外的方案." & Environment.NewLine
                                    Else
                                        sbErr &= "旅遊地區為國內 無法選擇(C)(D)國外的方案." & Environment.NewLine
                                    End If

                                End If
                            ElseIf TRAVEL_A.SelectedValue = "740" Then
                                '740
                                If checkSpecialChannel().Item1 = True Then
                                    If dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Then
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= "歐洲地區 請選擇 (C)(D)(J)(K)(L)的國外方案." & Environment.NewLine
                                    End If
                                ElseIf checkChannelO().Item1 Then
                                    If dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Or dr("iPlanNo") = "O" Then
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= "歐洲地區 請選擇 (C)(D)(J)(K)(L)(O)的國外方案." & Environment.NewLine
                                    End If
                                ElseIf checkChannelMN().Item1 Then
                                    If dr("iPlanNo") = "M" Or dr("iPlanNo") = "N" Then
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= "歐洲地區 請選擇 (M)(N)的國外方案." & Environment.NewLine
                                    End If
                                Else
                                    If dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Then
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= "歐洲地區 請選擇 (C)(D)的國外方案." & Environment.NewLine
                                    End If
                                End If
                            ElseIf TRAVEL_A.SelectedValue = "750" Then
                                If checkSpecialChannel().Item1 = True Then
                                    If Checkbox2.Checked = True Then
                                        '申根
                                        If dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "申根地區 請選擇 (D)(J)(K)的國外方案." & Environment.NewLine
                                        End If
                                    Else
                                        If dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "歐洲地區 請選擇 (C)(D)(J)(K)的國外方案." & Environment.NewLine
                                        End If
                                    End If
                                ElseIf checkChannelO().Item1 Then
                                    If Checkbox2.Checked = True Then
                                        '申根
                                        If dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Or dr("iPlanNo") = "O" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "申根地區 請選擇 (D)(J)(K)(O)的國外方案." & Environment.NewLine
                                        End If
                                    Else
                                        If dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "O" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "歐洲地區 請選擇 (C)(D)(J)(K)(O)的國外方案." & Environment.NewLine
                                        End If
                                    End If
                                ElseIf checkChannelMN().Item1 Then
                                    If Checkbox2.Checked = True Then
                                        '申根
                                        If dr("iPlanNo") = "N" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "申根地區 請選擇 (N)海外特選型方案." & Environment.NewLine
                                        End If
                                    Else
                                        If dr("iPlanNo") = "M" Or dr("iPlanNo") = "N" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "歐洲地區 請選擇 (M)(N)的國外方案." & Environment.NewLine
                                        End If
                                    End If
                                Else
                                    If Checkbox2.Checked = True Then
                                        '申根
                                        If dr("iPlanNo") = "D" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "申根地區 請選擇 (D)海外超值型方案." & Environment.NewLine
                                        End If
                                    Else
                                        If dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Then
                                        Else
                                            Me.IsInputErr = True
                                            sbErr &= "歐洲地區 請選擇 (C)(D)的國外方案." & Environment.NewLine
                                        End If
                                    End If
                                End If
                            Else
                                '國外
                                If dr("iPlanNo") = "A" Or dr("iPlanNo") = "B" Then
                                    Me.IsInputErr = True
                                    sbErr &= "旅遊地區為國外 無法選擇(A)(B)國內的方案." & Environment.NewLine
                                ElseIf checkSpecialChannel().Item1 = False AndAlso checkChannelO().Item1 = False AndAlso checkChannelMN().Item1 = False AndAlso (dr("iPlanNo") = "C" Or dr("iPlanNo") = "D") Then
                                ElseIf checkSpecialChannel().Item1 = False AndAlso checkChannelO().Item1 = False AndAlso checkChannelMN().Item1 = True AndAlso (dr("iPlanNo") = "M" Or dr("iPlanNo") = "N") Then
                                ElseIf checkSpecialChannel().Item1 = True AndAlso (dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L") Then
                                ElseIf checkChannelO().Item1 = True AndAlso (dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Or dr("iPlanNo") = "O") Then
                                Else
                                    Me.IsInputErr = True
                                    sbErr &= "旅遊方案 代碼錯誤." & Environment.NewLine
                                End If
                            End If

                        End If
                    End If

                    Try
                        age = B2BEntity.HasFunction.GetAgeInLaw(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate())

                        '國別 身份別為5護照號碼時，上方已判別過國別
                        If dr("IISSU_TP") <> "5" Then
                            If drE("IISSU_COUNTRY").ToString.Trim.Equals(String.Empty) Then
                                dr("IISSU_COUNTRY") = "TW"
                            Else
                                dr("IISSU_COUNTRY") = drE("IISSU_COUNTRY").ToString.Trim
                            End If
                        End If

                        If Checkbox2.Checked = True And TRAVEL_A.SelectedValue = "750" Then
                            '要排除81歲
                            If age < 81 Then
                                '英文姓名
                                If drE("PASSPORT_NAME").ToString.Trim.Equals(String.Empty) Then
                                    Me.IsInputErr = True
                                    sbErr &= "申根區-英文姓名不可為空白." & Environment.NewLine
                                    dr("PASSPORT_NAME") = ""
                                Else
                                    dr("PASSPORT_NAME") = drE("PASSPORT_NAME").ToString.Trim

                                End If
                                '護照號碼
                                If drE("PASSPORT_NUMBER").ToString.Trim.Equals(String.Empty) Then
                                    Me.IsInputErr = True
                                    sbErr &= "申根區-護照號碼不可為空白." & Environment.NewLine
                                    dr("PASSPORT_NUMBER") = ""
                                Else
                                    dr("PASSPORT_NUMBER") = drE("PASSPORT_NUMBER").ToString.Trim

                                End If
                            End If
                        ElseIf TRAVEL_A.SelectedValue <> "400" Then
                            '要排除81歲
                            If age < 81 Then
                                '英文姓名
                                Dim passportCheck As String = drE("PASSPORT_NAME").ToString.Trim & drE("PASSPORT_NUMBER").ToString.Trim
                                If String.IsNullOrWhiteSpace(passportCheck) = False AndAlso (String.IsNullOrWhiteSpace(drE("PASSPORT_NAME").ToString.Trim) = True Or String.IsNullOrWhiteSpace(drE("PASSPORT_NUMBER").ToString.Trim) = True) Then
                                    Me.IsInputErr = True
                                    sbErr &= "英文姓名 與 護照號碼 不可擇一填寫" & Environment.NewLine
                                Else
                                    dr("PASSPORT_NAME") = drE("PASSPORT_NAME").ToString.Trim
                                    dr("PASSPORT_NUMBER") = drE("PASSPORT_NUMBER").ToString.Trim
                                End If
                            End If
                        End If

                    Catch ex As Exception
                        Throw New Exception("保費試算錯誤，請調整後再行檢核")
                    End Try
#End Region'"============================================================================================="

                    '檢查年齡到底是幾歲
                    '判斷被保人年齡
                    age = B2BEntity.HasFunction.GetAgeInLaw(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate())
                    If age < 15 Then
                        InsType = "IT4Y"
                    Else
                        InsType = "IT41"
                    End If
                    InsDay = txtIQPITP_Days.Text.Trim
                    If dr("ISSU_1") > "0" Then
                        dr("ISSU_1") = dr("ISSU_1")
                    End If

                    Dim sbErr2 As String = ""
                    If sbErr <> "" Then
                        sbErr2 = sbErr
                        sbErr = dr("IISSU") & " " & sbErr2 & Environment.NewLine
                    End If

                    TOTsbErr += sbErr

                    InsCoverage = dr("ISSU_1")
                    InsMdeduct = "0"

                    '5.CountryType  代表費率別 ：3=個人費率、4=團費率
                    If CHK5.Checked = True Then
                        CountryType = "4"
                    Else
                        CountryType = "3"
                    End If

                    'ITP 不便險取保費的方式 
                    'IT41/IT4Y/IT42/IT62/IT63 這四檔 旅平的相關險種
                    '1:InsType(險種)
                    '2:InsDay(天數)
                    '3.InsCoverage 保額 元
                    '4.InsMdeduct 自負額 元
                    '5.CountryType  代表費率別 ：3=個人費率、4=團費率
                    'Dflag(要抓該費率檔內小於系統日的最大日)
                    Dim sumsum1 As Integer = 0
                    Dim sumsum2 As Integer = 0
                    Dim sumsum3 As Integer = 0
                    Dim sumsum4 As Integer = 0
                    Dim sumsum5 As Integer = 0
                    Dim sumsum6 As Integer = 0 'IT43
                    Dim sumsum7 As Integer = 0 'IT64
                    Dim sumsum8 As Integer = 0 'IT65 
                    Dim sumsum9 As Integer = 0 'IT66
                    Dim sumsum10 As Integer = 0 'IT6D
                    Dim sumsum11 As Integer = 0 'IT6F
                    Dim sumsum12 As Integer = 0 'IT6G

                    '算金額1

                    '只限國外地區==未滿15足歲投保旅平險，可以只投保「傷害醫療保險金」及「海外突發疾病保險金(甲型)，故刪除系統檢核「請選擇身故失能保險金 或 失能保險金」。
                    If TRAVEL_A.SelectedValue = "400" Then
                        '國內
                        If dr("ISSU_1").ToString.Trim = "0" Then
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "未滿15歲投保 最高建議保額為69萬-- 匯入金額錯誤不符." & Environment.NewLine
                        End If
                    Else
                        '國外
                        If age < 15 Then
                            If dr("ISSU_1").ToString.Trim = "0" Then
                                If checkChannelMN().Item1 AndAlso TargetProject = "ITP" Then 'checkChannelMN = true，IQP的海突傳染病是保IT63
                                    If dr("ISSU_2").ToString.Trim = "0" Or (dr("ISSU_3").ToString.Trim = "0" And dr("ISSU_IT64").ToString.Trim = "0") Then
                                        Me.IsInputErr = True
                                        sbErr &= dr("IISSU") + "未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金." & Environment.NewLine
                                    End If
                                ElseIf (checkSpecialChannel().Item1 = False And checkChannelO().Item1 = False) AndAlso (dr("ISSU_2").ToString.Trim = "0" Or (dr("ISSU_3").ToString.Trim = "0" And dr("ISSU_3D").ToString.Trim = "0")) Then
                                    Me.IsInputErr = True
                                    sbErr &= dr("IISSU") + "未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金." & Environment.NewLine
                                ElseIf ((checkSpecialChannel().Item1 = True Or checkChannelO().Item1 = True) AndAlso (dr("iPlanNo").ToString = "J" Or dr("iPlanNo").ToString = "K" Or dr("iPlanNo").ToString = "L")) AndAlso
                                        (dr("ISSU_2").ToString.Trim = "0" Or dr("ISSU_IT64").ToString.Trim = "0") Then
                                    Me.IsInputErr = True
                                    sbErr &= dr("IISSU") + "未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金." & Environment.NewLine
                                ElseIf checkChannelO().Item1 AndAlso dr("iPlanNo").ToString = "O" AndAlso (dr("ISSU_2").ToString.Trim = "0" Or (dr("ISSU_3").ToString.Trim = "0" And dr("ISSU_IT64").ToString.Trim = "0")) Then
                                    Me.IsInputErr = True
                                    sbErr &= dr("IISSU") + "未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金." & Environment.NewLine
                                ElseIf ((checkSpecialChannel().Item1 = True Or checkChannelO().Item1 = True) AndAlso (dr("iPlanNo").ToString <> "J" And dr("iPlanNo").ToString <> "K" And dr("iPlanNo").ToString <> "L" And dr("iPlanNo").ToString <> "O")) AndAlso
                                        (dr("ISSU_2").ToString.Trim = "0" Or (dr("ISSU_3").ToString.Trim = "0" And dr("ISSU_3D").ToString.Trim = "0")) Then
                                    Me.IsInputErr = True
                                    sbErr &= dr("IISSU") + "未滿15歲 請選擇身傷害醫療保險金 及 海外突發疾病保險金." & Environment.NewLine
                                End If
                            End If
                        End If
                    End If
                    ''重大燒燙傷保險金檢核
                    'If CheckFireBurn(dr("ISSU_1"), dr("ISSU_IT44")) <> String.Empty Then
                    '    Me.IsInputErr = True
                    '    sbErr &= "重大燒燙傷保險金， 上限不得超過死亡及失能保額4分之1" & Environment.NewLine
                    'End If
                    If checkSpecialChannel().Item1 = False AndAlso dr("ISSU_1").ToString.Trim = "0" AndAlso dr("ISSU_IT44").ToString.Trim <> "0" Then
                        Me.IsInputErr = True
                        sbErr &= "未滿15歲投保無喪葬費用方案 ，建議不投保重大燒燙傷保險金" & Environment.NewLine
                    End If

                    Dim dsMoney1 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                    If dsMoney1.Tables("ITPPrem").Rows.Count > 0 Then
                        sumsum1 = CInt(dsMoney1.Tables("ITPPrem").Rows(0)("PREM"))
                    Else
                        '不可能沒資料 沒資料就是有問題

                        If age < 15 Then
                            If dr("ISSU_1") = "0" Then
                                sumsum1 = 0
                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("IISSU") + "未滿15歲投保 最高建議保額為69萬-- 匯入金額錯誤不符." & Environment.NewLine
                            End If
                        Else
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "意外死亡及失能保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                        End If
                        sumsum1 = 0
                    End If

                    InsType = "IT42"
                    If dr("ISSU_2") > "0" Then
                        dr("ISSU_2") = dr("ISSU_2")
                    End If
                    InsCoverage = dr("ISSU_2")

                    '判斷金額是否正確
                    If InsCoverage = "0" Or InsCoverage = "" Then
                        sumsum2 = 0
                    Else
                        '算金額2
                        Dim dsMoney2 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                        If dsMoney2.Tables("ITPPrem").Rows.Count > 0 Then
                            sumsum2 = CInt(dsMoney2.Tables("ITPPrem").Rows(0)("PREM"))
                        Else
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "傷害醫療保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                            sumsum2 = 0
                        End If
                    End If

                    If TRAVEL_A.SelectedValue <> "400" Then
                        InsType = "IT62"
                        If dr("ISSU_3") > "0" Then
                            dr("ISSU_3") = dr("ISSU_3")
                        End If
                        InsCoverage = dr("ISSU_3")

                        '判斷金額是否正確
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum3 = 0
                        Else
                            '算金額3
                            Dim dsMoney3 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                            If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                                sumsum3 = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))
                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                                sumsum3 = 0
                            End If
                        End If
                    Else
                        If dr("ISSU_3") > "0" Then
                            dr("ISSU_3") = dr("ISSU_3")
                        End If
                        InsCoverage = dr("ISSU_3")
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum3 = 0
                        Else
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "國內地區無法投保 海外突發疾病醫療保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                            sumsum3 = 0
                        End If
                    End If

                    If TRAVEL_A.SelectedValue <> "400" Then
                        If TargetProject = "ITP" Then
                            InsType = "IT6E"
                        Else
                            InsType = "IT63"
                        End If
                        If dr("ISSU_3D") > "0" Then
                            dr("ISSU_3D") = dr("ISSU_3D")
                        End If
                        InsCoverage = dr("ISSU_3D")

                        '判斷金額是否正確
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum5 = 0
                        Else
                            '算金額3
                            Dim dsMoney3 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                            If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                                sumsum5 = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))

                                'F115060600 新增IT6D，方案基本型 & 超值型才有的附加保險金
                                If TargetProject = "ITP" AndAlso (dr("iPlanNo") = "C" Or dr("iPlanNo") = "D") Then

                                    '算金額IT6F
                                    InsCoverage = (CInt(dr("ISSU_3D")) * 0.01).ToString
                                    InsType = "IT6F"
                                    dsMoney3 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                    If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                                        sumsum11 = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= dr("IISSU") + "海外突發疾病急診醫療保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                                        sumsum5 = 0
                                        sumsum11 = 0
                                    End If

                                    '算金額IT6G
                                    InsType = "IT6G"
                                    dsMoney3 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                    If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                                        sumsum12 = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= dr("IISSU") + "海外突發疾病門診醫療保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                                        sumsum5 = 0
                                        sumsum11 = 0
                                        sumsum12 = 0
                                    End If

                                    '算金額IT6D
                                    InsCoverage = (CInt(dr("ISSU_3D")) * 0.05).ToString
                                    InsType = "IT6D"
                                    dsMoney3 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                    If dsMoney3.Tables("ITPPrem").Rows.Count > 0 Then
                                        sumsum10 = CInt(dsMoney3.Tables("ITPPrem").Rows(0)("PREM"))
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                        sumsum5 = 0
                                        sumsum11 = 0
                                        sumsum12 = 0
                                        sumsum10 = 0
                                    End If
                                End If
                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                sumsum5 = 0
                            End If
                        End If
                    Else
                        If dr("ISSU_3D") > "0" Then
                            dr("ISSU_3D") = dr("ISSU_3D")
                        End If
                        InsCoverage = dr("ISSU_3D")
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum5 = 0
                        Else
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "國內地區無法投保 海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                            sumsum5 = 0
                        End If
                    End If



                    InsType = "IT44"
                    If dr("ISSU_IT44") > "0" Then
                        dr("ISSU_IT44") = dr("ISSU_IT44")
                    End If
                    InsCoverage = dr("ISSU_IT44")

                    If InsCoverage = "0" Or InsCoverage = "" Then
                        sumsum4 = 0
                    Else
                        '算金額4
                        Dim dsMoney4 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                        If dsMoney4.Tables("ITPPrem").Rows.Count > 0 Then
                            sumsum4 = CInt(dsMoney4.Tables("ITPPrem").Rows(0)("PREM"))
                        Else
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "重大燒燙傷保險金(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                            sumsum4 = 0
                        End If
                    End If

                    If IsDBNull(dr("ISSU_IT43")) = False Then
                        '舊版IQP不要20260109
                        InsType = "IT43"
                        If dr("ISSU_IT43") > "0" Then
                            dr("ISSU_IT43") = dr("ISSU_IT43")
                        End If
                        InsCoverage = dr("ISSU_IT43")

                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum6 = 0
                        Else
                            '算金額4
                            Dim dsMoney5 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                            If dsMoney5.Tables("ITPPrem").Rows.Count > 0 Then
                                sumsum6 = CInt(dsMoney5.Tables("ITPPrem").Rows(0)("PREM"))
                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("IISSU") + "大眾運輸意外保險金(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                                sumsum6 = 0
                            End If
                        End If

                    End If

                    If TRAVEL_A.SelectedValue <> "400" Then
                        InsType = "IT64"
                        If dr("ISSU_IT64") > "0" Then
                            dr("ISSU_IT64") = dr("ISSU_IT64")
                        End If
                        InsCoverage = dr("ISSU_IT64")

                        '判斷金額是否正確
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum7 = 0
                        Else
                            '算金額IT64
                            Dim dsMoney64 As DataSet = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                            If dsMoney64.Tables("ITPPrem").Rows.Count > 0 Then
                                sumsum7 = CInt(dsMoney64.Tables("ITPPrem").Rows(0)("PREM"))

                                InsCoverage = (CInt(InsCoverage) * 0.05).ToString

                                '算金額IT65
                                InsType = "IT65"
                                dsMoney64 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                If dsMoney64.Tables("ITPPrem").Rows.Count > 0 Then
                                    sumsum8 = CInt(dsMoney64.Tables("ITPPrem").Rows(0)("PREM"))
                                Else
                                    Me.IsInputErr = True
                                    sbErr &= dr("IISSU") + "海外突發疾病門診醫療保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                                    sumsum7 = 0
                                    sumsum8 = 0
                                End If

                                '算金額IT66 
                                InsType = "IT66"
                                dsMoney64 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                If dsMoney64.Tables("ITPPrem").Rows.Count > 0 Then
                                    sumsum9 = CInt(dsMoney64.Tables("ITPPrem").Rows(0)("PREM"))
                                Else
                                    Me.IsInputErr = True
                                    sbErr &= dr("IISSU") + "海外突發疾病急診醫療保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                                    sumsum7 = 0
                                    sumsum8 = 0
                                    sumsum9 = 0
                                End If


                                If checkChannelMN().Item1 Or (checkChannelO().Item1 AndAlso dr("iPlanNo") = "O") Then
                                    '算金額IT6D
                                    InsType = "IT6D"
                                    dsMoney64 = Me.InvokeCommonService("B2BService.bsITP101000GetOracleValue", InsType, InsDay, InsCoverage, InsMdeduct, CountryType)

                                    If dsMoney64.Tables("ITPPrem").Rows.Count > 0 Then
                                        sumsum10 = CInt(dsMoney64.Tables("ITPPrem").Rows(0)("PREM"))
                                    Else
                                        Me.IsInputErr = True
                                        sbErr &= dr("IISSU") + "海外突發疾病醫療保額(萬)(含傳染病) -- 匯入金額錯誤不符." & Environment.NewLine
                                        sumsum7 = 0
                                        sumsum8 = 0
                                        sumsum9 = 0
                                        sumsum10 = 0
                                    End If

                                End If
                            Else
                                Me.IsInputErr = True
                                sbErr &= dr("IISSU") + "海外突發疾病住院醫療保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                                sumsum7 = 0
                            End If
                        End If
                    Else
                        If dr("ISSU_IT64") > "0" Then
                            dr("ISSU_IT64") = dr("ISSU_IT64")
                        End If
                        InsCoverage = dr("ISSU_IT64")
                        If InsCoverage = "0" Or InsCoverage = "" Then
                            sumsum7 = 0
                        Else
                            Me.IsInputErr = True
                            sbErr &= dr("IISSU") + "國內地區無法投保 海外突發疾病住院醫療保額(萬) -- 匯入金額錯誤不符." & Environment.NewLine
                            sumsum7 = 0
                        End If
                    End If

                    dr("ISSU_4") = sumsum1 + sumsum2 + sumsum3 + sumsum4 + sumsum5 + sumsum6 + sumsum7 + sumsum8 + sumsum9 + sumsum11 + sumsum12

                    'ITP 年齡檢核
                    Dim sSolution, IADD1, IADD2, IADD3, IADD4, IADD64, IADD43 As String

                    sSolution = dr("ISSU_1").ToString
                    IADD1 = dr("ISSU_2").ToString
                    IADD2 = dr("ISSU_3").ToString
                    IADD3 = dr("ISSU_IT44").ToString
                    IADD4 = dr("ISSU_3D").ToString
                    IADD64 = dr("ISSU_IT64").ToString
                    IADD43 = dr("ISSU_IT43").ToString

                    If IsDBNull(dr("iPlanNo")) Then
                        '先給他空白
                        dr("iPlanNo") = ""
                    End If

                    Dim ChkStr As String
                    If IsDate(dr("NISSUBD")) = True Then
                        If checkSpecialChannel().Item1 = True AndAlso (dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L") Then
                            ChkStr = CheckAge_ITP_SPEChannel(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate(), sSolution, IADD1, IADD2, dr("iPlanNo").ToString, dr("IISSU_TP"), dr("iPlanNo"))
                        ElseIf checkChannelO().Item1 = True AndAlso (dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Or dr("iPlanNo") = "O") Then
                            ChkStr = CheckAge_ITP_SPEChannel(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate(), sSolution, IADD1, IADD2, dr("iPlanNo").ToString, dr("IISSU_TP"), dr("iPlanNo"))
                        Else
                            ChkStr = CheckAge_ITP(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate(), sSolution, IADD1, IADD2, dr("iPlanNo").ToString, dr("IISSU_TP"))
                        End If

                        If ChkStr <> String.Empty Then
                            Me.IsInputErr = True

                            sbErr &= ChkStr
                        End If
                    End If

                    Dim ChkStr2 As String = String.Empty
                    If IsDate(dr("NISSUBD")) = True Then
                        If checkSpecialChannel().Item1 = True AndAlso (dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L") Then
                            ChkStr2 = CheckiPlanNo_SPEChannel(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate(), dr("iPlanNo").ToString, sSolution, IADD1, IADD2, IADD3, IADD64, dr("IISSU_TP"), IADD43)
                        ElseIf checkChannelO().Item1 = True AndAlso (dr("iPlanNo") = "J" Or dr("iPlanNo") = "K" Or dr("iPlanNo") = "L" Or dr("iPlanNo") = "O") Then
                            ChkStr2 = CheckiPlanNo_SPEChannel(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate(), dr("iPlanNo").ToString, sSolution, IADD1, IADD2, IADD3, IADD64, dr("IISSU_TP"), IADD43)
                        Else
                            ChkStr2 = CheckiPlanNo(dr("NISSUBD"), Me.UcRocDobjbgn.GetDate(), dr("iPlanNo").ToString, sSolution, IADD1, IADD2, IADD3, If(checkChannelMN().Item1 AndAlso TargetProject = "ITP", IADD64, IADD4), dr("IISSU_TP"), IADD43)
                        End If

                        If ChkStr2 <> String.Empty Then
                            Me.IsInputErr = True
                            sbErr &= ChkStr2
                        End If
                    End If

                    If Not sbErr.Equals(String.Empty) Then dr("ERR") = sbErr
                    ds.IQPISSU.AddIQPISSURow(dr)
                    BaseSetChildFormData(Me.dgdMain, Me.StateMain)

                    If age < 15 Then
                        'IT4Y
                        Dim dr2 As beITP101000.IQPOBJRow
                        dr2 = ds.IQPOBJ.NewIQPOBJRow
                        dr2("IQPMainOID") = OID
                        dr2("IISSU") = dr("IISSU")
                        dr2("INSTyp") = If(TargetProject = "ITP", "IT4Y", "IQP0Y")
                        dr2("CONTENT") = "旅行平安保險(未滿15歲)"
                        dr2("INSAMT") = dr("ISSU_1")
                        dr2("INSPREM") = sumsum1
                        ds.IQPOBJ.AddIQPOBJRow(dr2)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    Else
                        'IT41
                        Dim dr1 As beITP101000.IQPOBJRow
                        dr1 = ds.IQPOBJ.NewIQPOBJRow
                        dr1("IQPMainOID") = OID
                        dr1("IISSU") = dr("IISSU")
                        dr1("INSTyp") = If(TargetProject = "ITP", "IT41", "IQP01")
                        dr1("CONTENT") = "旅行平安保險[標準型]"
                        dr1("INSAMT") = GetDec(dr("ISSU_1"))
                        dr1("INSPREM") = sumsum1
                        ds.IQPOBJ.AddIQPOBJRow(dr1)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)

                    End If
                    'IT42
                    If sumsum2 <> 0 Then
                        Dim dr3 As beITP101000.IQPOBJRow
                        dr3 = ds.IQPOBJ.NewIQPOBJRow
                        dr3("IQPMainOID") = OID
                        dr3("IISSU") = dr("IISSU")
                        dr3("INSTyp") = If(TargetProject = "ITP", "IT42", "IQP02")
                        dr3("CONTENT") = "傷害醫療保險給付[實支實付型]"
                        dr3("INSAMT") = dr("ISSU_2")
                        dr3("INSPREM") = sumsum2
                        ds.IQPOBJ.AddIQPOBJRow(dr3)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT44
                    If sumsum4 <> 0 Then
                        Dim dr44 As beITP101000.IQPOBJRow
                        dr44 = ds.IQPOBJ.NewIQPOBJRow
                        dr44("IQPMainOID") = OID
                        dr44("IISSU") = dr("IISSU")
                        dr44("INSTyp") = If(TargetProject = "ITP", "IT44", "IQP04")
                        dr44("CONTENT") = "重大燒燙傷保險金"
                        dr44("INSAMT") = dr("ISSU_IT44")
                        dr44("INSPREM") = sumsum4
                        ds.IQPOBJ.AddIQPOBJRow(dr44)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT62
                    If sumsum3 <> 0 Then
                        Dim dr4 As beITP101000.IQPOBJRow
                        dr4 = ds.IQPOBJ.NewIQPOBJRow
                        dr4("IQPMainOID") = OID
                        dr4("IISSU") = dr("IISSU")
                        dr4("INSTyp") = If(TargetProject = "ITP", "IT62", "IQP11")
                        dr4("CONTENT") = "海外突發疾病醫療保險金(不含傳染病)"
                        dr4("INSAMT") = dr("ISSU_3")
                        dr4("INSPREM") = sumsum3
                        ds.IQPOBJ.AddIQPOBJRow(dr4)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT63 or IT6E
                    If sumsum5 <> 0 Then
                        Dim dr5 As beITP101000.IQPOBJRow
                        dr5 = ds.IQPOBJ.NewIQPOBJRow
                        dr5("IQPMainOID") = OID
                        dr5("IISSU") = dr("IISSU")
                        dr5("INSTyp") = If(TargetProject = "ITP", "IT6E", "IQP12")
                        dr5("CONTENT") = "海外突發疾病醫療保險金(含傳染病)"
                        dr5("INSAMT") = dr("ISSU_3D")
                        dr5("INSPREM") = sumsum5
                        ds.IQPOBJ.AddIQPOBJRow(dr5)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT43
                    If sumsum6 <> 0 Then
                        Dim dr6 As beITP101000.IQPOBJRow
                        dr6 = ds.IQPOBJ.NewIQPOBJRow
                        dr6("IQPMainOID") = OID
                        dr6("IISSU") = dr("IISSU")
                        dr6("INSTyp") = If(TargetProject = "ITP", "IT43", "IQP03")
                        dr6("CONTENT") = "大眾運輸意外保險金"
                        dr6("INSAMT") = dr("ISSU_IT43")
                        dr6("INSPREM") = sumsum6
                        ds.IQPOBJ.AddIQPOBJRow(dr6)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT64
                    If sumsum7 <> 0 Then
                        Dim dr7 As beITP101000.IQPOBJRow
                        dr7 = ds.IQPOBJ.NewIQPOBJRow
                        dr7("IQPMainOID") = OID
                        dr7("IISSU") = dr("IISSU")
                        dr7("INSTyp") = If(TargetProject = "ITP", "IT64", "IQP13")
                        dr7("CONTENT") = "海外突發疾病住院醫療保險金"
                        dr7("INSAMT") = dr("ISSU_IT64")
                        dr7("INSPREM") = sumsum7
                        ds.IQPOBJ.AddIQPOBJRow(dr7)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT65
                    If sumsum8 <> 0 Then
                        Dim dr8 As beITP101000.IQPOBJRow
                        dr8 = ds.IQPOBJ.NewIQPOBJRow
                        dr8("IQPMainOID") = OID
                        dr8("IISSU") = dr("IISSU")
                        dr8("INSTyp") = If(TargetProject = "ITP", "IT65", "IQP14")
                        dr8("CONTENT") = "海外突發疾病門診醫療保險金"
                        dr8("INSAMT") = (CInt(dr("ISSU_IT64")) * 0.05).ToString
                        dr8("INSPREM") = sumsum8
                        ds.IQPOBJ.AddIQPOBJRow(dr8)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT66
                    If sumsum9 <> 0 Then
                        Dim dr9 As beITP101000.IQPOBJRow
                        dr9 = ds.IQPOBJ.NewIQPOBJRow
                        dr9("IQPMainOID") = OID
                        dr9("IISSU") = dr("IISSU")
                        dr9("INSTyp") = If(TargetProject = "ITP", "IT66", "IQP15")
                        dr9("CONTENT") = "海外突發疾病急診醫療保險金"
                        dr9("INSAMT") = (CInt(dr("ISSU_IT64")) * 0.05).ToString
                        dr9("INSPREM") = sumsum9
                        ds.IQPOBJ.AddIQPOBJRow(dr9)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT6F
                    If sumsum11 <> 0 AndAlso TargetProject = "ITP" AndAlso (dr("iPlanNo") = "C" Or dr("iPlanNo") = "D") Then
                        Dim dr11 As beITP101000.IQPOBJRow
                        dr11 = ds.IQPOBJ.NewIQPOBJRow
                        dr11("IQPMainOID") = OID
                        dr11("IISSU") = dr("IISSU")
                        dr11("INSTyp") = "IT6F"
                        dr11("CONTENT") = "海外突發疾病急診醫療保險金"
                        dr11("INSAMT") = (CInt(dr("ISSU_3D")) * 0.01).ToString
                        dr11("INSPREM") = sumsum11
                        ds.IQPOBJ.AddIQPOBJRow(dr11)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT6G
                    If sumsum12 <> 0 AndAlso TargetProject = "ITP" AndAlso (dr("iPlanNo") = "C" Or dr("iPlanNo") = "D") Then
                        Dim dr12 As beITP101000.IQPOBJRow
                        dr12 = ds.IQPOBJ.NewIQPOBJRow
                        dr12("IQPMainOID") = OID
                        dr12("IISSU") = dr("IISSU")
                        dr12("INSTyp") = "IT6G"
                        dr12("CONTENT") = "海外突發疾病門診醫療保險金"
                        dr12("INSAMT") = (CInt(dr("ISSU_3D")) * 0.01).ToString
                        dr12("INSPREM") = sumsum12
                        ds.IQPOBJ.AddIQPOBJRow(dr12)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    'IT6D
                    If sumsum10 <> 0 AndAlso TargetProject = "ITP" AndAlso (dr("iPlanNo") = "C" Or dr("iPlanNo") = "D" Or dr("iPlanNo") = "M" Or dr("iPlanNo") = "N" Or dr("iPlanNo") = "O") Then
                        Dim dr10 As beITP101000.IQPOBJRow
                        dr10 = ds.IQPOBJ.NewIQPOBJRow
                        dr10("IQPMainOID") = OID
                        dr10("IISSU") = dr("IISSU")
                        dr10("INSTyp") = "IT6D"
                        dr10("CONTENT") = "返國繼續住院醫療保險金"
                        If checkChannelMN().Item1 Or (checkChannelO().Item1 AndAlso dr("iPlanNo") = "O") Then
                            dr10("INSAMT") = (CInt(dr("ISSU_IT64")) * 0.05).ToString
                        Else
                            dr10("INSAMT") = (CInt(dr("ISSU_3D")) * 0.05).ToString
                        End If
                        dr10("INSPREM") = sumsum10
                        ds.IQPOBJ.AddIQPOBJRow(dr10)
                        BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
                    End If
                    If TargetProject = "ITP" AndAlso dr("iPlanNo") <> "L" Then
#Region "checkmain不便險方案"
                        '不便險方案
                        'GET_IQPOBJ(dr("IISSU"), dr("iPlanNo"), i)
                        '險種代號       C 海外基本型	D 海外超值型
                        'IT11		    限額200000元	限額200000元
                        'IT2R		    限額100萬		限額300萬
                        'IT37(原IT2K)   限額10000元		限額30000元
                        'IT36(原IT2J)   限額20000元		限額60000元 
                        'IT3A(原IT2A)   每次4000元		每次6000元
                        'IT3B(原IT2L)   每次3000元		每次3000元
                        'IT3C           每次10000元		每次30000元
                        'IT3D           每次1000元		每次2000元
                        'IT3E           每次1000元		每次2000元
                        'IT39(原IT29)   每次4000元		每次6000元
                        'IT38(原IT2X)	每次4000元		每次6000元
                        'IT2M		    每次3000元		每次6000元
                        'IT27		    每日3000元		每日5000元
                        'IT2D		    限額20000元		限額30000元
                        'IT28		    每次3000元		每次5000元
                        'IT31		    限額50000元		限額50000元

                        '險種代號       J 海外標準型	K 海外豪華型
                        'IT11		    限額3000000元	限額5000000元
                        'IT2R		    限額250萬		限額250萬
                        'IT3B(原IT2L)   每次3000元		每次3000元
                        'IT3A(原IT2A)   每次4000元		每次6000元
                        'IT39(原IT29)   每次4000元		每次6000元
                        'IT38(原IT2X)   每次4000元		每次6000元
                        'IT2M		    每次3000元		每次5000元
                        'IT27		    每日3000元		每日6000元
                        'IT2D		    限額30000元		限額50000元
                        'IT28		    每次3000元		每次3000元
                        'IT31		    限額30000元		限額50000元
                        'IT36(原IT32)   限額120000元	限額150000元
                        'IT37(原IT33)   限額120000元	限額150000元
                        'IT34		    限額120000元	限額150000元
                        'IT35		    限額120000元	限額150000元
                        'IT2P		    限額3000元		限額5000元
                        'IT2T		    每日10000元		每日10000元
                        'IT2U		    每次3000元		每次5000元
                        'IT2S		    每次3000元		每次5000元

                        '險種代號       O 海外小資型	
                        'IT11		    限額200000元		
                        'IT36(原IT2J)   限額60000元		
                        'IT37(原IT2K)   限額30000元
                        'IT38(原IT2X)	每次5000元
                        'IT39(原IT29)   每次5000元
                        'IT3A(原IT2A)   每次5000元		
                        'IT3B(原IT2L)   每次3000元	
                        'IT2M		    每次5000元		
                        'IT28		    每次3000元	
                        'IT31		    限額50000元	
                        'IT2D		    限額50000元	
                        'IT3D           每次1000元	

                        Dim iPlanNo, IISSU As String
                        iPlanNo = dr("iPlanNo").ToString.Trim
                        IISSU = dr("IISSU").ToString.Trim
                        InsType = ""
                        InsDay = ""
                        InsCoverage = ""
                        InsMdeduct = ""
                        CountryType = ""


                        '20251218
                        '	身故500萬以上，方案J K優惠IT2R不算錢
                        Dim IT2RFree = False
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            If CType(sSolution, Int32) >= 5000000 Then IT2RFree = True
                        End If

                        'ITP 不便險取保費的方式
                        '0 IISSU
                        '1:InsType(險種)
                        '2:InsDay(天數)
                        '3.InsCoverage 保額 元
                        '4.InsMdeduct 自負額 元
                        '5.CountryType 1=國內 2=國外
                        '=====IT11=====
                        InsType = "IT11"
                        InsDay = txtIQPITP_Days.Text
                        If iPlanNo = "A" Then
                            InsCoverage = "50000"
                        ElseIf iPlanNo = "B" Then
                            InsCoverage = "50000"
                        ElseIf iPlanNo = "C" Or iPlanNo = "M" Then
                            InsCoverage = "200000"
                        ElseIf iPlanNo = "D" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsCoverage = "200000"
                        ElseIf iPlanNo = "J" Then
                            InsCoverage = "3000000"
                        ElseIf iPlanNo = "K" Then
                            InsCoverage = "5000000"
                        Else
                            InsCoverage = "0"
                            '20250124 方案不再控制內直接踢掉,往後處理會造成麻煩（由於修改是單一筆但檢核方案確是一整批(要相同方案)）;多筆改到起痟。好幾代前可以選不同方案
                            Throw New Exception($"[{iPlanNo}]方案非範圍內選項。請重新修改。")
                        End If
                        InsMdeduct = "2500"
                        '2016.2.16 (5集體國內 6集體國外) CHK5
                        'ITP費率檔(個人集體直接) 1啟用 0停用 
                        If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

                            If CHK5.Checked = True Then
                                If TRAVEL_A.SelectedValue = "400" Then
                                    CountryType = "5" '國外5 外6 別
                                Else
                                    CountryType = "6" '國外5 外6 別
                                End If
                            Else
                                If TRAVEL_A.SelectedValue = "400" Then
                                    CountryType = "1" '國內1 外2 別
                                Else
                                    CountryType = "2" '國內1 外2 別
                                End If
                            End If
                        Else
                            If TRAVEL_A.SelectedValue = "400" Then
                                CountryType = "1" '國內1 外2 別
                            Else
                                CountryType = "2" '國內1 外2 別
                            End If
                        End If
                        Dim sumsumIT11 As Integer
                        Dim sumNewIT11 As Integer
                        sumsumIT11 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT11, OID, iPlanNo)
                        sumNewIT11 = sumsumIT11

                        '=====IT2R=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Then
                            InsType = "IT2R"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "1000000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "3000000"
                            ElseIf iPlanNo = "J" Or iPlanNo = "K" Then
                                InsCoverage = "2500000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT2R As Integer
                            Dim sumNewIT2R As Integer
                            If IT2RFree Then
                                sumsumIT2R = InsertIQPOBJ_EXL_Free(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2R, OID)
                            Else
                                sumsumIT2R = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2R, OID, iPlanNo)
                            End If
                            sumNewIT2R = sumsumIT2R
                        End If
                        '=====IT37=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT37"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "10000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Or iPlanNo = "O" Then
                                InsCoverage = "30000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "120000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "150000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT37 As Integer
                            Dim sumNewIT37 As Integer
                            sumsumIT37 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT37, OID, iPlanNo)
                            sumNewIT37 = sumsumIT37
                        End If
                        '=====IT36=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT36"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "20000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Or iPlanNo = "O" Then
                                InsCoverage = "60000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "120000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "150000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT36 As Integer
                            Dim sumNewIT36 As Integer
                            sumsumIT36 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT36, OID, iPlanNo)
                            sumNewIT36 = sumsumIT36
                        End If
                        '=====IT3A=====
                        If iPlanNo <> "A" Then
                            InsType = "IT3A"
                            If iPlanNo = "B" Then
                                InsCoverage = "1000"
                            ElseIf iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "O" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT3A As Integer
                            Dim sumNewIT3A As Integer
                            sumsumIT3A = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT3A, OID, iPlanNo)
                            sumNewIT3A = sumsumIT3A
                        End If
                        '=====IT3B=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT3B"
                            If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                                InsCoverage = "3000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT3B As Integer
                            Dim sumNewIT3B As Integer
                            sumsumIT3B = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT3B, OID, iPlanNo)
                            sumNewIT3B = sumsumIT3B
                        End If
                        '=====IT3C=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "M" Or iPlanNo = "N" Then
                            InsType = "IT3C"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "10000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "30000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT3C As Integer
                            Dim sumNewIT3C As Integer
                            sumsumIT3C = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT3C, OID, iPlanNo)
                            sumNewIT3C = sumsumIT3C
                        End If
                        '=====IT3D=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT3D"
                            If iPlanNo = "C" Or iPlanNo = "M" Or iPlanNo = "O" Then
                                InsCoverage = "1000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "2000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT3D As Integer
                            Dim sumNewIT3D As Integer
                            sumsumIT3D = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT3D, OID, iPlanNo)
                            sumNewIT3D = sumsumIT3D
                        End If
                        '=====IT3E=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "M" Or iPlanNo = "N" Then
                            InsType = "IT3E"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "1000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "2000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT3E As Integer
                            Dim sumNewIT3E As Integer
                            sumsumIT3E = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT3E, OID, iPlanNo)
                            sumNewIT3E = sumsumIT3E
                        End If
                        '=====IT39=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT39"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "O" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT39 As Integer
                            Dim sumNewIT39 As Integer
                            sumsumIT39 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT39, OID, iPlanNo)
                            sumNewIT39 = sumsumIT39

                        End If
                        '=====IT38=====
                        If iPlanNo <> "A" Then
                            InsType = "IT38"
                            If iPlanNo = "B" Then
                                InsCoverage = "1500"
                            ElseIf iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "4000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "O" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT38 As Integer
                            Dim sumNewIT38 As Integer
                            sumsumIT38 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT38, OID, iPlanNo)
                            sumNewIT38 = sumsumIT38
                        End If
                        '=====IT2M=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT2M"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "6000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Or iPlanNo = "O" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT2M As Integer
                            Dim sumNewIT2M As Integer
                            sumsumIT2M = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2M, OID, iPlanNo)
                            sumNewIT2M = sumsumIT2M
                        End If
                        '=====IT27=====
                        If iPlanNo <> "A" And iPlanNo <> "O" Then
                            InsType = "IT27"
                            If iPlanNo = "B" Then
                                InsCoverage = "2000"
                            ElseIf iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                                InsCoverage = "5000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "6000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT27 As Integer
                            Dim sumNewIT27 As Integer
                            sumsumIT27 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT27, OID, iPlanNo)
                            sumNewIT27 = sumsumIT27
                        End If
                        '=====IT2D=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT2D"
                            If iPlanNo = "C" Or iPlanNo = "M" Then
                                InsCoverage = "20000"
                            ElseIf iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "N" Then
                                InsCoverage = "30000"
                            ElseIf iPlanNo = "K" Or iPlanNo = "O" Then
                                InsCoverage = "50000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT2D As Integer
                            Dim sumNewIT2D As Integer
                            sumsumIT2D = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2D, OID, iPlanNo)
                            sumNewIT2D = sumsumIT2D
                        End If
                        '=====IT28=====
                        InsType = "IT28"
                        If iPlanNo = "A" Or iPlanNo = "B" Or iPlanNo = "C" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "O" Then
                            InsCoverage = "3000"
                        ElseIf iPlanNo = "D" Or iPlanNo = "N" Then
                            InsCoverage = "5000"
                        End If
                        InsMdeduct = "0"

                        Dim sumsumIT28 As Integer
                        Dim sumNewIT28 As Integer
                        sumsumIT28 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT28, OID, iPlanNo)
                        sumNewIT28 = sumsumIT28
                        '=====IT2G=====
                        If iPlanNo = "A" Or iPlanNo = "B" Then
                            InsType = "IT2G"
                            InsCoverage = "3000"
                            InsMdeduct = "0"

                            Dim sumsumIT2G As Integer
                            Dim sumNewIT2G As Integer
                            sumsumIT2G = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2G, OID, iPlanNo)
                            sumNewIT2G = sumsumIT2G
                        End If
                        '=====IT31=====
                        If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                            InsType = "IT31"
                            If iPlanNo = "C" Or iPlanNo = "D" Or iPlanNo = "M" Or iPlanNo = "N" Or iPlanNo = "O" Then
                                InsCoverage = "50000"
                            ElseIf iPlanNo = "J" Then
                                InsCoverage = "30000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "50000"
                            End If
                            InsMdeduct = "0"
                            Dim sumsumIT31 As Integer
                            Dim sumNewIT31 As Integer
                            sumsumIT31 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT31, OID, iPlanNo)
                            sumNewIT31 = sumsumIT31
                        End If
                        '=====IT34=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT34"
                            If iPlanNo = "J" Then
                                InsCoverage = "120000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "150000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT34 As Integer
                            Dim sumNewIT34 As Integer
                            sumsumIT34 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT34, OID, iPlanNo)
                            sumNewIT34 = sumsumIT34
                        End If
                        '=====IT35=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT35"
                            If iPlanNo = "J" Then
                                InsCoverage = "120000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "150000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT35 As Integer
                            Dim sumNewIT35 As Integer
                            sumsumIT35 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT35, OID, iPlanNo)
                            sumNewIT35 = sumsumIT35
                        End If
                        '=====IT2P=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT2P"
                            If iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT2P As Integer
                            Dim sumNewIT2P As Integer
                            sumsumIT2P = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2P, OID, iPlanNo)
                            sumNewIT2P = sumsumIT2P
                        End If
                        '=====IT2U====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT2U"
                            If iPlanNo = "J" Then
                                InsCoverage = "10000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "10000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT2U As Integer
                            Dim sumNewIT2U As Integer
                            sumsumIT2U = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2U, OID, iPlanNo)
                            sumNewIT2U = sumsumIT2U
                        End If
                        '=====IT2T=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT2T"
                            If iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT2T As Integer
                            Dim sumNewIT2T As Integer
                            sumsumIT2T = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2T, OID, iPlanNo)
                            sumNewIT2T = sumsumIT2T
                        End If
                        '=====IT2S=====
                        If iPlanNo = "J" Or iPlanNo = "K" Then
                            InsType = "IT2S"
                            If iPlanNo = "J" Then
                                InsCoverage = "3000"
                            ElseIf iPlanNo = "K" Then
                                InsCoverage = "5000"
                            End If
                            InsMdeduct = "0"

                            Dim sumsumIT2S As Integer
                            Dim sumNewIT2S As Integer
                            sumsumIT2S = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2S, OID, iPlanNo)
                            sumNewIT2S = sumsumIT2S
                        End If
                        ''=====IT2N=====
                        'If iPlanNo = "D" Then
                        '    InsType = "IT2N"
                        '    InsCoverage = "3000"
                        '    InsMdeduct = "0"
                        '    InsertIQPOBJ_EXL(dr("IISSU"), InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2N)
                        'End If
#End Region
                    End If

                    Dim TOTPREM As Integer = 0
                    StateMain.AcceptChanges()
                    For Each dr9 As DataRow In StateMain.Tables("IQPOBJ").Select("IISSU = '" & dr("IISSU") & "' ")
                        'drD.Delete()
                        TOTPREM += CInt(dr9("INSPREM"))
                    Next
                    IQPMain_totPrem += TOTPREM
                    dr("ISSU_4") = TOTPREM

                    BaseSetChildFormData(Me.dgdMain, Me.StateMain)

                    i = i + 1
                    ALLsbErr &= sbErr
                End If
            Next

            '20250908 IQPMAIN給當前總保費，確保後續procChangedTargetProject()將正確的總保費寫入資料表。如此在要保確認前檢視的資料才會一致
            If IQPMain_totPrem > 0 Then
                StateMain.Tables("IQPMAIN").Rows(0)("TOTPREM") = IQPMain_totPrem.ToString()
            End If

            StateMain.AcceptChanges()
            txtIQPMainOid.Text = OID

            If Me.IsInputErr Then
                InsertUP_BTN.Text = "X"
                CheckSNNN.Text = "資料有誤, 請修正後重新檢核"
                sbErrAll = "資料有誤, 請修正後重新檢核"
                txtCase.Text = GetiPlanNo(txtCase.Text)
                Me.UIAppendError(ALLsbErr)

            Else
                If dsCopy.Tables("IQPISSU").Rows.Count = 0 Then
                    Me.UIAppendMessage("無資料可供檢核!")
                Else

                    InsertUP_BTN.Text = ""
                    sbErrAll = ""
                    CheckSNNN.Text = "檢核資料正確"

                End If
            End If

        Catch ex As Exception
            InsertUP_BTN.Text = "X"
            ProcessException(ex)
        End Try
    End Sub
    Protected Overrides Sub BaseDoUpdate()
        '修改
        Dim dsMain As DataSet = StateMain.Copy()
        Dim dr As DataRow = BaseGetMainRow(dsMain)
        AppCollectMainFormData(dr)
        'If Not dsMain.HasChanges Then
        '    UIFormIsDirty = False
        '    BaseMsgSaveNoChangeForm()
        '    Return
        'End If
        BaseUpdate(dsMain.Copy())
        If BaseIsClearFormAfterUpdate() Then
            BaseDoClear()
        End If
        BaseMsgUpdateOK()
    End Sub

    Public Property WebServiceURL() As String
        Get
            Return CType(PageSession("_WEBSERVICEURL"), String)
        End Get
        Set(ByVal Value As String)
            PageSession("_WEBSERVICEURL") = Value
        End Set
    End Property

    ''' <summary>
    ''' WebService false=測試機 ; true=正式機
    ''' </summary>
    ''' <returns></returns>
    Public Property WebServiceFlag() As Boolean
        Get
            Return CType(PageSession("_WEBSERVICEFLAG"), Boolean)
        End Get
        Set(ByVal Value As Boolean)
            PageSession("_WEBSERVICEFLAG") = Value
        End Set
    End Property

    Function CallCommonService(ByVal rqtMISServiceName As String, ByVal args As String())
        Dim iTimerOut As Double = 0
        Dim dsTmp As New DataSet
        WebServiceFlag = CBool(ConfigurationManager.AppSettings("bHASB2CRequest"))
        Dim ws As B2BWebReference.BusinessService = InitsWS_IQP(GetWsURL_IQP(WebServiceFlag), CDbl(iTimerOut))

        Try
            'Dim ws As B2BWebReference.BusinessService = InitsWS_IQP(WebServiceURL, CDbl(iTimerOut))

            dsTmp = ws.InvokeCommonService(rqtMISServiceName, args)
            'ws.Dispose()
        Catch ex As Exception
            Throw ex
        Finally
            ws.Dispose()
        End Try

        Return dsTmp
    End Function
    Function CallBussinessService(ByVal rqtMISServiceName As String, ByVal ds As DataSet)
        Dim iTimerOut As Double = 0
        Dim dsTmp As New DataSet
        Try
            'Dim ws As B2BWebReference.BusinessService = InitsWS_IQP(WebServiceURL, CDbl(iTimerOut))
            Dim ws As B2BWebReference.BusinessService = InitsWS_IQP(GetWsURL_IQP(CBool(ConfigurationManager.AppSettings("bHASB2CRequest"))), CDbl(iTimerOut))

            dsTmp = ws.InvokeBusinessService(rqtMISServiceName, ds)
            ws.Dispose()
        Catch ex As Exception
            Throw ex
        End Try

        Return dsTmp
    End Function
    Public Function BrokerCheck() As String
        '通路檢核
        Dim param As New PUBEntity.bePubBrokerCheck
        Dim dr As PUBEntity.bePubBrokerCheck.BrokerCheckRow = param.BrokerCheck.NewBrokerCheckRow
        Dim strMsg As String = ""
        '經紀人代號(IAGENT)             ==> 通路經辦代號轉經紀人
        '經辦代號(IOFFICER)             ==> 通路經辦代號
        '險別代號(IINSCLS)              ==> I
        '大宗收件號(IREC)               ==> 畫面欄位
        '通路單位(IRECUNIT)             ==> 畫面欄位
        '通路單位名稱(NRECUNIT)         ==> 畫面欄位
        '保單號碼(IPOLICY)              ==> ""
        '類別(FLAG) 0:保單 1:批單       ==> 0
        '險種(IINSKIND)                 ==> ITP OR IQP
        '續保單號(IPREPLY)              ==> ""
        '是否分期(NRECRMK1)             ==> N
        '通路人員代號(IRECPESN)         ==> 畫面欄位
        '通路人員名稱(NRECPESN)         ==> 畫面欄位



        '找出該帳號經辦代號
        Dim IOFFICER As String = ""
        Dim IAGENT As String = ""
        Dim TextBoxUserID As String = GetUserIDFromTextBox()
        Dim ds0 As DataSet = Me.InvokeCommonService("B2BService.bsGetItemLists", "b2b.getiofficer", TextBoxUserID)
        If ds0.Tables("Organization").Rows.Count > 0 Then
            IOFFICER = ds0.Tables("Organization").Rows(0).Item("Iofficer")
        End If

        '串出經紀人
        Dim dsRtn As DataSet = Me.CallCommonService("PUBService.bsSingleSQL", New String() {"CMNP_OFF_SOURCE", "IOFFICER ='" + IOFFICER + "'", "IAGENT"})
        If dsRtn.Tables(0).DefaultView.Count > 0 Then
            IAGENT = dsRtn.Tables(0).Rows(0)("IAGENT").trim
        End If

        'If Me.ScsOrgID = "K077" Or Me.ScsOrgID = "KA77" Or Me.ScsOrgID = "KKEW" Then
        If IAGENT = "855" Or IAGENT = "869" Or IAGENT = "293" Or IAGENT = "957" Or IAGENT = "028" Or IAGENT = "728" Or IAGENT = "589" Or IAGENT = "970" Or IAGENT = "923" Then


            If IAGENT = "855" Or IAGENT = "293" Then
                If Me.txtIREC.Text.Trim = "" Then  '收件號
                    strMsg &= "通路檢核 需輸入收件號." & Environment.NewLine
                End If
                If Me.txtIrecunit.Text.Trim = "" Then '通路單位人員
                    strMsg &= "通路檢核 需輸入通路單位." & Environment.NewLine
                End If
                'If IAGENT = "855" Then
                '    txtIREC.Text = "2017000XXXXXX"
                'End If
            End If
            If IAGENT = "869" Or IAGENT = "957" Or IAGENT = "970" Or IAGENT = "923" Then
                If Me.txtIrecunit.Text.Trim = "" Then '通路單位人員
                    strMsg &= "通路檢核 需輸入通路單位." & Environment.NewLine
                End If
            End If

            If IAGENT = "589" Then
                If Me.txtIRECPESN.Text.Trim = "" Then '通路人員代號
                    strMsg &= "「通路人員代號」需登打業務員登錄證號" & Environment.NewLine
                End If
            End If

            If IAGENT = "728" Then
                If Me.txtIrecunit.Text.Trim = "" Then '通路單位人員
                    strMsg &= "「通路單位」需登打業務員登錄證號及業務員姓名" & Environment.NewLine
                End If
            End If

            If IAGENT = "028" Then
                If Me.txtIrecunit.Text.Trim = "" Then '通路單位人員
                    strMsg &= "「通路單位」需登打業務員登錄證號" & Environment.NewLine
                End If

                If txtNrecpesn.Text.Trim = "" Then
                    strMsg &= "「通路人員名稱」需登打業務員姓名" & Environment.NewLine
                End If
            End If

            If strMsg = "" Then
                dr("IAGENT") = IAGENT '經紀人代號
                dr("IOFFICER") = IOFFICER '經辦人代號
                dr("IINSCLS") = "I" '險別代號
                dr("IREC") = Me.txtIREC.Text '收件號
                dr("IRECUNIT") = Me.txtIrecunit.Text '通路單位人員
                dr("NRECUNIT") = Me.txtNRECUNIT.Text '通路單位名稱
                dr("IPOLICY") = ""  '保單號碼
                dr("FLAG") = "0"  '0:保單 1:批單
                dr("IINSKIND") = "ITP" '險種
                dr("IPREPLY") = "" '續保單號
                dr("NRECRMK1") = "N" '是否分期
                dr("IRECPESN") = Me.txtIRECPESN.Text '通路人員代號
                dr("NRECPESN") = Me.txtNrecpesn.Text '通路人員名稱

                param.BrokerCheck.Rows.Add(dr)

                Dim ds As DataSet = Me.CallBussinessService("PUBService.bsPubBrokerCheck", param)
                'Dim aryNG() As DataRow = ds.Tables("BrokerCheck").Select("RESULT='N'")
                'For Each dr In aryNG
                'strMsg &= dr("MSG").ToString.Trim & vbNewLine
                'Next
                strMsg &= ds.Tables("BrokerCheck").Rows(0).Item("msg")

            End If

        End If

        Return strMsg

    End Function
    Protected Sub btnIINS2_Click(sender As Object, e As EventArgs) Handles btnIINS2.Click
        Try
            Dim entity = GetClient(txtIins_IINS2.Text)
            If entity.name <> String.Empty Then
                NAME_IINS.Text = entity.name
                IINS_ADDR.Text = entity.address
            Else
                Throw New Exception("查無資料")
            End If
        Catch ex As Exception
            Me.UIAppendError(ex.Message)
        End Try
    End Sub

    Public Function GetClient(id As String) As Has101000
        If id = String.Empty Then
            Throw New Exception("請輸入要保人ID")
        End If
        If id.Length = 8 Then
            Return New Has101000(id)
        Else Return New Has101000()
        End If


    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Shows3.Visible = True
        Me.Shows4.Visible = True
    End Sub

    Private Sub btnOpenFlight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenFlight.Click
        If btnOpenFlight.Text = "+" Then
            btnOpenFlight.Text = "-"
            Me.Show5.Visible = True
            Me.Show6.Visible = True
            Me.Show7.Visible = True
        Else
            btnOpenFlight.Text = "+"
            Me.Show5.Visible = False
            Me.Show6.Visible = False
            Me.Show7.Visible = False
        End If
    End Sub

    Function CallBussinessServiceSKL_GETIREC()
        'K077統一改抓車險 bsGetReceiveNumber
        Try
            Dim dsRtn As New DataSet
            Dim UnitID, Introducer, EmpolyeeNumber, PolicyKind, PolicyType, CustomerType As String
            UnitID = txtIrecunit.Text.Trim.Substring(0, 6)
            Introducer = txtIrecunit.Text.Trim.Substring(6, 6)
            EmpolyeeNumber = txtIrecunit.Text.Trim.Substring(6, 6)
            PolicyKind = "5"
            PolicyType = "1"
            If ddl_IINS.SelectedValue = "2" Then
                CustomerType = "2" '統一編號
            ElseIf ddl_IINS.SelectedValue = "3" Then
                CustomerType = "3" '居留證號
            Else
                CustomerType = "1" '身份證字號
            End If

            '====================================================================================================================
            Dim url = ""
            Dim urlGetToken = ""
            Dim K077NewKey = ""
            Dim K077NewIV = ""
            Dim dsK077 As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", $"TYPE = 'K077_PlatForm' AND CODE='F1920'", "DATA1,DATA2,DATA3"})
            If Not IsNothing(dsK077) AndAlso dsK077.Tables(0).Rows.Count > 0 Then
                If dsK077.Tables(0).Rows(0)("DATA1") = "1" Then
                    K077NewKey = dsK077.Tables(0).Rows(0)("DATA2").ToString.Trim
                    K077NewIV = dsK077.Tables(0).Rows(0)("DATA3").ToString.Trim
                    Dim dsSingK077 As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", $"TYPE = 'BRK001000' AND CODE='K077New'", "DATA1,DATA3"})
                    If Not IsNothing(dsSingK077) AndAlso dsSingK077.Tables(0).Rows.Count > 0 Then
                        K077NewFlag = True
                        url = GetString(dsSingK077.Tables(0).Rows(0)("DATA1")) 'SSO網址
                        urlGetToken = GetString(dsSingK077.Tables(0).Rows(0)("DATA3")) '新光金新系統取得TOKEN網址
                    Else
                        Throw New Exception("找不到對應的網址!")
                    End If
                End If
            End If
            If K077NewFlag = True Then
                K077Token = B2BEntity.HasFunction.DecryptCFBK077(K077GetToken(urlGetToken, K077NewKey, K077NewIV), K077NewKey, K077NewIV)
                K077Token = K077Token.Substring(0, K077Token.LastIndexOf("}") + 1)
                Dim Obj As Linq.JObject = Newtonsoft.Json.JsonConvert.DeserializeObject(K077Token)
                HeaderToken = Obj.Item("JwtToken").ToString
            End If
            '====================================================================================================================
            'dsRtn = Me.InvokeCommonService("B2BService.bsBCAR855API", New String() {"GetReceiveNumber", UnitID, Introducer, EmpolyeeNumber, PolicyKind, PolicyType, CustomerType, Me.txtIins_IINS2.Text.Trim, HeaderToken})
            dsRtn = Me.InvokeCommonService("B2BService.bsGetReceiveNumber", New String() {"GetReceiveNumberITP", UnitID, Introducer, EmpolyeeNumber, PolicyKind, PolicyType, CustomerType, Me.txtIins_IINS2.Text.Trim, HeaderToken})
            If dsRtn.Tables(0).Rows.Count > 0 AndAlso dsRtn.Tables(0).Rows(0).Item("ReturnCode").ToString = "00" Then
                Me.txtIREC.Text = dsRtn.Tables(0).Rows(0).Item("ReceiveNumber").ToString
            ElseIf dsRtn.Tables(0).Rows.Count > 0 Then
                Throw New Exception("呼叫自動產生收件編號錯誤 ex:" & dsRtn.Tables(0).Rows(0).Item("ReturnStatus").ToString & dsRtn.Tables(0).Rows(0).RowError)
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function InitsWS(ByVal sWsUrl As String, ByVal iTimerOut As Double)
        Dim ws As New skpiaweb.ReceiveNumberWS
        '若有傳遞URL就用傳入的URL當來源WS
        If iTimerOut = 0 Then
            ws.Timeout = 60000
        Else
            ws.Timeout = iTimerOut
        End If
        ws.Credentials = System.Net.CredentialCache.DefaultCredentials '排除http error 401問題
        If sWsUrl <> "" Then
            ws.Url = sWsUrl
        End If
        Return ws
    End Function
    '轉數字為0
    Protected Function GetDec(ByVal d As Object) As Decimal
        If IsDBNull(d) Then
            Return 0
        ElseIf d Is Nothing Then
            Return 0
        ElseIf d.ToString.Trim.Length = 0 Then
            Return 0
        ElseIf Not IsNumeric(d) Then
            Return 0
        Else
            Return CDec(d)
        End If
    End Function
    Function checklock()
        '新壽需要鎖控團體費率
        'K105也有
        If Me.txtCompanyid.Text = "K105" Or Me.txtCompanyid.Text = "K077" Or Me.txtCompanyid.Text = "KKKK" Or Me.txtCompanyid.Text = "KKEW" Or
           Me.txtCompanyid.Text = "KKFN" Or Me.txtCompanyid.Text = "KKFP" Or Me.txtCompanyid.Text = "KKEX" Or Me.txtCompanyid.Text = "KJ48" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Class TrustAllCertificatePolicy
        Implements ICertificatePolicy

        Public Function CheckValidationResult(ByVal srvPoint As ServicePoint,
                         ByVal cert As X509Certificate, ByVal request As WebRequest,
                         ByVal certificateProblem As Integer) _
                     As Boolean Implements ICertificatePolicy.CheckValidationResult
            Return True
        End Function
    End Class
    Function CheckAreaDply_bgn() As Boolean
        Dim str_CHK_ID As String() = {"ddlCountry_A2", "ddlCountry_B2", "ddlCountry_C2", "ddlCountry_D2", "ddlCountry_E2", "ddlCountry_F2", "ddlCountry_G2", "ddlCountry_H2"}
        Dim CHK_AREA_COUNTRY As String() = {"JP720"}
        Dim conddl As New Control
        'conddl = FindControl(s1)

        For Each AREA_COUNTRY As String In CHK_AREA_COUNTRY
            '國家
            For Each CHK_ID As String In str_CHK_ID
                '檢查各選項
                If CType(FindControl(CHK_ID), DropDownList).SelectedValue = AREA_COUNTRY Then
                    Dim DATE_BGN As DateTime = CType(CDate(UcRocDobjbgn.DateText).ToString("yyyy/MM/dd") & " " & txtTobjbgn.Text & ":00", DateTime)
                    Dim DATEts As TimeSpan = DATE_BGN - Now
                    If DATEts.Days < 30 Then 'If DATEts.Days > 5 And DATEts.Days < 15 Then
                        Return False
                    Else
                        Return True
                    End If
                End If
            Next
        Next
        Return False
    End Function
    Function CheckAreaLock() As String
        Try
            Dim strErr As String = ""
            Me.IsInputErr = False
            Dim TCODE As String = ""
            TCODE = TRAVEL_A.SelectedValue.Trim
            '400 台灣地區及離島地區
            '500 大陸地區(包含港、澳)
            '600 美加地區
            '710 東南亞、南亞地區
            '720 日本地區
            '725 韓國地區
            '730 紐澳地區
            '740 歐洲地區
            '750 申根地區
            '760 非洲地區
            '770 中東地區
            '800 全球各地
            '810 東南亞地區

            Dim AREALOCK As String = ""
            Dim AREALOCK2 As String = ""
            Dim dsSing As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='HAS_B2BAREALOCK' AND CODE = '" + TCODE + "'", "DATA1,DATA2,DATA3"})

            If Not IsNothing(dsSing) AndAlso dsSing.Tables.Count > 0 AndAlso dsSing.Tables(0).Rows.Count > 0 Then
                AREALOCK = dsSing.Tables(0).Rows(0)("DATA1").ToString.Trim
                AREALOCK2 = dsSing.Tables(0).Rows(0)("DATA2").ToString.Trim
            End If

            If AREALOCK <> "" Then
                If (UcRocDobjbgn.DateText >= CDate(AREALOCK) And UcRocDobjbgn.DateText <= CDate(AREALOCK2)) Then
                    Me.IsInputErr = True
                    strErr &= dsSing.Tables(0).Rows(0)("DATA3").ToString.Trim
                End If
            End If

            Return strErr
        Catch ex As Exception
            Me.ProcessException(ex)
        End Try
    End Function
    Function CHECK_COUNCODE1() As String
        Try
            Dim COUNCODE1 As String = ""
            If ddlCountry_B2.SelectedValue = "" And ddlCountry_C2.SelectedValue = "" And ddlCountry_D2.SelectedValue = "" And
                ddlCountry_E2.SelectedValue = "" And ddlCountry_F2.SelectedValue = "" And ddlCountry_G2.SelectedValue = "" And
                ddlCountry_H2.SelectedValue = "" Then

                '只有一組 用第一組判斷就好
                COUNCODE1 = HttpUtility.HtmlEncode(ddlCountry_A2.SelectedValue.Substring(2, 3))
                If COUNCODE1 = "750" Then
                    '申根地區判斷
                    If Checkbox2.Checked = True Then
                        '申根地區
                        COUNCODE1 = "750"
                    Else
                        '改成歐洲地區
                        COUNCODE1 = "740"
                    End If
                End If
            Else
                Dim ddlA, ddlB, ddlC, ddlD, ddlE, ddlF, ddlG, ddlH As String

                If ddlCountry_A2.SelectedValue <> "" Then
                    ddlA = HttpUtility.HtmlEncode(ddlCountry_A2.SelectedValue.Substring(2, 3))
                Else
                    ddlA = ""
                End If
                If ddlCountry_B2.SelectedValue <> "" Then
                    ddlB = HttpUtility.HtmlEncode(ddlCountry_B2.SelectedValue.Substring(2, 3))
                Else
                    ddlB = ""
                End If
                If ddlCountry_C2.SelectedValue <> "" Then
                    ddlC = HttpUtility.HtmlEncode(ddlCountry_C2.SelectedValue.Substring(2, 3))
                Else
                    ddlC = ""
                End If
                If ddlCountry_D2.SelectedValue <> "" Then
                    ddlD = HttpUtility.HtmlEncode(ddlCountry_D2.SelectedValue.Substring(2, 3))
                Else
                    ddlD = ""
                End If
                If ddlCountry_E2.SelectedValue <> "" Then
                    ddlE = HttpUtility.HtmlEncode(ddlCountry_E2.SelectedValue.Substring(2, 3))
                Else
                    ddlE = ""
                End If
                If ddlCountry_F2.SelectedValue <> "" Then
                    ddlF = HttpUtility.HtmlEncode(ddlCountry_F2.SelectedValue.Substring(2, 3))
                Else
                    ddlF = ""
                End If
                If ddlCountry_G2.SelectedValue <> "" Then
                    ddlG = HttpUtility.HtmlEncode(ddlCountry_G2.SelectedValue.Substring(2, 3))
                Else
                    ddlG = ""
                End If
                If ddlCountry_H2.SelectedValue <> "" Then
                    ddlH = HttpUtility.HtmlEncode(ddlCountry_H2.SelectedValue.Substring(2, 3))
                Else
                    ddlH = ""
                End If

                If ddlA = "400" Then
                    If (ddlB <> "400" And ddlB <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlC <> "400" And ddlC <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlD <> "400" And ddlD <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlE <> "400" And ddlE <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlF <> "400" And ddlF <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlG <> "400" And ddlG <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    If (ddlH <> "400" And ddlH <> "") Then
                        Throw New Exception("選擇台灣地區時無法再選擇國外區域。")
                    End If
                    COUNCODE1 = "400"
                Else
                    '全球各地
                    If ddlA = "750" Or ddlB = "750" Or ddlC = "750" Or ddlD = "750" Or ddlE = "750" Or ddlF = "750" Or ddlG = "750" Or ddlH = "750" Then
                        '申根地區判斷
                        If Checkbox2.Checked = True Then
                            '申根地區
                            COUNCODE1 = "750"
                        Else
                            '改成歐洲地區
                            COUNCODE1 = "740"
                        End If
                    Else
                        If ddlB <> "" Then
                            If String.Compare(ddlA, ddlB) = 0 Then
                                COUNCODE1 = ddlA
                            Else
                                '改成全球地區
                                COUNCODE1 = "800"
                            End If
                        End If
                        If ddlC <> "" And COUNCODE1 <> "800" Then
                            If String.Compare(ddlA, ddlC) = 0 Then
                                COUNCODE1 = ddlA
                            Else
                                '改成全球地區
                                COUNCODE1 = "800"
                            End If
                        End If
                        If ddlD <> "" And COUNCODE1 <> "800" Then
                            If String.Compare(ddlA, ddlD) = 0 Then
                                COUNCODE1 = ddlA
                            Else
                                '改成全球地區
                                COUNCODE1 = "800"
                            End If
                        End If
                        If ddlE <> "" And COUNCODE1 <> "800" Then
                            If String.Compare(ddlA, ddlE) = 0 Then
                                COUNCODE1 = ddlA
                            Else
                                '改成全球地區
                                COUNCODE1 = "800"
                            End If
                        End If
                        If ddlF <> "" And COUNCODE1 <> "800" Then
                            If String.Compare(ddlA, ddlF) = 0 Then
                                COUNCODE1 = ddlA
                            Else
                                '改成全球地區
                                COUNCODE1 = "800"
                            End If
                        End If
                        If ddlG <> "" And COUNCODE1 <> "800" Then
                            If String.Compare(ddlA, ddlG) = 0 Then
                                COUNCODE1 = ddlA
                            Else
                                '改成全球地區
                                COUNCODE1 = "800"
                            End If
                        End If
                        If ddlH <> "" And COUNCODE1 <> "800" Then
                            If String.Compare(ddlA, ddlH) = 0 Then
                                COUNCODE1 = ddlA
                            Else
                                '改成全球地區
                                COUNCODE1 = "800"
                            End If
                        End If

                    End If
                End If
            End If

            Return COUNCODE1

        Catch ex As Exception
            Me.ProcessException(ex)
        End Try
    End Function
    Function CheckTyphoon() As String
        Try
            Dim msg As String = ""
            Dim Region As String = ""
            Dim country As String = ""

            If TRAVEL_A.SelectedValue = "400" Then
                Region = "1"
            Else
                Region = "2"
            End If
            country = ddlCountry_A2.SelectedValue.Substring(0, 2).Trim
            Dim dsReturn As DataSet = Me.InvokeCommonService("B2BService.bsB2BCheckTyphoon", UcRocDobjbgn.DateText, txtTobjbgn.Text.Trim, UcRocDobjend.DateText, txtTobjend.Text.Trim, Region, country)
            If dsReturn.Tables("CheckB2B").Rows(0)("PASS") = "N" Then
                msg = dsReturn.Tables("CheckB2B").Rows(0)("MSG")
            End If
            If ddlCountry_B2.SelectedValue <> "" Then
                country = ddlCountry_B2.SelectedValue.Substring(0, 2).Trim
                Dim dsReturn2 As DataSet = Me.InvokeCommonService("B2BService.bsB2BCheckTyphoon", UcRocDobjbgn.DateText, txtTobjbgn.Text.Trim, UcRocDobjend.DateText, txtTobjend.Text.Trim, Region, country)
                If dsReturn2.Tables("CheckB2B").Rows(0)("PASS") = "N" Then
                    msg = dsReturn2.Tables("CheckB2B").Rows(0)("MSG")
                End If
            End If
            If ddlCountry_C2.SelectedValue <> "" Then
                country = ddlCountry_C2.SelectedValue.Substring(0, 2).Trim
                Dim dsReturn3 As DataSet = Me.InvokeCommonService("B2BService.bsB2BCheckTyphoon", UcRocDobjbgn.DateText, txtTobjbgn.Text.Trim, UcRocDobjend.DateText, txtTobjend.Text.Trim, Region, country)
                If dsReturn3.Tables("CheckB2B").Rows(0)("PASS") = "N" Then
                    msg = dsReturn3.Tables("CheckB2B").Rows(0)("MSG")
                End If
            End If
            If ddlCountry_D2.SelectedValue <> "" Then
                country = ddlCountry_D2.SelectedValue.Substring(0, 2).Trim
                Dim dsReturn4 As DataSet = Me.InvokeCommonService("B2BService.bsB2BCheckTyphoon", UcRocDobjbgn.DateText, txtTobjbgn.Text.Trim, UcRocDobjend.DateText, txtTobjend.Text.Trim, Region, country)
                If dsReturn4.Tables("CheckB2B").Rows(0)("PASS") = "N" Then
                    msg = dsReturn4.Tables("CheckB2B").Rows(0)("MSG")
                End If
            End If
            If ddlCountry_E2.SelectedValue <> "" Then
                country = ddlCountry_E2.SelectedValue.Substring(0, 2).Trim
                Dim dsReturn5 As DataSet = Me.InvokeCommonService("B2BService.bsB2BCheckTyphoon", UcRocDobjbgn.DateText, txtTobjbgn.Text.Trim, UcRocDobjend.DateText, txtTobjend.Text.Trim, Region, country)
                If dsReturn5.Tables("CheckB2B").Rows(0)("PASS") = "N" Then
                    msg = dsReturn5.Tables("CheckB2B").Rows(0)("MSG")
                End If
            End If
            If ddlCountry_F2.SelectedValue <> "" Then
                country = ddlCountry_F2.SelectedValue.Substring(0, 2).Trim
                Dim dsReturn6 As DataSet = Me.InvokeCommonService("B2BService.bsB2BCheckTyphoon", UcRocDobjbgn.DateText, txtTobjbgn.Text.Trim, UcRocDobjend.DateText, txtTobjend.Text.Trim, Region, country)
                If dsReturn6.Tables("CheckB2B").Rows(0)("PASS") = "N" Then
                    msg = dsReturn6.Tables("CheckB2B").Rows(0)("MSG")
                End If
            End If
            If ddlCountry_G2.SelectedValue <> "" Then
                country = ddlCountry_G2.SelectedValue.Substring(0, 2).Trim
                Dim dsReturn7 As DataSet = Me.InvokeCommonService("B2BService.bsB2BCheckTyphoon", UcRocDobjbgn.DateText, txtTobjbgn.Text.Trim, UcRocDobjend.DateText, txtTobjend.Text.Trim, Region, country)
                If dsReturn7.Tables("CheckB2B").Rows(0)("PASS") = "N" Then
                    msg = dsReturn7.Tables("CheckB2B").Rows(0)("MSG")
                End If
            End If
            If ddlCountry_H2.SelectedValue <> "" Then
                country = ddlCountry_H2.SelectedValue.Substring(0, 2).Trim
                Dim dsReturn8 As DataSet = Me.InvokeCommonService("B2BService.bsB2BCheckTyphoon", UcRocDobjbgn.DateText, txtTobjbgn.Text.Trim, UcRocDobjend.DateText, txtTobjend.Text.Trim, Region, country)
                If dsReturn8.Tables("CheckB2B").Rows(0)("PASS") = "N" Then
                    msg = dsReturn8.Tables("CheckB2B").Rows(0)("MSG")
                End If
            End If
            Return msg
        Catch ex As Exception
            Me.ProcessException(ex)
        End Try
    End Function

    Public Sub GET_IQPOBJ(IISSU As String, iPlanNo As String, i As Integer)
        '##############無叫用，方案內容無更新#################

        '險種代號  C 海外基本型 D 海外超值型
        'IT11    限額200000元	限額200000元
        'IT2R    限額150萬	    限額150萬
        'IT37    限額10000元	限額20000元
        'IT36    限額20000元	限額40000元
        'IT3B    每次2000元	    每次3000元
        'IT3A    每次3000元	    每次5000元
        'IT39    每次3000元	    每次5000元
        'IT38    每次3000元	    每次5000元
        'IT2M    每次3000元	    每次6000元
        'IT27    每日3000元	    每日5000元
        'IT2D    限額20000元	限額30000元
        'IT28    每次3000元	    每次5000元
        'IT31    限額50000元	限額100000元

        Dim InsType, InsDay, InsCoverage, InsMdeduct, CountryType As String

        InsType = ""
        InsDay = ""
        InsCoverage = ""
        InsMdeduct = ""
        CountryType = ""

        'ITP 不便險取保費的方式
        '0 IISSU
        '1:InsType(險種)
        '2:InsDay(天數)
        '3.InsCoverage 保額 元
        '4.InsMdeduct 自負額 元
        '5.CountryType 1=國內 2=國外
        '=====IT11=====
        InsType = "IT11"
        InsDay = txtIQPITP_Days.Text
        If iPlanNo = "A" Then
            InsCoverage = "50000"
        ElseIf iPlanNo = "B" Then
            InsCoverage = "50000"
        ElseIf iPlanNo = "C" Then
            InsCoverage = "200000"
        ElseIf iPlanNo = "D" Then
            InsCoverage = "200000"
        End If
        InsMdeduct = "2500"
        '2016.2.16 (5集體國內 6集體國外) CHK5
        'ITP費率檔(個人集體直接) 1啟用 0停用 
        If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then

            If CHK5.Checked = True Then
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "5" '國外5 外6 別
                Else
                    CountryType = "6" '國外5 外6 別
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If
        Else
            If TRAVEL_A.SelectedValue = "400" Then
                CountryType = "1" '國內1 外2 別
            Else
                CountryType = "2" '國內1 外2 別
            End If
        End If
        Dim sumsumIT11 As Integer
        Dim sumNewIT11 As Integer
        sumsumIT11 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT11, "", iPlanNo)
        sumNewIT11 = sumsumIT11

        '=====IT2R=====
        If iPlanNo = "C" Or iPlanNo = "D" Then
            InsType = "IT2R"
            InsCoverage = "1500000"
            InsMdeduct = "0"
            Dim sumsumIT2R As Integer
            Dim sumNewIT2R As Integer
            sumsumIT2R = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2R, "", iPlanNo)
            sumNewIT2R = sumsumIT2R
        End If
        '=====IT37=====
        If iPlanNo = "C" Or iPlanNo = "D" Then
            InsType = "IT37"
            If iPlanNo = "C" Then
                InsCoverage = "10000"
            ElseIf iPlanNo = "D" Then
                InsCoverage = "20000"
            End If
            InsMdeduct = "0"
            Dim sumsumIT37 As Integer
            Dim sumNewIT37 As Integer
            sumsumIT37 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT37, "", iPlanNo)
            sumNewIT37 = sumsumIT37
        End If
        '=====IT36=====
        If iPlanNo = "C" Or iPlanNo = "D" Then
            InsType = "IT36"
            If iPlanNo = "C" Then
                InsCoverage = "20000"
            ElseIf iPlanNo = "D" Then
                InsCoverage = "40000"
            End If
            InsMdeduct = "0"
            Dim sumsumIT36 As Integer
            Dim sumNewIT36 As Integer
            sumsumIT36 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT36, "", iPlanNo)
            sumNewIT36 = sumsumIT36
        End If
        '=====IT3B=====
        If iPlanNo = "C" Or iPlanNo = "D" Then
            InsType = "IT3B"
            If iPlanNo = "C" Then
                InsCoverage = "2000"
            ElseIf iPlanNo = "D" Then
                InsCoverage = "3000"
            End If
            InsMdeduct = "0"
            Dim sumsumIT3B As Integer
            Dim sumNewIT3B As Integer
            sumsumIT3B = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT3B, "", iPlanNo)
            sumNewIT3B = sumsumIT3B
        End If
        '=====IT3A=====
        If iPlanNo <> "A" Then
            InsType = "IT3A"
            If iPlanNo = "B" Then
                InsCoverage = "1000"
            ElseIf iPlanNo = "C" Then
                InsCoverage = "3000"
            ElseIf iPlanNo = "D" Then
                InsCoverage = "5000"
            End If
            InsMdeduct = "0"
            Dim sumsumIT3A As Integer
            Dim sumNewIT3A As Integer
            sumsumIT3A = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT3A, "", iPlanNo)
            sumNewIT3A = sumsumIT3A
        End If
        '=====IT39=====
        If iPlanNo = "C" Or iPlanNo = "D" Then
            InsType = "IT39"
            If iPlanNo = "C" Then
                InsCoverage = "3000"
            ElseIf iPlanNo = "D" Then
                InsCoverage = "5000"
            End If
            InsMdeduct = "0"
            Dim sumsumIT39 As Integer
            Dim sumNewIT39 As Integer
            sumsumIT39 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT39, "", iPlanNo)
            sumNewIT39 = sumsumIT39

        End If
        '=====IT38=====
        If iPlanNo <> "A" Then
            InsType = "IT38"
            If iPlanNo = "B" Then
                InsCoverage = "1500"
            ElseIf iPlanNo = "C" Then
                InsCoverage = "3000"
            ElseIf iPlanNo = "D" Then
                InsCoverage = "5000"
            End If
            InsMdeduct = "0"

            Dim sumsumIT38 As Integer
            Dim sumNewIT38 As Integer
            sumsumIT38 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT38, "", iPlanNo)
            sumNewIT38 = sumsumIT38
        End If
        '=====IT25=====
        If iPlanNo = "C" Or iPlanNo = "D" Then
            InsType = "IT25"
            If iPlanNo = "C" Then
                InsCoverage = "3000"
            ElseIf iPlanNo = "D" Then
                InsCoverage = "6000"
            End If
            InsMdeduct = "0"

            Dim sumsumIT25 As Integer
            Dim sumNewIT25 As Integer
            sumsumIT25 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT25, "", iPlanNo)
            sumNewIT25 = sumsumIT25
        End If
        '=====IT27=====
        If iPlanNo <> "A" Then
            InsType = "IT27"
            If iPlanNo = "B" Then
                InsCoverage = "2000"
            ElseIf iPlanNo = "C" Then
                InsCoverage = "3000"
            ElseIf iPlanNo = "D" Then
                InsCoverage = "5000"
            End If
            InsMdeduct = "0"

            Dim sumsumIT27 As Integer
            Dim sumNewIT27 As Integer
            sumsumIT27 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT27, "", iPlanNo)
            sumNewIT27 = sumsumIT27
        End If
        '=====IT2D=====
        If iPlanNo = "C" Or iPlanNo = "D" Then
            InsType = "IT2D"
            If iPlanNo = "C" Then
                InsCoverage = "20000"
            ElseIf iPlanNo = "D" Then
                InsCoverage = "30000"
            End If
            InsMdeduct = "0"

            Dim sumsumIT2D As Integer
            Dim sumNewIT2D As Integer
            sumsumIT2D = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2D, "", iPlanNo)
            sumNewIT2D = sumsumIT2D
        End If
        '=====IT28=====
        InsType = "IT28"
        If iPlanNo = "A" Or iPlanNo = "B" Or iPlanNo = "C" Then
            InsCoverage = "3000"
        ElseIf iPlanNo = "D" Then
            InsCoverage = "5000"
        End If
        InsMdeduct = "0"

        Dim sumsumIT28 As Integer
        Dim sumNewIT28 As Integer
        sumsumIT28 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT28, "", iPlanNo)
        sumNewIT28 = sumsumIT28

        '=====IT2G=====
        If iPlanNo = "A" Or iPlanNo = "B" Then
            InsType = "IT2G"
            InsCoverage = "3000"
            InsMdeduct = "0"

            Dim sumsumIT2G As Integer
            Dim sumNewIT2G As Integer
            sumsumIT2G = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2G, "", iPlanNo)
            sumNewIT2G = sumsumIT2G
        End If
        '=====IT31=====
        If iPlanNo = "C" Or iPlanNo = "D" Then
            InsType = "IT31"
            If iPlanNo = "C" Then
                InsCoverage = "50000"
            ElseIf iPlanNo = "D" Then
                InsCoverage = "100000"
            End If
            InsMdeduct = "0"
            Dim sumsumIT31 As Integer
            Dim sumNewIT31 As Integer
            sumsumIT31 = InsertIQPOBJ_EXL(IISSU, InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT31, "", iPlanNo)
            sumNewIT31 = sumsumIT31

        End If
        ''=====IT2N=====
        'If iPlanNo = "D" Then
        '    InsType = "IT2N"
        '    InsCoverage = "3000"
        '    InsMdeduct = "0"
        '    InsertIQPOBJ_EXL(dr("IISSU"), InsType, InsDay, InsCoverage, InsMdeduct, CountryType, i, sumNewIT2N)
        'End If


    End Sub
    Public Class Has101000
        Inherits NEPage
        Property id As String
        Private Property _name As String
        Private Property _address As String
        Public ReadOnly Property name() As String
            Get
                Return _name
            End Get
        End Property
        Public ReadOnly Property address() As String
            Get
                Return _address
            End Get
        End Property

        Friend Sub New()
            _name = ""
            _address = ""
        End Sub

        Friend Sub New(ByVal id As String)
            Me.id = id
            GetClientInfo()
        End Sub

        Private Sub GetClientInfo()
            Dim dsParam As New beOpApiGcis
            Dim ds As New DataSet
            Dim dr As beOpApiGcis.beOpApiGcis001Row = CType(dsParam, beOpApiGcis).beOpApiGcis001.NewRow
            dr.Business_Accounting_NO = id '統一編號
            dsParam.beOpApiGcis001.AddbeOpApiGcis001Row(dr)
            ds = CType(New BHAS101002().CallBussinessService("PUBService.bsOpApiGcis001", dsParam), DataSet)
            If Not ds Is Nothing AndAlso ds.Tables.Count > 0 Then
                If ds.Tables(0).Rows.Count > 0 Then
                    '公司名稱
                    _name = ds.Tables(0).Rows(0).Item("Company_Name").ToString
                    '公司地址
                    _address = ds.Tables(0).Rows(0).Item("Company_Location").ToString
                End If
            End If
        End Sub
    End Class

    Protected Sub rad_flight_Y_CheckedChanged(sender As Object, e As EventArgs) Handles rad_flight_Y.CheckedChanged
        CheckSNNN.Text = ""
    End Sub

    Protected Sub rad_flight_N_CheckedChanged(sender As Object, e As EventArgs) Handles rad_flight_N.CheckedChanged
        CheckSNNN.Text = ""
    End Sub

    ''' <summary>
    ''' 要保人ID
    ''' </summary>
    Protected Sub txtIins_IINS2_TextChanged(sender As Object, e As EventArgs) Handles txtIins_IINS2.TextChanged
        Try
            '共用 關係人欄位只有1碼 
            CommonFunc.Relationship_add_Item(Page, Relationship, txtIins_IINS2)
        Catch ex As Exception

        End Try
    End Sub

    Shared Function GetString(ByVal str As Object) As String
        If IsDBNull(str) Then
            Return ""
        ElseIf IsNothing(str) Then
            Return ""
        ElseIf str.ToString.Trim = "" Then
            Return ""
        Else
            Return str.ToString.Trim
        End If
    End Function
    Protected Function K077GetToken(ByVal url As String, K077key As String, K077iv As String)
        Try
            'K077新系統 API:GetExternalSystemToken,用途:更新並獲得驗證Token
            Dim client = New RestClient(url)
            Dim request = New RestRequest(Method.POST)
            Dim requestData As New K077GetTokenRequest
            requestData.system = "CompApi"
            requestData.encodeString = B2BEntity.HasFunction.EncryptCFBK077("CompApi," & System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss zzz"), K077key, K077iv)
            request.AddHeader("Content-Type", "application/json")
            request.AddJsonBody(JsonConvert.SerializeObject(requestData))
            Dim response = client.Execute(request)
            Dim Obj As Linq.JObject = Newtonsoft.Json.JsonConvert.DeserializeObject(response.Content)
            Dim returnCode = Obj.Item("respHeader")("returnCode").ToString
            Dim encryptString = Obj.Item("transResp")("encryptString").ToString
            Dim message = Obj.Item("respHeader")("message").ToString

            If returnCode = "0000" Then
                Return encryptString
            Else
                Throw New Exception("returnCode=" & returnCode & ",message=" & message)
            End If
        Catch ex As Exception
            ErrMail("K077GetToken錯誤", ex.ToString)
            Me.JSRegisterFormTailScriptBlock("Init", "alert('" + ex.Message.ToString + "');window.location='/SKERP';")
        End Try
    End Function
    '   Public Function EncryptCFB(source As String) As String
    '	'K077新系統 加密用
    '	If String.IsNullOrWhiteSpace(source) Then
    '		Return String.Empty
    '	End If
    '	Try
    '		Dim key As Byte() = Convert.FromBase64String("Wm9HVjYrYU9DZ1BsTlRJc2xyLzZXVTVLb2d5bnJpQk8=")
    '		Dim iv As Byte() = Convert.FromBase64String("TnRJZ1ZjVzNNd2ZZeWdOSg==")
    '		If (key.Length <> 32) Then
    '			Throw New Exception("加密KEY長度異常 Length=" & key.Length)
    '		End If
    '		If (iv.Length <> 16) Then
    '			Throw New Exception("加密IV長度異常 Length=" & iv.Length)
    '		End If
    '		Using aes As Aes = Aes.Create()
    '			aes.Key = key
    '			aes.IV = iv
    '			aes.Mode = CipherMode.CFB
    '			aes.Padding = PaddingMode.PKCS7

    '			Using encryptor = aes.CreateEncryptor(aes.Key, aes.IV)
    '				Dim plainBytes As Byte() = Encoding.UTF8.GetBytes(source)
    '				Dim cipherBytes As Byte() = encryptor.TransformFinalBlock(plainBytes, 0, plainBytes.Length)
    '				Return Convert.ToBase64String(cipherBytes)
    '			End Using
    '		End Using
    '	Catch ex As Exception
    '		Throw New Exception("EncryptCFB error,Source:" + source)
    '	End Try
    'End Function
    'Public Function DecryptCFB(source As String) As String
    '	'K077新系統 解密用
    '	If String.IsNullOrWhiteSpace(source) Then
    '		Return String.Empty
    '	End If
    '	Try
    '		Dim key As Byte() = Convert.FromBase64String("Wm9HVjYrYU9DZ1BsTlRJc2xyLzZXVTVLb2d5bnJpQk8=")
    '		Dim iv As Byte() = Convert.FromBase64String("TnRJZ1ZjVzNNd2ZZeWdOSg==")
    '		If (key.Length <> 32) Then
    '			Throw New Exception("解密KEY長度異常 Length=" & key.Length)
    '		End If
    '		If (iv.Length <> 16) Then
    '			Throw New Exception("解密IV長度異常 Length=" & iv.Length)
    '		End If
    '		Using aes As Aes = Aes.Create()
    '			aes.Key = key
    '			aes.IV = iv
    '			aes.Mode = CipherMode.CFB
    '			aes.Padding = PaddingMode.None

    '			Using decryptor = aes.CreateDecryptor(aes.Key, aes.IV)
    '				Dim cipherBytes As Byte() = Convert.FromBase64String(source)
    '				Dim blockSize As Integer = aes.BlockSize / 8
    '				Dim paddingLength As Integer = blockSize - (cipherBytes.Length Mod blockSize)
    '				If (paddingLength < blockSize) Then
    '					Array.Resize(cipherBytes, cipherBytes.Length + paddingLength)
    '				End If
    '				Dim plainBytes As Byte() = decryptor.TransformFinalBlock(cipherBytes, 0, cipherBytes.Length)
    '				Return Encoding.UTF8.GetString(plainBytes).TrimEnd(vbNullChar)
    '			End Using
    '		End Using
    '	Catch ex As Exception
    '		Throw New Exception("DecryptCFB error,Source:" + source)
    '	End Try
    'End Function
#Region "Model"
    Public Class K077GetTokenRequest
        Public Property system As String
        Public Property encodeString As String
    End Class

#End Region


    Protected Sub ddlTargetProject_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTargetProject.SelectedIndexChanged
        '20250320 允許方案隨時變更；所以加人後[保費試算]再重算金額,顯示結果。
        Try
            '20250512 user提IQP & ITP方案選擇不使用下拉，改透過方案選擇按鈕設定，故將可選國家最大化，待﹝保費試算﹞時檢核該方案能否選擇該國家。
            Return

            '20250320 必定要重新選一次方案,IQP現在無；此狀況會依最後一次出現lblTargetProject.Text是什麼就依此重算
            'IQPOBJ細項做法全部刪掉 重新帶入
            lblTargetPlan.Text = ""
            txtCase.Text = ""
            Dim tr As HtmlTable = Form1.FindControl("divShow")
            If ddlTargetProject.SelectedValue = "ITP" Then
                tr.Attributes.Add("style", "")
            Else
                tr.Attributes.Add("style", "display:none")
            End If

            '20250411 方案變更後，旅遊地區清空，因為IQP & ITP可選擇的國家會有差異
            Dim locationLists As String() =
                New String() {
                    "ddlCountry_A1", "ddlCountry_A2", "ddlCountry_B1", "ddlCountry_B2", "ddlCountry_C1", "ddlCountry_C2", "ddlCountry_D1", "ddlCountry_D2",
                    "ddlCountry_E1", "ddlCountry_E2", "ddlCountry_F1", "ddlCountry_F2", "ddlCountry_G1", "ddlCountry_G2", "ddlCountry_H1", "ddlCountry_H2"
                }
            For Each dropdownItem As String In locationLists
                Dim tempItem As DropDownList = CType(Form1.FindControl(dropdownItem), DropDownList)
                tempItem.SelectedValue = ""
            Next

        Catch ex As Exception
            UIAppendError(ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' 重新整理一次 stateMain
    ''' </summary>
    Private Sub procChangedTargetProject()
        Try
            Dim ds As beITP101000 = CType(Me.StateMain, beITP101000)

            'F115031101 改依Statemain IQPISSU的IPlanNo當作目標方案，詳述參考GetTargetProject()的說明
            Dim TargetProject As String = GetTargetProject()
            'Dim TargetProject As String = ddlTargetProject.SelectedValue.Trim() '20250320
            Dim IPLANNO As String = ""

            '因修改被保人，方案改動僅能改主被保人，且主被保人改動、所有被保人方案會一同調整
            Dim IISSUPlanNo As String =
                StateMain.Tables("IQPISSU").AsEnumerable() _
                .Where(Function(r) Not IsDBNull(r("ISEQ")) AndAlso r("ISEQ").ToString() <> "") _
                .OrderBy(Function(r) Convert.ToInt32(r("ISEQ"))) _
                .Select(Function(r) r("IPLANNO")) _
                .FirstOrDefault()
            IPLANNO = IISSUPlanNo.Trim.ToUpper

            '若僅人身保險會是空的
            If IPLANNO = "A" Or IPLANNO = "C" Or IPLANNO = "D" Or IPLANNO = "J" Or IPLANNO = "K" Or IPLANNO = "L" Or IPLANNO = "M" Or IPLANNO = "N" Or IPLANNO = "O" Then
                TargetProject = "ITP"
            ElseIf String.IsNullOrEmpty(IPLANNO) Then
                TargetProject = "IQP"
            Else
                Throw New Exception($"被保人方案代碼{IPLANNO}不在處理範圍內")
            End If

            If TargetProject = "ITP" And String.IsNullOrEmpty(lblTargetPlan.Text.Trim) Then
                txtCase.Text = ""
                Throw New Exception("ITP必需選擇一個方案！！！")
            Else
                ' 20251008 移除團體費率，因得知在此之前早就無團體費率
                ' !!!!若需復原團體費率，bsGetItemLists的b2b_bhas101003CheckInstyp需留意要一併調整判斷!!!!
                If TargetProject = "ITP" Then
                    'F115031101 改IPLANNO以Statemain被保人資料為準
                    'IPLANNO = lblTargetPlan.Text.Trim.Substring(0, 1)

                    'B2C(ITP的MAIN_B_NRMK1將會是3)
                    'B2B ITP 的MAIN_B_NRMK1  則依照頁面費率別欄位進行變動，個人費率=2 團費率=3
                    ds.Tables("IQPMAIN").Rows(0)("MAIN_B_NRMK1") = "2" '集體投保否2=個人費率,3=團體費率

                    If (checkSpecialChannel().Item1 = True AndAlso (IPLANNO = "J" Or IPLANNO = "K" Or IPLANNO = "L")) Or
                        (checkChannelO().Item1 = True AndAlso (IPLANNO = "J" Or IPLANNO = "K" Or IPLANNO = "L" Or IPLANNO = "O")) Or
                        (checkChannelMN().Item1 = True AndAlso (IPLANNO = "M" Or IPLANNO = "N")) Then
                        ds.Tables("IQPMAIN").Rows(0)("MAIN_B_NRMK1") = "9"
                    End If
                Else
                    'B2C(ITP的MAIN_B_NRMK1將會是3)
                    'B2B IQP 的MAIN_B_NRMK1  則依照頁面費率別欄位進行變動，個人費率=0 團費率=1
                    ds.Tables("IQPMAIN").Rows(0)("MAIN_B_NRMK1") = "0" '集體投保否2=個人費率,3=團體費率
                End If
            End If

            Dim ORIOID As String = ds.Tables("IQPMAIN").Rows(0)("OID").ToString.Trim
            'Dim ORIIPLANNO As String = ds.Tables("IQPISSU").AsEnumerable().ToList().FirstOrDefault()("IPLANNO").ToString.Trim
            Dim ORITargetProject As String = ds.Tables("IQPMAIN").Rows(0)("IINSKIND").ToString.Trim
            'ITP 不便險取保費的方式 
            'IT41/IT4Y/IT42/IT62 這四檔 旅平的相關險種
            '1:InsType(險種)
            '2:InsDay(天數)
            '3.InsCoverage 保額 元
            '4.InsMdeduct 自負額 元
            '5.CountryType  代表費率別 ：3=個人費率、4=團費率
            'Dflag(要抓該費率檔內小於系統日的最大日)
            Dim IT414Y As Int32 = 0
            For Each row As DataRow In ds.Tables("IQPISSU").Rows
                ds.Tables("IQPOBJ").AsEnumerable().
                    Where(Function(x) x.Field(Of String)("IISSU") = row("IISSU")).Cast(Of DataRow).ToList().
                    ForEach(Sub(x)
                                If TargetProject = "IQP" Then
                                    If x.Field(Of String)("INSTYP") = "IT4Y" Then x.SetField(Of String)("INSTYP", "IQP0Y")
                                    If x.Field(Of String)("INSTYP") = "IT41" Then x.SetField(Of String)("INSTYP", "IQP01")
                                    If x.Field(Of String)("INSTYP") = "IT42" Then x.SetField(Of String)("INSTYP", "IQP02")
                                    If x.Field(Of String)("INSTYP") = "IT43" Then x.SetField(Of String)("INSTYP", "IQP03")
                                    If x.Field(Of String)("INSTYP") = "IT44" Then x.SetField(Of String)("INSTYP", "IQP04")
                                    If x.Field(Of String)("INSTYP") = "IT62" Then x.SetField(Of String)("INSTYP", "IQP11")
                                    If x.Field(Of String)("INSTYP") = "IT63" Then x.SetField(Of String)("INSTYP", "IQP12")
                                    If x.Field(Of String)("INSTYP") = "IT6E" Then x.SetField(Of String)("INSTYP", "IQP12")
                                    If x.Field(Of String)("INSTYP") = "IT64" Then x.SetField(Of String)("INSTYP", "IQP13")
                                    If x.Field(Of String)("INSTYP") = "IT65" Then x.SetField(Of String)("INSTYP", "IQP14")
                                    If x.Field(Of String)("INSTYP") = "IT66" Then x.SetField(Of String)("INSTYP", "IQP15")
                                End If
                                If TargetProject = "ITP" Then
                                    If x.Field(Of String)("INSTYP") = "IQP0Y" Then x.SetField(Of String)("INSTYP", "IT4Y")
                                    If x.Field(Of String)("INSTYP") = "IQP01" Then x.SetField(Of String)("INSTYP", "IT41")
                                    If x.Field(Of String)("INSTYP") = "IQP02" Then x.SetField(Of String)("INSTYP", "IT42")
                                    If x.Field(Of String)("INSTYP") = "IQP03" Then x.SetField(Of String)("INSTYP", "IT43")
                                    If x.Field(Of String)("INSTYP") = "IQP04" Then x.SetField(Of String)("INSTYP", "IT44")
                                    If x.Field(Of String)("INSTYP") = "IQP11" Then x.SetField(Of String)("INSTYP", "IT62")
                                    If x.Field(Of String)("INSTYP") = "IQP12" Then x.SetField(Of String)("INSTYP", "IT6E") '原IT63
                                    If x.Field(Of String)("INSTYP") = "IQP13" Then x.SetField(Of String)("INSTYP", "IT64")
                                    If x.Field(Of String)("INSTYP") = "IQP14" Then x.SetField(Of String)("INSTYP", "IT65")
                                    If x.Field(Of String)("INSTYP") = "IQP15" Then x.SetField(Of String)("INSTYP", "IT66")
                                End If
                            End Sub)
                'F114122301 20251218	
                '	海外標準型（J）IT41 or IT4Y 500萬以下，iProjectNo = B005 ；
                '	海外標準型（J）IT41 500萬以上，iProjectNo = B006；
                '	海外豪華型（K） iProjectNo = B007
                'F115031200
                '   海外基本型（C）B002 改成 B009
                '   海外超值型（D）B003 改成 B010
                '   海外標準型（J）IT41 or IT4Y 500萬以下 B005 改成 B012
                '   海外標準型（J）IT41 or IT4Y 500萬以上 B006 改成 B012A
                '   海外豪華型（K）B002 改成 B013
                'F115060500
                '   海外基本型（C）B009 改成 B014
                '   海外超值型（D）B010 改成 B015
                'F115082802
                '   海外小資型（O）B018

                If row("IPLANNO").ToString = "A" Then
                    row("IPROJECTNO") = "B001"
                ElseIf row("IPLANNO").ToString = "C" Then
                    row("IPROJECTNO") = "B014"
                ElseIf row("IPLANNO").ToString = "D" Then
                    row("IPROJECTNO") = "B015"
                ElseIf row("IPLANNO").ToString = "M" Then
                    row("IPROJECTNO") = "B016"
                ElseIf row("IPLANNO").ToString = "N" Then
                    row("IPROJECTNO") = "B017"
                ElseIf row("IPLANNO").ToString = "L" Then
                    row("IPROJECTNO") = "B008"
                ElseIf row("IPLANNO").ToString = "K" Then
                    row("IPROJECTNO") = "B013"
                ElseIf row("IPLANNO").ToString = "O" Then
                    row("IPROJECTNO") = "B018"
                ElseIf row("IPLANNO").ToString = "J" Then
                    IT414Y = (
                        From r In ds.Tables("IQPOBJ").AsEnumerable()
                        Where r.Field(Of String)("IISSU") = row("IISSU") AndAlso r.Field(Of String)("INSTYP") = "IT41"
                        Select r.Field(Of Int32)("INSAMT")
                    ).FirstOrDefault()

                    If String.IsNullOrEmpty(IT414Y) Or IT414Y = "0" Then
                        IT414Y = (
                            From r In ds.Tables("IQPOBJ").AsEnumerable()
                            Where r.Field(Of String)("IISSU") = row("IISSU") AndAlso r.Field(Of String)("INSTYP") = "IT4Y"
                            Select r.Field(Of Int32)("INSAMT")
                        ).FirstOrDefault()
                    End If

                    If IT414Y >= 0 AndAlso IT414Y < 5000000 Then
                        row("IPROJECTNO") = "B012"
                    ElseIf IT414Y >= 5000000 Then
                        row("IPROJECTNO") = "B012A"
                    End If

                End If


                'F114122301 20251218 要求更改成，同要保號，被保人ITP方案可各自不同。
                'row("IPLANNO") = If(TargetProject = "ITP", IPLANNO, DBNull.Value)
                ds.Tables("IQPOBJ").AcceptChanges()
            Next
            ds.Tables("IQPISSU").AcceptChanges()

            ' 保留將IT2J、IT2X、IT2K、IT29、IT2A、IT2L從IQPOBJ移除，因為雖然改代號，但user取出已暫存的舊資料仍會有舊代號
            ds.Tables("IQPOBJ").AsEnumerable().
                Where(Function(x)
                          Return Regex.IsMatch(x.Field(Of String)("INSTYP"), "IT11|IT2R|IT37|IT36|IT3A|IT3B|IT3C|IT3D|IT3E|IT39|IT38|IT2M|IT27|IT2D|IT28|IT31|IT32|IT33|IT34|IT35|IT2P|IT2U|IT2T|IT2S|IT2J|IT2X|IT2K|IT29|IT2A|IT2L")
                      End Function).ToList().ForEach(Sub(x) x.Delete())
            ds.Tables("IQPOBJ").AcceptChanges()
            If IPLANNO = "A" Then
                ds.Tables("IQPOBJ").AsEnumerable().
                    Where(Function(x)
                              Return Regex.IsMatch(x.Field(Of String)("INSTYP"), "IT2G")
                          End Function).ToList().ForEach(Sub(x) x.Delete())
                ds.Tables("IQPOBJ").AcceptChanges()
            End If
            If TargetProject = "ITP" AndAlso IPLANNO <> "L" Then
                '險種代號       C 海外基本型	D 海外超值型
                'IT11		    限額200000元	限額200000元
                'IT2R		    限額100萬		限額300萬
                'IT37(原IT2K)   限額10000元		限額30000元
                'IT36(原IT2J)   限額20000元		限額60000元 
                'IT3A(原IT2A)   每次4000元		每次6000元
                'IT3B(原IT2L)   每次3000元		每次3000元
                'IT3C           每次10000元		每次30000元
                'IT3D           每次1000元		每次2000元
                'IT3E           每次1000元		每次2000元
                'IT39(原IT29)   每次4000元		每次6000元
                'IT38(原IT2X)	每次4000元		每次6000元
                'IT2M		    每次3000元		每次6000元
                'IT27		    每日3000元		每日5000元
                'IT2D		    限額20000元		限額30000元
                'IT28		    每次3000元		每次5000元
                'IT31		    限額50000元		限額50000元

                '險種代號       J 海外標準型	K 海外豪華型
                'IT11		    限額3000000元	限額5000000元
                'IT2R		    限額250萬		限額250萬
                'IT3B(原IT2L)   每次3000元		每次3000元
                'IT3A(原IT2A)   每次4000元		每次6000元
                'IT39(原IT29)   每次4000元		每次6000元
                'IT38(原IT2X)   每次4000元		每次6000元
                'IT2M		    每次3000元		每次5000元
                'IT27		    每日3000元		每日6000元
                'IT2D		    限額30000元		限額50000元
                'IT28		    每次3000元		每次3000元
                'IT31		    限額30000元		限額50000元
                'IT36(原IT32)   限額120000元	限額150000元
                'IT37(原IT33)   限額120000元	限額150000元
                'IT34		    限額120000元	限額150000元
                'IT35		    限額120000元	限額150000元
                'IT2P		    限額3000元		限額5000元
                'IT2T		    每日10000元		每日10000元
                'IT2U		    每次3000元		每次5000元
                'IT2S		    每次3000元		每次5000元

                '險種代號       O 海外小資型	
                'IT11		    限額200000元		
                'IT36(原IT2J)   限額60000元		
                'IT37(原IT2K)   限額30000元
                'IT38(原IT2X)	每次5000元
                'IT39(原IT29)   每次5000元
                'IT3A(原IT2A)   每次5000元		
                'IT3B(原IT2L)   每次3000元	
                'IT2M		    每次5000元		
                'IT28		    每次3000元	
                'IT31		    限額50000元	
                'IT2D		    限額50000元	
                'IT3D           每次1000元	

                GetDay()
                Dim xday As String = txtIQPITP_Days.Text.Trim
                Dim NRMK5 As String = ds.Tables("IQPMAIN").Rows(0)("NRMK5").ToString.Trim
                Dim countryType As String = getCountryType(NRMK5, ds.Tables("IQPISSU").Rows.Count)
                Dim IT2RFree As Boolean = False
                Dim IT41 As Int32 = 0
                For Each row As DataRow In ds.Tables("IQPISSU").Rows
                    '20251218
                    '	身故500萬以上，方案J K優惠IT2R不算錢
                    If row("IPLANNO").ToString = "J" Or row("IPLANNO").ToString = "K" Then
                        IT2RFree = False
                        IT41 = (
                            From r In ds.Tables("IQPOBJ").AsEnumerable()
                            Where r.Field(Of String)("IISSU") = row("IISSU") AndAlso r.Field(Of String)("INSTYP") = "IT41"
                            Select r.Field(Of Int32)("INSAMT")
                        ).FirstOrDefault()

                        If IT41 >= 5000000 Then IT2RFree = True
                    End If

                    Dim IISSU As String = row("IISSU").ToString
                    'If ORITargetProject.ToUpper = "IQP" Then
                    '    IPLANNO = lblTargetPlan.Text.Trim.Substring(0, 1)
                    'Else
                    '    IPLANNO = row("IPLANNO").ToString
                    'End If

                    Select Case IPLANNO
                        Case "A"
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT11", xday, "50000", "2500", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT28", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2G", xday, "3000", "", countryType)
                        Case "C", "M"
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT11", xday, "200000", "2500", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT36", xday, "20000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT38", xday, "4000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT37", xday, "10000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT39", xday, "4000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3A", xday, "4000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3B", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3C", xday, "10000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3D", xday, "1000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3E", xday, "1000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2M", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT27", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT28", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2D", xday, "20000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2R", xday, "1000000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT31", xday, "50000", "", countryType)
                        Case "D", "N"
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT11", xday, "200000", "2500", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT36", xday, "60000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT38", xday, "6000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT37", xday, "30000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT39", xday, "6000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3A", xday, "6000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3B", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3C", xday, "30000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3D", xday, "2000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3E", xday, "2000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2M", xday, "6000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT27", xday, "5000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT28", xday, "5000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2D", xday, "30000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2R", xday, "3000000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT31", xday, "50000", "", countryType)
                        Case "J"
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT11", xday, "3000000", "2500", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT38", xday, "4000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT39", xday, "4000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3A", xday, "4000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3B", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2M", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT27", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT28", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2D", xday, "30000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT31", xday, "30000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT36", xday, "120000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT37", xday, "120000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT34", xday, "120000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT35", xday, "120000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2P", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2U", xday, "10000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2T", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2S", xday, "3000", "", countryType)

                            If IT2RFree Then
                                '2026-01-12 要求不寫入IT2R
                                'addIQPObjFreePrem(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2R", xday, "2500000", "", countryType)
                            Else
                                addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2R", xday, "2500000", "", countryType)
                            End If
                        Case "K"
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT11", xday, "5000000", "2500", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT38", xday, "6000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT39", xday, "6000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3A", xday, "6000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3B", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2M", xday, "5000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT27", xday, "6000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT28", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2D", xday, "50000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT31", xday, "50000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT36", xday, "150000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT37", xday, "150000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT34", xday, "150000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT35", xday, "150000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2P", xday, "5000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2U", xday, "10000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2T", xday, "5000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2S", xday, "5000", "", countryType)

                            If IT2RFree Then
                                '2026-01-12 要求不寫入IT2R
                                'addIQPObjFreePrem(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2R", xday, "2500000", "", countryType)
                            Else
                                addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2R", xday, "2500000", "", countryType)
                            End If
                        Case "O"
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT11", xday, "200000", "2500", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT36", xday, "60000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT38", xday, "5000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT37", xday, "30000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT39", xday, "5000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3A", xday, "5000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3B", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2M", xday, "5000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT28", xday, "3000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT31", xday, "50000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT2D", xday, "50000", "", countryType)
                            addIQPObj(ds.Tables("IQPOBJ"), ORIOID, IISSU, "IT3D", xday, "1000", "", countryType)
                        Case Else
                            Throw New Exception($"{IPLANNO}方案不在處理範圍內！！！")
                    End Select
                    ds.Tables("IQPOBJ").AcceptChanges()
                Next
                BaseSetChildFormData(Me.dgdMain2, Me.StateMain)
            End If
            ds.Tables("IQPMAIN").Rows(0)("IINSKIND") = TargetProject
            ds.Tables("IQPMAIN").AcceptChanges()
            Dim dsTemp As New DataSet

            If ORITargetProject <> TargetProject Or ORIOID.Substring(0, 2).ToUpper <> TargetProject.Substring(0, 2).ToUpper Then
                dsTemp = Me.InvokeBusinessService("B2BService.bsBHAS101003SaveTemp", StateMain)
            Else
                dsTemp = Me.InvokeBusinessService("B2BService.bsBHAS101003SaveEdit", StateMain)
            End If

            'If ORITargetProject = TargetProject Then
            '    dsTemp = Me.InvokeBusinessService("B2BService.bsBHAS101003SaveEdit", StateMain)
            'Else
            '    dsTemp = Me.InvokeBusinessService("B2BService.bsBHAS101003SaveTemp", StateMain)
            'End If
            If dsTemp IsNot Nothing AndAlso dsTemp.Tables("IQPMAIN") IsNot Nothing AndAlso dsTemp.Tables("IQPMAIN").Rows.Count > 0 Then
                txtIQPMainOid.Text = dsTemp.Tables("IQPMAIN").Rows(0)("OID").ToString
                StateMain = dsTemp
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Function getCountryType(NRMK5 As String, Total As Integer) As String
        Dim CountryType As String = ""
        If ConfigurationManager.AppSettings("bCHECK_5") = "1" Then
            If CHK5.Checked = True Then
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "5" '國外5 外6 別
                Else
                    CountryType = "6" '國外5 外6 別
                End If
            Else
                If TRAVEL_A.SelectedValue = "400" Then
                    CountryType = "1" '國內1 外2 別
                Else
                    CountryType = "2" '國內1 外2 別
                End If
            End If
        Else
            If NRMK5 = "400" Then
                CountryType = "1" '國內1 外2 別
            Else
                CountryType = "2" '國內1 外2 別
            End If
        End If
        Return CountryType
    End Function


    Private Sub addIQPObj(dt As beITP101000.IQPOBJDataTable, OID As String, ID As String,
                          INSTYPE As String, InsDay As String, InsCoverage As String,
                          InsMdeduct As String, CountryType As String)
        Try
            'ITP 不便險取保費的方式
            '1:InsType(險種)
            '2:InsDay(天數)
            '3.InsCoverage 保額 元
            '4.InsMdeduct 自負額 元
            '5.CountryType 1=國內 2=國外 國外5 外6 別
            Dim row As beITP101000.IQPOBJRow = dt.NewIQPOBJRow()
            row("IQPMAINOID") = OID
            row("IISSU") = ID
            row("INSTyp") = INSTYPE
            row("INSAMT") = InsCoverage
            row("MDEDUCT") = If(String.IsNullOrEmpty(InsMdeduct), DBNull.Value, InsMdeduct)
            row("INSPREM") = GetITPPrem(INSTYPE, InsDay, If(String.IsNullOrEmpty(InsCoverage), 0, InsCoverage), If(String.IsNullOrEmpty(InsMdeduct), 0, InsMdeduct), CountryType)
            dt.AddIQPOBJRow(row)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub addIQPObjFreePrem(dt As beITP101000.IQPOBJDataTable, OID As String, ID As String,
                          INSTYPE As String, InsDay As String, InsCoverage As String,
                          InsMdeduct As String, CountryType As String)
        Try
            'ITP 不便險取保費的方式
            '1:InsType(險種)
            '2:InsDay(天數)
            '3.InsCoverage 保額 元
            '4.InsMdeduct 自負額 元
            '5.CountryType 1=國內 2=國外 國外5 外6 別
            Dim row As beITP101000.IQPOBJRow = dt.NewIQPOBJRow()
            row("IQPMAINOID") = OID
            row("IISSU") = ID
            row("INSTyp") = INSTYPE
            row("INSAMT") = InsCoverage
            row("MDEDUCT") = If(String.IsNullOrEmpty(InsMdeduct), DBNull.Value, InsMdeduct)
            row("INSPREM") = "0"
            dt.AddIQPOBJRow(row)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary> 取回 ITP 項目金額 </summary>
    Private Function GetITPPrem(INSType As String, InsDay As Integer, InsCoverage As Integer,
                                InsMdeduct As Integer, CountryType As Integer) As Integer 'INSType 4碼險種 EX.IT11, InsCoverage 保額
        'ITP 不便險取保費的方式
        '1:InsType(險種)
        '2:InsDay(天數)
        '3.InsCoverage 保額 元
        '4.InsMdeduct 自負額 元
        '5.CountryType 3=國內 4=國外
        'Dflag(要抓該費率檔內小於系統日的最大日)
        Dim ds As DataSet = InvokeCommonService("B2BService.bsITP101000GetOracleValue", INSType, InsDay, InsCoverage, InsMdeduct, CountryType)
        If ds.Tables("ITPPrem").Rows.Count > 0 Then
            Return CInt(ds.Tables("ITPPrem").Rows(0)("PREM").ToString.Trim) '寫入試算金額
        Else
            Return 0
        End If

    End Function

    Private Function ContainFullWidthChar(input As String) As Boolean
        For Each c As Char In input
            Dim code As Integer = AscW(c)

            If code >= &HFF01 AndAlso code <= &HFF5E Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function checkTOTPREM() As String
        Dim result As String = ""
        Dim SumTotal As Integer = 0
        Dim TOTPREM As Integer = 0

        TOTPREM = CType((StateMain.Tables("IQPMAIN").Rows(0)("TOTPREM").ToString()), Integer)

        For Each drE As DataRow In StateMain.Tables("IQPOBJ").Rows
            If drE.RowState <> DataRowState.Deleted Then
                SumTotal += CType(drE("INSPREM"), Integer)
            End If
        Next

        If TOTPREM <> SumTotal Then
            result = txtIQPMainOid.Text & " 總保額加總不符合。IQPMAIN.TOTPREM：" & TOTPREM.ToString() & " ； IQPOBJ.Total_INSPREM：" & SumTotal.ToString()
        End If

        Return result
    End Function

    Private Function checkISSUPlanNo() As String
        Dim result As String = ""
        Dim chkResult As String = ""
        Dim firstRound As Boolean = True
        For Each row As DataRow In StateMain.Tables("IQPISSU").Rows
            If row.RowState <> DataRowState.Deleted Then
            End If
        Next

        Return result
    End Function

    Private Function checkProjectNo() As String

        Dim IT414Y As Int32 = 0
        Dim result As String = ""
        Dim chkResult As Boolean = True
        Dim flagPlanNo As String = ""
        Dim firstRound As Boolean = True
        For Each row As DataRow In StateMain.Tables("IQPISSU").Rows
            If row.RowState <> DataRowState.Deleted Then

#Region "檢查被保險人方案是否一致"
                If firstRound Then
                    flagPlanNo = row("IPLANNO").ToString
                Else
                    If flagPlanNo <> row("IPLANNO").ToString Then
                        result = txtIQPMainOid.Text & " 被保險人方案不一致。"
                        Exit For
                    End If
                End If

                firstRound = False
#End Region


#Region "檢查被保險人IPLANNO & IPROJECTNO"

                'F114122301 20251218	
                '	海外標準型（J）IT41 or IT4Y 500萬以下，iProjectNo = B005 ；
                '	海外標準型（J）IT41 500萬以上，iProjectNo = B006；
                '	海外豪華型（K） iProjectNo = B007
                'F115031200
                '   海外基本型（C）B002 改成 B009
                '   海外超值型（D）B003 改成 B010
                '   海外標準型（J）IT41 or IT4Y 500萬以下 B005 改成 B012
                '   海外標準型（J）IT41 or IT4Y 500萬以上 B006 改成 B012A
                '   海外豪華型（K）B002 改成 B013
                'F115060500
                '   海外基本型（C）B009 改成 B014
                '   海外超值型（D）B010 改成 B015
                'F115082802
                '   海外小資型（O）B018

                If String.IsNullOrEmpty(row("IPLANNO").ToString) And String.IsNullOrEmpty(row("IPROJECTNO").ToString) = False Then
                    chkResult = False
                ElseIf row("IPLANNO").ToString = "A" And row("IPROJECTNO").ToString <> "B001" Then
                    chkResult = False
                ElseIf row("IPLANNO").ToString = "C" And row("IPROJECTNO").ToString <> "B014" Then
                    chkResult = False
                ElseIf row("IPLANNO").ToString = "D" And row("IPROJECTNO").ToString <> "B015" Then
                    chkResult = False
                ElseIf row("IPLANNO").ToString = "M" And row("IPROJECTNO").ToString <> "B016" Then
                    chkResult = False
                ElseIf row("IPLANNO").ToString = "N" And row("IPROJECTNO").ToString <> "B017" Then
                    chkResult = False
                ElseIf row("IPLANNO").ToString = "L" And row("IPROJECTNO").ToString <> "B008" Then
                    chkResult = False
                ElseIf row("IPLANNO").ToString = "K" And row("IPROJECTNO").ToString <> "B013" Then
                    chkResult = False
                ElseIf row("IPLANNO").ToString = "O" And row("IPROJECTNO").ToString <> "B018" Then
                    chkResult = False
                ElseIf row("IPLANNO").ToString = "J" Then
                    IT414Y = (
                                 From r In StateMain.Tables("IQPOBJ").AsEnumerable()
                                 Where r.Field(Of String)("IISSU") = row("IISSU") AndAlso r.Field(Of String)("INSTYP") = "IT41"
                                 Select r.Field(Of Int32)("INSAMT")
                             ).FirstOrDefault()

                    If String.IsNullOrEmpty(IT414Y) Or IT414Y = "0" Then
                        IT414Y = (
                                     From r In StateMain.Tables("IQPOBJ").AsEnumerable()
                                     Where r.Field(Of String)("IISSU") = row("IISSU") AndAlso r.Field(Of String)("INSTYP") = "IT4Y"
                                     Select r.Field(Of Int32)("INSAMT")
                                 ).FirstOrDefault()
                    End If

                    If (IT414Y >= 0 AndAlso IT414Y < 5000000) And row("IPROJECTNO") <> "B012" Then
                        chkResult = False
                    ElseIf IT414Y >= 5000000 And row("IPROJECTNO") <> "B012A" Then
                        chkResult = False
                    End If
                End If

                If chkResult = False Then
                    result = txtIQPMainOid.Text & " 方案代碼不符合。" & "IPlanNo：" & row("IPLANNO").ToString & " ； IProjectNo：" & row("IPROJECTNO").ToString & " 。"
                    Exit For
                End If
#End Region

            End If
        Next

        Return result
    End Function

    Private Function checkOIDNIissu() As String

        Dim result As String = ""
        Dim IISSUP1anNo As String = getIISSUPlanNo()

        '若僅人身保險會是空的
        If IISSUP1anNo = "A" Or IISSUP1anNo = "C" Or IISSUP1anNo = "D" Or IISSUP1anNo = "J" Or IISSUP1anNo = "K" Or IISSUP1anNo = "L" Or IISSUP1anNo = "M" Or IISSUP1anNo = "N" Or IISSUP1anNo = "O" Then
            If (txtIQPMainOid.Text).ToUpper.StartsWith("IT") = False Then
                result = txtIQPMainOid.Text & " 要保號與被保人方案不符。IQPIISSU.IPLANNO：" & IISSUP1anNo
            End If
        ElseIf String.IsNullOrEmpty(IISSUP1anNo) Then
            If (txtIQPMainOid.Text).ToUpper.StartsWith("IQ") = False Then
                result = txtIQPMainOid.Text & " 要保號與被保人方案不符。IQPIISSU.IPLANNO：" & IISSUP1anNo
            End If
        Else
            Throw New Exception($"被保人方案代碼{IISSUP1anNo}不在處理範圍內")
        End If

        Return result
    End Function

    Private Function checkMAIN_B_NRMK() As String
        Dim result As String = ""
        Dim B_NRMK As String = StateMain.Tables("IQPMAIN").Rows(0)("MAIN_B_NRMK1")
        Dim IISSUP1anNo As String = getIISSUPlanNo()

        If (txtIQPMainOid.Text).ToUpper.StartsWith("IT") Then
            If (checkSpecialChannel().Item1 = True AndAlso (IISSUP1anNo = "J" Or IISSUP1anNo = "K" Or IISSUP1anNo = "L")) Or
               (checkChannelO().Item1 = True AndAlso (IISSUP1anNo = "J" Or IISSUP1anNo = "K" Or IISSUP1anNo = "L" Or IISSUP1anNo = "O")) Or
               (checkChannelMN().Item1 = True AndAlso (IISSUP1anNo = "N" Or IISSUP1anNo = "M")) Then
                If B_NRMK <> "9" Then
                    result = txtIQPMainOid.Text & " 方案與費率不符。IQPIISSU.IPLANNO：" & IISSUP1anNo & "； MAIN_B_NRMK1：" & B_NRMK
                End If
            Else
                If B_NRMK <> "2" Then
                    result = txtIQPMainOid.Text & " 方案與費率不符。IQPIISSU.IPLANNO：" & IISSUP1anNo & "； MAIN_B_NRMK1：" & B_NRMK
                End If
            End If
        Else
            If B_NRMK <> "0" Then
                result = txtIQPMainOid.Text & " 方案與費率不符。IQPIISSU.IPLANNO：" & IISSUP1anNo & "； MAIN_B_NRMK1：" & B_NRMK
            End If
        End If

        Return result
    End Function

    Private Function checkNINSREL() As String
        ' 檢查 DataSet 與 DataTable 是否存在
        If StateMain Is Nothing OrElse Not StateMain.Tables.Contains("IQPISSU") OrElse Not StateMain.Tables.Contains("IQPMAIN") Then
            Return ""
        End If

        Dim dtIssu As DataTable = StateMain.Tables("IQPISSU")
        Dim dtMain As DataTable = StateMain.Tables("IQPMAIN")

        ' 取得 被保人關係為要保本人(NINSREL = "1") 做比對
        Dim issuRows As DataRow() = dtIssu.Select("NINSREL = '1'")
        If issuRows.Length = 0 Then
            Return ""
        End If

        Dim drIssu As DataRow = issuRows(0)
        Dim drMain As DataRow = dtMain.Rows(0)

        ' 取出字串欄位並處理 DBNull
        Dim nissu As String = If(drIssu.IsNull("NISSU"), String.Empty, drIssu("NISSU").ToString())
        Dim iissu As String = If(drIssu.IsNull("IISSU"), String.Empty, drIssu("IISSU").ToString())
        Dim nins As String = If(drMain.IsNull("NINS"), String.Empty, drMain("NINS").ToString())
        Dim iins As String = If(drMain.IsNull("IINS"), String.Empty, drMain("IINS").ToString())

        ' 取出日期欄位並轉換為 DateTime? (Nullable(Of DateTime)) 以確保比對準確
        Dim nissubd As DateTime? = If(drIssu.IsNull("NISSUBD"), CType(Nothing, DateTime?), Convert.ToDateTime(drIssu("NISSUBD")))
        Dim insbd As DateTime? = If(drMain.IsNull("INSBD"), CType(Nothing, DateTime?), Convert.ToDateTime(drMain("INSBD")))

        ' 判斷是否不符
        If nissu <> nins OrElse nissu <> NAME_IINS.Text.Trim Then
            Return "要被保人關係為本人，姓名、身分證、生日需相同"
        End If
        If iissu <> iins OrElse iissu <> txtIins_IINS2.Text.Trim Then
            Return "要被保人關係為本人，姓名、身分證、生日需相同"
        End If

        If nissubd <> insbd OrElse nissubd <> Birtdy_IINS.GetDate Then
            Return "要被保人關係為本人，姓名、身分證、生日需相同"
        End If

        ' 如果完全符合，則依邏輯預設回傳空字串
        Return ""

    End Function

    Private Sub checkFlightOrNot()

        If StateMain.Tables.Contains("HASM_FLIGHT_DETAIL") Then
            For Each dr4 As DataRow In StateMain.Tables("HASM_FLIGHT_DETAIL").Rows
                If dr4.RowState <> DataRowState.Deleted Then
                    StateMain.Tables("IQPMAIN").Rows(0)("IFLIGHT") = "Y"
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Function checkIRECPESN() As String
        Dim Result As String = ""
        Dim DATA As String = ""
        Dim IAGENT As String = ""
        Dim IChannelList As List(Of String)
        Dim IChannel As String = Me.txtCompanyid.Text

        If IChannel = "SKI" Then
            Return ""
        End If

        Try
            Dim ChannelTuple As Tuple(Of Boolean, String) = GetChannelList()
            If ChannelTuple.Item1 = False Then
                Result = ChannelTuple.Item2
            Else
                DATA = ChannelTuple.Item2

                IChannelList = DATA.Split(",").ToList()

                If IChannelList.Contains(IChannel) Then
                    If String.IsNullOrWhiteSpace(txtIRECPESN.Text) Then
                        Result = "通路人員代號（登錄證字號） 不可為空白."
                    Else
                        '檢查 通路人員代號（登錄證字號） 格式是否正確
                        Dim checkResult As String = HasFunction.Check_MONEY_PAPERS(txtIRECPESN.Text, "855", New Date())
                        Result = checkResult
                    End If

                    If String.IsNullOrWhiteSpace(txtNrecpesn.Text) Then
                        Result = If(String.IsNullOrEmpty(Result), "", Result + "   ") + "通路人員名稱 不可為空白."
                    End If
                End If
            End If
        Catch ex As Exception
            Result = "通路代碼清單取得失敗..."
        End Try

        Return Result
    End Function

    Private Function checkSpecialChannel() As Tuple(Of Boolean, String)
        Dim IsSpecialChannel As Boolean = False
        Dim DATA As String = ""
        Dim IChannelList As List(Of String)
        Dim IChannel As String = Me.txtCompanyid.Text

        Try
            If String.IsNullOrEmpty(txtIsSPEChannel.Text.Trim) Then
                Dim ChannelTuple As Tuple(Of Boolean, String) = GetSpecialChannelList()
                If ChannelTuple.Item1 = False Then
                    Return Tuple.Create(Of Boolean, String)(False, ChannelTuple.Item2)
                Else
                    DATA = ChannelTuple.Item2

                    IChannelList = DATA.Split(",").ToList()

                    If IChannelList.Contains(IChannel) Then
                        IsSpecialChannel = True
                        txtIsSPEChannel.Text = "TRUE"
                    Else
                        IsSpecialChannel = False
                        txtIsSPEChannel.Text = "FALSE"
                    End If
                End If
            ElseIf txtIsSPEChannel.Text.ToUpper = "TRUE" Then
                IsSpecialChannel = True
            Else
                IsSpecialChannel = False
            End If
        Catch ex As Exception
            Return Tuple.Create(Of Boolean, String)(False, "通路代碼清單取得失敗...")
        End Try

        Return Tuple.Create(Of Boolean, String)(IsSpecialChannel, "")
    End Function

    ''' <summary>
    ''' 取得哪些特定通路可以指定要保的通路代碼&帳號
    ''' </summary>
    ''' <returns></returns>
    Private Function checkAssignACC() As Tuple(Of Boolean, String)
        Dim IsSepcialChannel As Boolean = False
        Dim DATA As String = ""
        Dim IChannelList As List(Of String)
        Dim IChannel As String = Me.ScsOrgID

        Try
            Dim ChannelTuple As Tuple(Of Boolean, String) = GetAssignACCList()
            If ChannelTuple.Item1 = False Then
                Return Tuple.Create(Of Boolean, String)(False, ChannelTuple.Item2)
            Else
                DATA = ChannelTuple.Item2

                IChannelList = DATA.Split(",").ToList()

                If IChannelList.Contains(IChannel) Then
                    IsSepcialChannel = True
                Else
                    IsSepcialChannel = False
                End If
            End If
        Catch ex As Exception
            Return Tuple.Create(Of Boolean, String)(False, "通路代碼清單取得失敗...")
        End Try

        Return Tuple.Create(Of Boolean, String)(IsSepcialChannel, "")
    End Function

    ''' <summary>
    ''' 取得哪些指定通路可以要安心型&特選型的通路代碼&帳號
    ''' </summary>
    ''' <returns></returns>
    Private Function checkChannelMN() As Tuple(Of Boolean, String)
        Dim IsSpecialChannel As Boolean = False
        Dim DATA As String = ""
        Dim IChannelList As List(Of String)
        Dim IChannel As String = Me.txtCompanyid.Text

        Try
            Dim ChannelTuple As Tuple(Of Boolean, String) = GetChannelMNList()
            If ChannelTuple.Item1 = False Then
                Return Tuple.Create(Of Boolean, String)(False, ChannelTuple.Item2)
            Else
                DATA = ChannelTuple.Item2

                IChannelList = DATA.Split(",").ToList()

                If IChannelList.Contains(IChannel) Then
                    IsSpecialChannel = True
                Else
                    IsSpecialChannel = False
                End If
            End If
        Catch ex As Exception
            Return Tuple.Create(Of Boolean, String)(False, "MN方案用的通路代碼清單取得失敗...")
        End Try

        Return Tuple.Create(Of Boolean, String)(IsSpecialChannel, "")
    End Function

    ''' <summary>
    ''' 取得哪些指定通路(目前:大陸保代)可以要小資型的通路代碼&帳號
    ''' </summary>
    ''' <returns></returns>
    Private Function checkChannelO() As Tuple(Of Boolean, String)
        Dim IsSpecialChannel As Boolean = False
        Dim DATA As String = ""
        Dim IChannelList As List(Of String)
        Dim IChannel As String = Me.txtCompanyid.Text

        Try
            Dim ChannelTuple As Tuple(Of Boolean, String) = GetChannelOList()
            If ChannelTuple.Item1 = False Then
                Return Tuple.Create(Of Boolean, String)(False, ChannelTuple.Item2)
            Else
                DATA = ChannelTuple.Item2

                IChannelList = DATA.Split(",").ToList()

                If IChannelList.Contains(IChannel) Then
                    IsSpecialChannel = True
                Else
                    IsSpecialChannel = False
                End If
            End If
        Catch ex As Exception
            Return Tuple.Create(Of Boolean, String)(False, "MN方案用的通路代碼清單取得失敗...")
        End Try

        Return Tuple.Create(Of Boolean, String)(IsSpecialChannel, "")
    End Function

    Private Function GetChannelFromCommonCode(ByVal code As String) As Tuple(Of Boolean, String)
        Dim DATA1 As String = "", DATA2 As String, DATA3 As String

        Try
            Dim dsIChannel As DataSet = InvokeCommonService("PUBService.bsSingleSQL", New String() {"CMNP_COMMONCODE", "TYPE ='BHAS101003' AND CODE ='" & code & "'"})

            If Not IsNothing(dsIChannel) AndAlso dsIChannel.Tables.Count > 0 AndAlso dsIChannel.Tables(0).Rows.Count > 0 Then
                DATA1 = dsIChannel.Tables(0).Rows(0)("DATA1").ToString.Trim
                DATA2 = dsIChannel.Tables(0).Rows(0)("DATA2").ToString.Trim
                DATA3 = dsIChannel.Tables(0).Rows(0)("DATA3").ToString.Trim
                DATA1 = DATA1 + If(String.IsNullOrEmpty(DATA2), "", "," + DATA2) + If(String.IsNullOrEmpty(DATA3), "", "," + DATA3)

                If String.IsNullOrEmpty(DATA1) Then
                    Return Tuple.Create(Of Boolean, String)(False, "通路代碼清單取得失敗...")
                End If
            End If
        Catch ex As Exception
            Return Tuple.Create(Of Boolean, String)(False, "通路代碼清單取得失敗...")
        End Try

        Return Tuple.Create(Of Boolean, String)(True, DATA1)
    End Function

    ''' <summary>
    ''' 取特殊通路代碼清單。存在之代碼表示可以投保指定方案【海外標準型】&【海外豪華型】
    ''' </summary>
    ''' <returns></returns>
    Private Function GetSpecialChannelList() As Tuple(Of Boolean, String)
        Return GetChannelFromCommonCode("SPECIAL_CHANNEL")
    End Function

    ''' <summary>
    ''' 取指定通路代碼清單。存在之代碼表示可以投保指定方案【海外安心型】&【海外特選型】
    ''' </summary>
    ''' <returns></returns>
    Private Function GetChannelMNList() As Tuple(Of Boolean, String)
        Return GetChannelFromCommonCode("CHANNEL_MN")
    End Function

    ''' <summary>
    ''' 取指定通路代碼清單。存在之代碼表示可以投保指定方案【海外小資型】
    ''' </summary>
    ''' <returns></returns>
    Private Function GetChannelOList() As Tuple(Of Boolean, String)
        Return GetChannelFromCommonCode("CHANNEL_O")
    End Function

    ''' <summary>
    ''' 取通路代碼清單。存在之代碼表示需要求user輸入﹝登錄證字號﹞，且預設將﹝通路內控資料﹞一欄開啟
    ''' </summary>
    ''' <returns>當tuple.item1 = true，tuple.item2為﹝通路代碼清單﹞字串；當tuple.item1 = false，tuple.item2為﹝錯誤訊息﹞</returns>
    Private Function GetChannelList() As Tuple(Of Boolean, String)
        Return GetChannelFromCommonCode("ICHANNEL_CHECK")
    End Function

    ''' <summary>
    ''' 取通路代碼清單。存在之代碼表示可以指定要保的通路代碼與帳號
    ''' </summary>
    ''' <returns>當tuple.item1 = true，tuple.item2為﹝通路代碼清單﹞字串；當tuple.item1 = false，tuple.item2為﹝錯誤訊息﹞</returns>
    Private Function GetAssignACCList() As Tuple(Of Boolean, String)
        Return GetChannelFromCommonCode("ASSIGN_ACCOUNT")
    End Function

    Private Function ChannelColsShowOrNot() As Boolean
        Dim Result As Boolean = False
        Dim Data As String = ""
        Dim IChannelList As List(Of String)
        Dim IChannel As String = Me.txtCompanyid.Text

        Try
            Dim ChannelTuple As Tuple(Of Boolean, String) = GetChannelList()

            If ChannelTuple.Item1 = True Then
                Data = ChannelTuple.Item2

                IChannelList = Data.Split(",").ToList()
            End If

            If Not IChannelList Is Nothing AndAlso IChannelList.Count > 0 AndAlso IChannelList.Contains(IChannel) Then
                Result = True

            ElseIf IChannel = "K077" Or IChannel = "KA77" Or IChannel = "KKEW" Or IChannel = "KEJ0" Or IChannel = "KKFM" Or IChannel = "KKFP" Or IChannel = "KKEX" Then
                '抓指定通路代號的權限 通路檢核
                Result = True
            End If

        Catch ex As Exception
            Result = False
        End Try

        Return Result
    End Function

    ''' <summary>
    ''' 複製同被保人與要保人關係的下拉選單Relationship, 而selectedValue則依照StateMain的IQPISSU資料表中【與要保人關係(NINSREL)】是否有"其他"給值
    ''' </summary>
    ''' <returns></returns>
    Private Function GetTempRelationshipDDL() As DropDownList

        Dim tempDDL As New DropDownList()
        For Each item As ListItem In Relationship.Items
            tempDDL.Items.Add(New ListItem(item.Text, item.Value))
        Next

        If (CType(Me.StateMain, beITP101000).Tables("IQPISSU")).AsEnumerable().Where(Function(x) x.Field(Of String)("NINSREL") = "0").Count >= 1 Then
            tempDDL.SelectedValue = "0"
        End If
        Return tempDDL
    End Function


    Private Sub DirectToEKYC()
        '要保人 
        Dim nins As String = StateMain.Tables("IQPMAIN").Rows(0)("NINS").ToString.Trim
        '要保人ID
        Dim iins As String = StateMain.Tables("IQPMAIN").Rows(0)("IINS").ToString.Trim
        '要保人國籍
        Dim iins_country As String = StateMain.Tables("IQPMAIN").Rows(0)("IINS_COUNTRY").ToString.Trim
        '要保人ID若為居留證，則國籍由EKYC頁面選擇
        If ddl_IINS.SelectedValue.ToString = "3" Then
            iins_country = ""
        Else
            If String.IsNullOrWhiteSpace(iins_country) Then
                iins_country = "TW"
            End If
        End If
        '負責人
        Dim nins_charger As String = ""
        If iins.Length = 8 Then
            nins_charger = StateMain.Tables("IQPMAIN").Rows(0)("NINS_CHARGER").ToString.Trim
        End If
        '主被保人
        Dim nissu As String = ""
        '主被保人ID
        Dim iissu As String = ""
        '主被保人國別
        Dim iissu_country As String = ""
        For Each dr_ISSU As DataRow In StateMain.Tables("IQPISSU").Select(" ISEQ = '1'")
            nissu = dr_ISSU("NISSU").ToString.Trim
            iissu = dr_ISSU("IISSU").ToString.Trim
            iissu_country = dr_ISSU("IISSU_COUNTRY").ToString.Trim
        Next
        '起保日
        Dim dply_bgn As String = StateMain.Tables("IQPMAIN").Rows(0)("DPLY_BGN").ToString.Trim


        PageSession("commandname") = "BHAS_KYCBill"
        Dim argKYC = New Object() {
            txtIQPMainOid.Text,
            nins, iins, iins_country, nins_charger,
            nissu, iissu, iissu_country, dply_bgn, "I", labIOFFICER.Text
        }
        Me.CallOpenModel("BHAS_KYCBill_Model.aspx", "BHAS_KYCBill_Model", "", "NewKYC", argKYC)
    End Sub


    Private Sub ddlIT41_Select(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlIT41.SelectedIndexChanged
        If Not String.IsNullOrEmpty(Me.txtCase.Text) Then
            Dim iPlanNo As String = (Me.txtCase.Text).Substring(0, 1)
            ddlIT4Y.SelectedValue = ""

            If (iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "L") Then
                GetIT44FromIT41(25)
                GetIT43FromIT41(20)
            Else
                GetIT44FromIT41()
            End If
        End If
    End Sub
    Private Sub ddlIT4Y_Select(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlIT4Y.SelectedIndexChanged
        If Not String.IsNullOrEmpty(Me.txtCase.Text) Then
            Dim iPlanNo As String = (Me.txtCase.Text).Substring(0, 1)
            ddlIT41.SelectedValue = ""

            If (iPlanNo = "J" Or iPlanNo = "K" Or iPlanNo = "L") Then
                GetIT44FromIT4Y(25)
                GetIT43FromIT4Y()
            Else
                GetIT44FromIT4Y()
                GetIT43FromIT4Y()
            End If
        End If
    End Sub

    ''' <summary>
    ''' 計算重大燒燙燒傷保額 
    ''' 將 ddlIT41 的值 * XX%（以萬為單位），確保 ddlIT44 有該項並選取，且鎖定 ddlIT44
    ''' </summary>
    ''' <param name="multiplier">比例%，都不傳為解鎖選項</param>
    Private Sub GetIT44FromIT41(Optional multiplier As Integer = -1)
        Try
            ' 先解鎖
            ddlIT44.Enabled = True
            ' 檢核是否有新增過重大燒燙燒傷保額，要刪掉
            For i As Integer = ddlIT44.Items.Count - 1 To 0 Step -1
                Dim v As Integer = 0
                If Integer.TryParse(ddlIT44.Items(i).Value, v) AndAlso v > 250 Then
                    ddlIT44.Items.RemoveAt(i)
                End If
            Next

            Dim src As String = String.Empty
            If Not IsNothing(ddlIT41) Then
                'IT41大人
                src = GetString(ddlIT41.SelectedValue)
            Else
                multiplier = -1
            End If

            ' -1為表示解鎖就好
            If multiplier < 0 Then
                Return
            End If

            ' 沒來源值 -> 選空並鎖定
            If String.IsNullOrWhiteSpace(src) Then
                If Not IsNothing(ddlIT44) Then
                    ' 選擇空白（若有）
                    Dim emptyOpt = ddlIT44.Items.FindByValue(String.Empty)
                    If Not IsNothing(emptyOpt) Then
                        ddlIT44.SelectedValue = String.Empty
                    ElseIf ddlIT44.Items.Count > 0 Then
                        ddlIT44.SelectedIndex = 0
                    End If
                    ddlIT44.Enabled = False
                End If
                Return
            End If

            ' 解析來源數值（ddlIT41 的值為萬，例如 100 表示 100 萬）
            Dim n As Double = 0
            If Double.TryParse(src, n) Then
                ' 計算目標：ddlIT41 * 0.01 (1%)，並四捨五入為整數萬
                Dim targetWan As Integer = CInt(Math.Round(n * 0.01 * multiplier))
                Dim sval As String = targetWan.ToString()

                ' 若 ddlIT44 沒有該值則新增
                If ddlIT44.Items.FindByValue(sval) Is Nothing Then
                    ddlIT44.Items.Add(New ListItem(sval, sval))
                End If

                ' 選取並鎖定
                Try
                    ddlIT44.SelectedValue = sval
                Catch ex As Exception
                    ' 若選取失敗，保留最接近的選項（可根據需求改為 log）
                    If ddlIT44.Items.Count > 0 Then ddlIT44.SelectedIndex = ddlIT44.Items.Count - 1
                End Try
                ddlIT44.Enabled = False
            Else
                ' 無法解析來源值 -> 鎖定並不設定具體值
                If Not IsNothing(ddlIT44) Then ddlIT44.Enabled = False
            End If
        Catch ex As Exception
            ' 可以記錄或忽略
        End Try
    End Sub

    ''' <summary>
    ''' 計算大眾運輸保額
    ''' 將 ddlIT41 的值 * XX%（以萬為單位），確保 ddlIT43 有該項並選取，且鎖定 ddlIT43
    ''' </summary>
    ''' <param name="multiplier">比例%，都不傳為解鎖選項</param>
    Private Sub GetIT43FromIT41(Optional multiplier As Integer = -1)
        Try
            ' 先解鎖
            ddlIT43.Enabled = False

            Dim src As String = String.Empty
            If Not IsNothing(ddlIT41) Then
                'IT41大人
                src = GetString(ddlIT41.SelectedValue)
            Else
                multiplier = -1
            End If

            ' -1為表示解鎖就好
            If multiplier < 0 Then
                Return
            End If

            ' 沒來源值 -> 選空並鎖定
            If String.IsNullOrWhiteSpace(src) Then
                If Not IsNothing(ddlIT43) Then
                    ' 選擇空白（若有）
                    Dim emptyOpt = ddlIT43.Items.FindByValue(String.Empty)
                    If Not IsNothing(emptyOpt) Then
                        ddlIT43.SelectedValue = String.Empty
                    ElseIf ddlIT43.Items.Count > 0 Then
                        ddlIT43.SelectedIndex = 0
                    End If
                    ddlIT43.Enabled = False
                End If
                Return
            End If

            ' 解析來源數值（ddlIT43 的值為萬，例如 100 表示 100 萬）
            Dim n As Double = 0
            If Double.TryParse(src, n) Then
                ' 計算目標：ddlIT43 * 0.01 (1%)，並四捨五入為整數萬
                Dim targetWan As Integer = CInt(Math.Round(n * 0.01 * multiplier))
                Dim sval As String = targetWan.ToString()

                ' 若 ddlIT43 沒有該值則新增
                If ddlIT43.Items.FindByValue(sval) Is Nothing Then
                    ddlIT43.Items.Add(New ListItem(sval, sval))
                End If

                ' 選取並鎖定
                Try
                    ddlIT43.SelectedValue = sval
                Catch ex As Exception
                    ' 若選取失敗，保留最接近的選項（可根據需求改為 log）
                    If ddlIT43.Items.Count > 0 Then ddlIT43.SelectedIndex = ddlIT43.Items.Count - 1
                End Try
                ddlIT43.Enabled = False
            Else
                ' 無法解析來源值 -> 鎖定並不設定具體值
                If Not IsNothing(ddlIT43) Then ddlIT43.Enabled = False
            End If
        Catch ex As Exception
            ' 可以記錄或忽略
        End Try
    End Sub

    ''' <summary>
    ''' 計算重大燒燙燒傷保額
    ''' 將 ddlIT4Y 的值 * XX%（以萬為單位），確保 ddlIT44 有該項並選取，且鎖定 ddlIT44
    ''' </summary>
    ''' <param name="multiplier">比例%，都不傳為解鎖選項</param>
    Private Sub GetIT44FromIT4Y(Optional multiplier As Integer = -1)
        Try
            ' 先解鎖
            ddlIT44.Enabled = True
            ' 檢核是否有新增過重大燒燙燒傷保額，要刪掉
            For i As Integer = ddlIT44.Items.Count - 1 To 0 Step -1
                Dim v As Integer = 0
                If Integer.TryParse(ddlIT44.Items(i).Value, v) AndAlso v > 250 Then
                    ddlIT44.Items.RemoveAt(i)
                End If
            Next

            'For Each item As ListItem In ddlIT44.Items.ToArray()
            '    Dim v As Integer = 0
            '    If Integer.TryParse(item.Value, v) AndAlso v > 250 Then
            '        If ddlIT44.Items.Contains(item) Then
            '            ddlIT44.Items.Remove(item)
            '        End If
            '    End If
            'Next

            Dim src As String = String.Empty
            If Not IsNothing(ddlIT4Y) Then
                'IT4Y
                src = GetString(ddlIT4Y.SelectedValue)
            Else
                multiplier = -1
            End If

            ' -1為表示解鎖就好
            If multiplier < 0 Then
                Return
            End If

            ddlIT44.SelectedValue = 50
            ddlIT44.Enabled = False
        Catch ex As Exception
            ' 可以記錄或忽略
        End Try
    End Sub

    ''' <summary>
    ''' 計算重大燒燙燒傷保額
    ''' 將 ddlIT4Y 的值 * XX%（以萬為單位），確保 ddlIT43 有該項並選取，且鎖定 ddlIT43
    ''' </summary>
    ''' <param name="multiplier">比例%，都不傳為解鎖選項</param>
    Private Sub GetIT43FromIT4Y(Optional multiplier As Integer = -1)
        Try
            ' 先解鎖
            ddlIT43.SelectedValue = ""
            ddlIT43.Enabled = False
        Catch ex As Exception
            ' 可以記錄或忽略
        End Try
    End Sub

    Private Function GetUserIDFromTextBox() As String
        Return Me.txtCompanyid.Text.Trim & "\" & Me.txtUserid.Text.Trim
    End Function

    ''' <summary>
    ''' 從Statemain被保人的IPlanNo判斷該單當前方案為ITP or IQP。因被保人方案需要與主被保人一致，故基本[增、修被保人]把關無誤，此處IPlanNo應一致。
    ''' （原本使用ddlTargetProject當作目標方案，但user增、修完畢被保人後，可能改點其他方案但沒有增、修被保人。此狀態下依ddlTargetProject當作目標方案會錯誤。）
    ''' </summary>
    ''' <returns></returns>
    Private Function GetTargetProject() As String
        Dim TargetProject As String = ""

        '因修改被保人，方案改動僅能改主被保人，且主被保人改動、所有被保人方案會一同調整
        Dim boo As Boolean = True ' 預設為 True
        Dim IISSUP1anNo As String = getIISSUPlanNo()

        '若僅人身保險會是空的
        If IISSUP1anNo = "A" Or IISSUP1anNo = "C" Or IISSUP1anNo = "D" Or IISSUP1anNo = "J" Or IISSUP1anNo = "K" Or IISSUP1anNo = "L" Or IISSUP1anNo = "M" Or IISSUP1anNo = "N" Or IISSUP1anNo = "O" Then
            TargetProject = "ITP"
        ElseIf String.IsNullOrEmpty(IISSUP1anNo) Then
            TargetProject = "IQP"
        Else
            Throw New Exception($"被保人方案代碼{IISSUP1anNo}不在處理範圍內")
        End If

        Return TargetProject
    End Function

    Private Function getIISSUPlanNo() As String
        Dim rows = StateMain.Tables("IQPISSU").AsEnumerable().Where(Function(r) Not IsDBNull(r("ISEQ")) AndAlso r("ISEQ").ToString() <> "")
        Dim IISSUP1anNo As String = ""
        Dim GetPlanNo

        If Not rows.Any() Then
            ' 情況：完全沒有符合條件的資料
            Throw New Exception("無被保人資料，建議再資料更新後操作。")
        Else
            ' 情況：有資料，取出第一筆
            GetPlanNo = rows.OrderBy(Function(r) Convert.ToInt32(r("ISEQ"))).Select(Function(r) r("IPLANNO")).FirstOrDefault()

            ' 安全處理：處理可能為空的 IPLANNO
            IISSUP1anNo = If(GetPlanNo, "").ToString.Trim.ToUpper
        End If
        Return IISSUP1anNo
    End Function

    ''' <summary>
    ''' 刪除被保人後，將被保人重新排序，如果ISEQ大於被刪除的被保人ISEQ，則ISEQ -1 往前遞補。
    ''' ex:刪除ISEQ 1 (主被保人)，則ISEQ 2 - 1往前遞補成為主被保人
    ''' 若遞補後沒有ISEQ = 1，則取當前ISEQ最小值設為"1"
    ''' </summary>
    ''' <param name="ISEQ_Now"></param>
    Private Sub ResetISEQAfterDEL(ISEQ_Now As String, TBName As String)
        Dim ISEQCount As Int32 = 1
        For Each dr As DataRow In StateMain.Tables(TBName).Rows
            If CType(dr("ISEQ"), Int32) > ISEQ_Now Then
                dr("ISEQ") = (CType(dr("ISEQ"), Int32) - 1).ToString
            End If
        Next

        '將ISEQ最小值設為'1'
        Dim minRow = StateMain.Tables(TBName).AsEnumerable() _
                    .Where(Function(r) Not IsDBNull(r("ISEQ")) AndAlso r("ISEQ").ToString() <> "") _
                    .OrderBy(Function(r) Convert.ToInt32(r("ISEQ"))) _
                    .FirstOrDefault()

        If minRow IsNot Nothing Then
            minRow("ISEQ") = "1"
        End If


        StateMain.Tables(TBName).AcceptChanges()
    End Sub

    Private Sub ShowPremInfoOrNot(show As Boolean)
        lblsum01.Visible = show
        lblsum02.Visible = show
        sum01.Visible = show
        sum02.Visible = show
        lblTotal.Visible = show
        txtTotal.Visible = show
        'If dgdMain.Items.Count > 0 Then
        '    For Each column As DataGridColumn In dgdMain.Columns
        '        If column.HeaderText = "保費(NTD)" Then
        '            column.Visible = show
        '            Exit For
        '        End If
        '    Next
        'End If
    End Sub

    Private Function DateAMPM(inputDate As String) As String
        Return inputDate.Replace("上午", "AM").Replace("下午", "PM")
    End Function

    '####程式中有一處呼叫此function，若要復原下方function須留意
    'Protected Sub rad_Disease_CheckedChanged(sender As Object, e As EventArgs) Handles rad_Disease_N.CheckedChanged, rad_Disease_Y.CheckedChanged
    '	If rad_Disease_N.Checked Then
    '		'ddlIT61.Enabled = True
    '		'ddlIT63.Enabled = False
    '		ddlIT63.SelectedValue = ""
    '	Else
    '		'ddlIT61.Enabled = False
    '		'ddlIT63.Enabled = True
    '		ddlIT61.SelectedValue = ""
    '	End If
    'End Sub

    Public Function IsMobileNumber(ByVal instr As String) As Boolean
        ' 檢查傳入的字串是否為空值或 Nothing
        If String.IsNullOrWhiteSpace(instr) Then
            Return False
        End If

        ' 定義台灣手機號碼的正規表達式：以 09 開頭，後面接著 8 個數字
        Dim pattern As String = "^09\d{8}$"

        ' 回傳比對結果 (符合回傳 True，否則回傳 False)
        Return Regex.IsMatch(instr, pattern)
    End Function


#Region "航班快速理賠"

    Protected Sub btnGetFlightInfo_Click(sender As Object, e As EventArgs) Handles btnGetFlightInfo.Click
        CheckSNNN.Text = ""
        '日期檢核
        Dim bolDateChk As Boolean = True
        Try
            If IsDate((UcRocDobjbgn.GetDate).ToString("yyyy/MM/dd")) = False Then
                bolDateChk = False
            End If
            If IsDate((UcRocDobjend.GetDate).ToString("yyyy/MM/dd")) = False Then
                bolDateChk = False
            End If
        Catch ex As Exception
            bolDateChk = False
        End Try
        If Not bolDateChk Then
            Me.UIAppendError("請正確填寫 保險期間")
            Return
        End If

        Dim bolOutbound As Boolean = False
        Dim bolReturn As Boolean = False

        '115/08/19說不檢核去、回航程的數量
        'If Me.StateMain.Tables.Contains("HASM_FLIGHT_DETAIL") AndAlso Me.StateMain.Tables("HASM_FLIGHT_DETAIL").Rows.Count > 0 Then
        '    bolOutbound = Me.StateMain.Tables("HASM_FLIGHT_DETAIL").AsEnumerable() _
        '        .Any(Function(r) Not IsDBNull(r("Voyage")) AndAlso r("Voyage").ToString().Contains("去程"))
        '
        '    bolReturn = Me.StateMain.Tables("HASM_FLIGHT_DETAIL").AsEnumerable() _
        '        .Any(Function(r) Not IsDBNull(r("Voyage")) AndAlso r("Voyage").ToString().Contains("回程"))
        'End If

        PageSession("commandname") = "FlightSearch"
        Dim argKYC = New Object() {
            (UcRocDobjbgn.GetDate).ToString("yyyy/MM/dd"),
            txtTobjbgn.Text,
            (UcRocDobjend.GetDate).ToString("yyyy/MM/dd"),
            txtTobjend.Text, bolOutbound, bolReturn
        }
        Me.CallOpenModel("BHAS_FlightSearch.aspx", "BHAS_FlightSearch", "", "GetFlight", argKYC)
    End Sub

    Private Sub dgdFlight_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dgdFlight.ItemCommand
        Try

            Dim view As DataView = Me.AppGetDataGridBoundDataView(Me.dgdFlight, Me.StateMain)
            If (e.Item.DataSetIndex) = -1 Then
                Exit Sub
            End If
            If e.CommandName = "DeleteFlight" Then
                CheckSNNN.Text = ""
                Dim rowView As DataRowView = view(e.Item.DataSetIndex)

                Dim row As DataRow = rowView.Row


                '刪除dgdFlight 
                Dim ISWQ_D As String = row("ISEQ")
                For Each drD As DataRow In StateMain.Tables("HASM_FLIGHT_DETAIL").Select("ISEQ = '" & ISWQ_D & "'")
                    drD.Delete()
                Next

                '立即針對Table做刪除，避免後續查找StateMain時發生錯誤
                StateMain.Tables("HASM_FLIGHT_DETAIL").AcceptChanges()
                'ISEQ重新給號
                ResetISEQAfterDEL(ISWQ_D, "HASM_FLIGHT_DETAIL")

                BaseSetChildFormData(dgdFlight, StateMain)
            End If


        Catch ex As Exception
            ProcessException(ex)
            'Me.UIAppendError(ex.Message.ToString)
        End Try
    End Sub


    ''' <summary>
    ''' 判斷航班時間是否在保險範圍內
    ''' </summary>
    Private Function ConfirmFlightDateTime(fDepartureStr As DateTime, fArrivalStr As DateTime, DPLY_BGN As String, DPLY_END As String, TPLY_BGN As Integer, TPLY_END As Integer) As Tuple(Of Boolean, String)
        ' fDepartureDateTime 航班出發日期
        ' fArrivalDateTime 航班到達日期
        ' iDepartureDateTime 保險出發日期
        ' iarrivalDateTime 保險到達日期
        ' 航班出發日期以及航班到達日期 都必須在保險範圍內才顯示航班
        Try

            Dim iDepartureStr As String = DPLY_BGN & " " & TPLY_BGN.ToString("D2") & ":00"
            Dim iArrivalStr As String = DPLY_END & " " & TPLY_END.ToString("D2") & ":00" ' 時間一率用dateFormTM為主   

            Dim iDepartureDateTime As DateTime = DateTime.Parse(DateAMPM(iDepartureStr))
            Dim iarrivalDateTime As DateTime = DateTime.Parse(DateAMPM(iArrivalStr))
            Dim fDepartureDateTime As DateTime = DateTime.Parse(DateAMPM(fDepartureStr))
            Dim fArrivalDateTime As DateTime = DateTime.Parse(DateAMPM(fArrivalStr))

            If fDepartureDateTime >= iDepartureDateTime AndAlso fDepartureDateTime <= iarrivalDateTime AndAlso fArrivalDateTime >= iDepartureDateTime AndAlso fArrivalDateTime <= iarrivalDateTime Then
                Return Tuple.Create(Of Boolean, String)(True, "")
            End If

        Catch ex As Exception
            Return Tuple.Create(Of Boolean, String)(False, "航班時間判斷發生異常...")
        End Try

        Return Tuple.Create(Of Boolean, String)(False, "航班時間必須在保險期間")
    End Function

    Private Function CheckAllowFlightOrNot() As String
        Dim ISSUCount As Int32 = 0
        Dim FlightCount As Int32 = 0
        Dim result As String = ""
        Dim chkResult As Boolean = True
        Dim flagPlanNo As String = ""
        Dim firstRound As Boolean = True


        If StateMain.Tables.Contains("HASM_FLIGHT_DETAIL") Then
            For Each dr4 As DataRow In StateMain.Tables("HASM_FLIGHT_DETAIL").Rows
                If dr4.RowState <> DataRowState.Deleted Then
                    FlightCount += 1
                    Exit For
                End If
            Next
        End If

        If FlightCount > 0 Then
            For Each row As DataRow In StateMain.Tables("IQPISSU").Rows
                If row.RowState <> DataRowState.Deleted Then

#Region "檢查被保險人方案是否一致"
                    If firstRound Then
                        flagPlanNo = row("IPLANNO").ToString
                    Else
                        If flagPlanNo <> row("IPLANNO").ToString Then
                            result = txtIQPMainOid.Text & " 被保險人方案不一致。"
                            Exit For
                        End If
                    End If

                    ISSUCount += 1
                    firstRound = False
#End Region
                End If
            Next

            '檢查航班卡控
            If String.IsNullOrEmpty(result) Then
                If flagPlanNo = "A" Or flagPlanNo = "L" Or String.IsNullOrEmpty(flagPlanNo) Then
                    If checkSpecialChannel().Item1 = True Or checkChannelO().Item1 = True Then
                        result = "「僅人身保險、國內海陸型、一般型或申根型」不適用〔班機延誤快速理賠服務〕，請將航班清空。"
                    Else
                        result = "「僅人身保險、國內海陸型」不適用〔班機延誤快速理賠服務〕，請將航班清空。"
                    End If
                ElseIf ISSUCount > 9 Then
                    result = "〔班機延誤快速理賠服務〕僅限於9位(含)以下被保險人，請將航班清空。"
                End If
            End If
        End If

        Return result
    End Function

#End Region

End Class

Public Class DiseaseFields
    Public Property IT61 As Boolean
    Public Property IT63 As Boolean
    Public Property IT64 As Boolean
    Public Property IT65 As Boolean
    Public Property IT66 As Boolean
    Public Property IT6D As Boolean
    Public Property IT6E As Boolean
    Public Property IT6F As Boolean
    Public Property IT6G As Boolean
End Class
