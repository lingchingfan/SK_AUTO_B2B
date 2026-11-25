<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="BHAS101003.aspx.vb" Inherits="B2BWeb.BHAS101003" %>

<%@ Register TagPrefix="uc1" TagName="ucDate" Src="../../ISNEControl/ucDate.ascx" %>
<%@ Register TagPrefix="uc1" TagName="ucRocDate" Src="../../ISNEControl/ucRocDate.ascx" %>
<%--<%@ Register TagPrefix="uc1" TagName="ucRocDate" Src="../../PUBControl/CMN/ucRocDateJS.ascx" %>--%>
<%@ Register TagPrefix="uc2" TagName="ucDate" Src="../../ISNEControl/ucDate.ascx" %>

<%@ Register TagPrefix="uc1" TagName="ApplicationBanner" Src="../../ISNEControl/ApplicationBanner.ascx" %>
<%@ Register TagPrefix="uc1" TagName="DataGridBanner" Src="../../ISNEControl/DataGridBanner.ascx" %>
<%@ Register TagPrefix="uc1" TagName="ucCmnZipCodeNew" Src="~/PubControl/CMN/ucCmnZipCodeNew.ascx" %>
<%@ Register TagPrefix="ucbs" TagName="ucBSModel" Src="../ucBSModel.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title><%=ApplicationTitle%>
    </title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../Styles.css" type="text/css" rel="stylesheet">
    <script src="../CAS/Scripts/sweetalert.min.js"></script>
    <link rel="stylesheet" type="text/css" href="../CAS/css/sweetalert.min.css" />
    <%: Scripts.Render("~/bundles/jquery") %>
    <link rel="stylesheet" type="text/css" href="../CAS/css/jquery-ui.min.css">
    <script src="../CAS/js/jquery-ui.min.js"></script>
    <script language="javascript">

        //$(function () {
        //    $(".datepicker").datepicker({
        //        showAnim: "blind",
        //        showOtherMonths: true,
        //        selectOtherMonths: true,
        //        showOn: "button",
        //        buttonImage: "../CAS/css/images/calendar.gif",
        //        buttonImageOnly: true,
        //        buttonText: "Select date",
        //        changeMonth: true,
        //        changeYear: true,
        //        dateFormat: "yymmdd",
        //        minDate: 0,
        //        showButtonPanel: true
        //    });

        //    $(".datepicker").on("change", function () {
        //        $(this).val(DateT($(this).val()));
        //    });
        //});

        $(function () {
            if (document.getElementById("CHK_P2").checked)
                document.getElementById("ddlIISSU_COUNTRY").style.display = "block";

            ChangeMF0();
            chkRelationshipVal();
            //radDiseaseClick("optionN");
            if ($("#hidUrl").val() != '') {
                $("#aPrint").attr('href', $("#hidUrl")[0].textContent);
                $("#hidUrl").val('');
                $("#aPrint")[0].click();

            }
            //20230419 加show message
            var str = document.getElementById('Applicationbanner1_StatusBar1_lblMessage_lblText').innerHTML;
            if (str == '') {
            } else {
                swal("系統訊息", str + "\n", "warning");
                //alert(str);
            }
            $("#show").click(showHandler);
            $("#hide").click(hideHandler);

            $("#uploadM").click(function () {
                $("#p1").show();

                //20250514 詢問使用者要此次要傳送IQP or ITP的名冊，並設定ddlTargetProject & lblTargetPlan & txtCase的值，
                //  因為user提IQP & ITP方案選擇不使用下拉，改透過方案選擇按鈕設定。但按鈕點選會檢核、並要求User需填被保人資訊。 
                ExlUpCheck();
            });

            $('#Relationship').on("change", function () {
                $("#txtRelationshipOther").val('');
                chkRelationshipVal();
            });


            if ($("#<%= rad_Disease_Y.ClientID %>").is(':checked')) {
                yesit63();
            }
            else {
                noit63();
            }

            if (document.getElementById("<%= CHK_P1.ClientID %>") != null) {
                if (document.getElementById("<%= CHK_P1.ClientID %>").checked) {
                    document.getElementById("<%= ddlIINS_COUNTRY.ClientID %>").style.display = "block";
                }
            }
        });

        $(function () {
            $("#dllGetIrec").blur(
                function () {
                    var $dllGetIrec = $("#dllGetIrec");
                    var sValue = $dllGetIrec.val();
                    var arr1;
                    if (sValue != '' && sValue != null) {
                        arr1 = sValue.split("|");
                        $("#txtIrecunit").val(arr1[0]);
                        $("#txtNrecpesn").val(arr1[1]);
                        $("#txtIREC").val(arr1[2]);
                    }
                });
        });
        var isShow = false;
        function change() {
            if (!isShow) {
                isShow = true;
                document.getElementById('ShowS2').style.display = '';
            }
            else {
                isShow = false;
                document.getElementById('ShowS2').style.display = 'none';
            }
        }

        //點選uploadM(名冊上傳)時，詢問該次excel匯入為IQP or ITP，依user選擇隨之設定ddlTargetProject & lblTargetPlan & txtCase的值
        function ExlUpCheck() {
            var ans = confirm("請問此次作業是否處理﹝旅綜﹞投保?\n按﹝確定﹞為﹝旅綜﹞作業；\n按﹝取消﹞為﹝旅平﹞作業");
            if (ans) {
                document.getElementById('ddlTargetProject').value = "ITP";
                document.getElementById('txtCase').value = "";
                document.getElementById('lblTargetPlan').value = "";
            }
            else {
                document.getElementById('ddlTargetProject').value = "IQP"
                document.getElementById('txtCase').value = "E 不選擇方案";
                document.getElementById('lblTargetPlan').value = "";
            }
        }

        function confirmUncheck(chk) {
            if (!chk.checked) {
                var ans = confirm("取消勾選會將當前登打的被保人資料清空，請問是否繼續?");
                if (ans) {
                    return true;
                }
                else {
                    return false;
                }
            }
            return true;
        }

        var isShow2 = false;
        function change2() {
            if (!isShow2) {
                isShow2 = true;
                document.getElementById('ShowS3').style.display = '';
                document.getElementById('ShowS4').style.display = '';
            }
            else {
                isShow2 = false;
                document.getElementById('ShowS3').style.display = 'none';
                document.getElementById('ShowS4').style.display = 'none';
            }
        }
        //var isNono = false;

        //F114010802 旅遊險b2b被保人與要保人關係#9078(1.自然人拿掉關係8、9項目2.法人只能選8、9項目(含其他)3.選其他必需輸入最長5個字SKERP_B2B.IQPISSU. JOBOFFICE)
        function chkRelationshipVal() {
            //if ($("#Relationship option:selected").val() == "0") {
            //    $("#txtRelationshipOther").show();
            //} else {
            //    $("#txtRelationshipOther").hide();
            //}
        }

        function chkboxChange1() {
            ChangeMF0();
            if (document.getElementById("<%= CHK_P1.ClientID %>").checked)
                document.getElementById("<%= ddlIINS_COUNTRY %>").style.display = "block";
            else
                document.getElementById("<%= ddlIINS_COUNTRY %>").style.display = "none";
        }

        function chkboxChange2() {
            ChangeMF();
            if (document.getElementById("CHK_P2").checked)
                document.getElementById("ddlIISSU_COUNTRY").style.display = "block";
            else
                document.getElementById("ddlIISSU_COUNTRY").style.display = "none";
        }

        function TxtToupper(tagId) {
            var txt = document.getElementById(tagId)
            txt.value = txt.value.toUpperCase();
        }

        function doBeforeSubmit() {
            ChangeMF0();
        }

        // 點選radioButton海突是否含傳染病時，是：海突(不含傳染病)的下拉.enable = false ； 否：海突(含傳染病)的下拉.enable = false
        function radDiseaseClick(value) {
            //rad_Disease_N
            if (value == "optionN") {
                var ddlN = document.getElementById('<%= ddlIT61.ClientID %>');
                var ddlY = document.getElementById('<%= ddlIT63.ClientID %>');
                ddlN.disabled = false;
                ddlN.enabled = true;
                ddlY.disabled = true;
                ddlY.enabled = false;
                ddlY.selectedIndex = 0;
            }
            else {
                //$("#ddlIT61").disableSelection();
                var ddlN = document.getElementById('<%= ddlIT61.ClientID %>');
                var ddlY = document.getElementById('<%= ddlIT63.ClientID %>');
                ddlN.disabled = true;
                ddlN.enabled = false;
                ddlY.disabled = false;
                ddlY.enabled = true;
                ddlN.selectedIndex = 0;
            }
        }

        //     //IT63
        //     $(function () {
        //$('#ddlIT63').on('change',ddlChange63);
        //ddlChange63();
        //})
        //     function ddlChange63() {
        //         var sWorkType = $('#ddlIT63').val();

        //     	switch(sWorkType){
        //     		case "":
        //     				document.getElementById("txtIT63").value="";
        //     				TotalSum();
        //     			break;
        //             default:
        //                 document.getElementById("txtIT63").value = sWorkType;
        //     	}
        //     }

    </script>
    <style type="text/css">
        .auto-style2 {
            height: 27px;
        }

        .auto-style3 {
            font-size: large;
            color: #FF0000;
            text-decoration: underline;
        }

        .auto-style4 {
            border: 1px solid #abd0e0;
            padding: 1px;
            background-color: #CBE9F5;
            color: #2C2C85;
            text-align: right;
            letter-spacing: 1px;
            white-space: nowrap;
            word-spacing: 3px;
            height: 26px;
        }

        .auto-style5 {
            border: 1px solid #abd0e0;
            padding: 1px;
            background-color: #E0EEF7;
            color: #010305;
            text-align: left;
            height: 26px;
        }

        .auto-style6 {
            text-align: right;
        }

        .auto-style7 {
            font-size: medium;
        }

        .auto-style8 {
            border: 1px solid #abd0e0;
            padding: 1px;
            background-color: #E0EEF7;
            color: #010305;
            text-align: left;
            font-size: medium;
        }

        .auto-style9 {
            font-size: small;
        }

        .auto-style10 {
            border: 1px solid #abd0e0;
            padding: 1px;
            background-color: #CBE9F5;
            color: #2C2C85;
            text-align: right;
            letter-spacing: 1px;
            white-space: nowrap;
            word-spacing: 3px;
            height: 57px;
        }

        .auto-style11 {
            border: 1px solid #abd0e0;
            padding: 1px;
            background-color: #E0EEF7;
            color: #010305;
            text-align: left;
            height: 57px;
        }

        .auto-style12 {
            color: #0000CC;
        }

        .txt150 {
            width:150px !important;
        }

        .button {
            width: 20px;
            text-align: left;
            font-size: 1em;
            text-decoration: underline;
            color: blue;
        }

            .button:hover .hover {
                display: block;
            }

        .hover {
            display: none;
        }

        .auto-style13 {
            /* 	background-color:Silver; 	border-right: gray 1px solid; 	border-top: gray 1px solid; 	text-justify: distribute-all-lines; 	font-size: 10pt; 	background-image: url(images/bgSilverX.gif); 	border-left: gray 1px solid; 	cursor: pointer; 	color: black;	 	padding-top: 3px; 	border-bottom: gray 1px solid; 	background-repeat: repeat-x; 	font-family: 'Courier New' , 細明體; 	letter-spacing: 2px; 	border-collapse: collapse; 	text-align: center; */
            padding-top: 2px;
            background-color: Silver;
            border: solid 1px gray;
            font-size: 10pt;
            background-image: url('../../images/bgSilverX24.gif');
            cursor: pointer;
            color: black;
            background-repeat: repeat-x;
            font-family: 'Courier New', 細明體;
            text-align: center;
        }

        #Form1 {
            color: #17345f;
        }

        table.b2b-form-layout {
            display: block;
            width: 100% !important;
            margin: 12px auto;
            border: 1px solid #d5deeb;
            border-radius: 8px;
            border-collapse: separate;
            border-spacing: 0;
            background: #fff;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(27, 55, 94, .06);
        }

        table.b2b-form-layout > tbody {
            display: block;
        }

        table.b2b-form-layout > tbody > tr {
            display: grid;
            grid-template-columns: repeat(10, minmax(0, 1fr));
            border-bottom: 1px solid #e6ebf2;
        }

        table.b2b-form-layout > tbody > tr:last-child {
            border-bottom: 0;
        }

        table.b2b-form-layout > tbody > tr > td {
            display: block;
            min-width: 0;
            box-sizing: border-box;
            padding: 8px 10px;
            vertical-align: top;
            overflow-wrap: anywhere;
            border: 0 !important;
        }

        table.b2b-form-layout td[colspan="2"] { grid-column: span 2; }
        table.b2b-form-layout td[colspan="3"] { grid-column: span 3; }
        table.b2b-form-layout td[colspan="4"] { grid-column: span 4; }
        table.b2b-form-layout td[colspan="5"] { grid-column: span 5; }
        table.b2b-form-layout td[colspan="6"] { grid-column: span 6; }
        table.b2b-form-layout td[colspan="7"] { grid-column: span 7; }
        table.b2b-form-layout td[colspan="8"] { grid-column: span 8; }
        table.b2b-form-layout td[colspan="9"] { grid-column: span 9; }
        table.b2b-form-layout td[colspan="10"] { grid-column: span 10; }

        table.b2b-form-layout .tdDataCaption {
            color: #233f6e;
            font-weight: 600;
            text-align: left;
            white-space: normal;
            background: #f7f9fc;
        }

        table.b2b-form-layout .tdDataField,
        table.b2b-form-layout .TDInput,
        table.b2b-form-layout .auto-style10,
        table.b2b-form-layout .auto-style11 {
            color: #263a59;
            background: #fff;
        }

        table.b2b-form-layout tr.b2b-section-heading {
            margin-top: 10px;
            border-top: 1px solid #d8e1ee;
            border-bottom: 1px solid #d8e1ee;
        }

        table.b2b-form-layout tr.b2b-section-heading > td {
            color: #17345f;
            font-weight: 600;
            background: #f1f5fa !important;
        }

        table.b2b-form-layout input[type="text"],
        table.b2b-form-layout input[type="email"],
        table.b2b-form-layout select,
        table.b2b-form-layout textarea {
            box-sizing: border-box;
            max-width: 100%;
            min-height: 34px;
            padding: 6px 9px;
            color: #17345f;
            border: 1px solid #c5d0e0;
            border-radius: 5px;
            background-color: #fff;
        }

        table.b2b-form-layout input[type="checkbox"],
        table.b2b-form-layout input[type="radio"] {
            accent-color: #1d3b6d;
        }

        table.b2b-form-layout input[type="button"],
        table.b2b-form-layout input[type="submit"],
        table.b2b-form-layout button {
            min-height: 34px;
            padding: 6px 12px;
            color: #17345f;
            border: 1px solid #c5d0e0;
            border-radius: 5px;
            background: #fff;
            cursor: pointer;
        }

        table.b2b-form-layout input[type="button"]:hover,
        table.b2b-form-layout input[type="submit"]:hover,
        table.b2b-form-layout button:hover {
            background: #f1f5fa;
        }

        @media (max-width: 900px) {
            table.b2b-form-layout > tbody > tr {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }

            table.b2b-form-layout td[colspan] {
                grid-column: 1 / -1;
            }
        }

        @media (max-width: 560px) {
            #Form1 {
                padding: 0 8px;
            }

            table.b2b-form-layout > tbody > tr {
                grid-template-columns: minmax(0, 1fr);
            }

            table.b2b-form-layout td:not([colspan]) {
                grid-column: 1 / -1;
            }

            table.b2b-form-layout > tbody > tr > td {
                padding: 7px 10px;
            }
        }
    </style>
</head>
<body ms_positioning="FlowLayout">
    <form id="Form1" method="post" runat="server">
        <ucbs:ucBSModel ID="ucBSModelModal" runat="server" />
        <asp:Label ID="Label1" ForeColor="Red" runat="server" Style="padding: 10px; font-size: 20pt;"></asp:Label>
        <input id="hidUrl"
            type="hidden" name="hidUrl" runat="server">
        <a
            style="display: none" id="aPrint" href="" target="_blank"
            download="DownLoad"></a>
        <uc1:applicationbanner id="Applicationbanner1" runat="server"></uc1:applicationbanner>
        <table border="0" cellspacing="0" cellpadding="0" width="95%" align="center">
            <tr>
                <td style="width: 20%"><font face="新細明體" size="3">
                    <p align="left">
                        報價單號：<asp:Label ID="txtIQUO" runat="server"></asp:Label>
                </font>
                    </P></td>
                <td style="width: 20%">
                    <font face="新細明體" size="3">繳費單號：<asp:Label ID="txtPAY_NUM" runat="server"></asp:Label></font></td>
                <td>
                    <p align="left">
                        <font face="新細明體" size="3">要保號：<asp:Label ID="txtIQPMainOid" runat="server" ReadOnly="true"></asp:Label><asp:TextBox Style="display: none" ID="txtRtype" runat="server" ReadOnly="True" MaxLength="1" BorderStyle="None" Width="20px" ForeColor="Black"></asp:TextBox>
                            <asp:Label ID="txtPESTATUS_NAME" runat="server"></asp:Label><asp:Label ID="txtPESTATUS" runat="server"></asp:Label></font>
                    </p>
                </td>
                <td style="width: 20%"><font face="新細明體" size="3">KYC序號：<asp:Label ID="txtIKYC" runat="server"></asp:Label></font></td>
                <td style="width: 20%"><font face="新細明體" size="3">總保費：<asp:Label ID="txtTOTPREM_NEW" runat="server"></asp:Label></font></td>
            </tr>
        </table>
        <table border="0" cellspacing="0" cellpadding="0" width="95%" align="center">
            <tr>
                <td class="tdDataField">
                    <asp:DropDownList ID="ddlTargetProject" runat="server" Width="200" AutoPostBack="true" style="display:none">
                        <asp:ListItem Value="ITP" Selected="True">ITP 旅行綜合險</asp:ListItem>
                        <asp:ListItem Value="IQP">IQP 旅行平安險</asp:ListItem>
                    </asp:DropDownList>
                    <strong>方案：<asp:Label ID="lblTargetPlan" runat="server" Text="" Style="color: red; font-weight: bold;"></asp:Label></strong>
                    <asp:TextBox  ID="txtIsSPEChannel" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off" style="display:none"/>
                    <asp:HiddenField ID="hdValidateFlag" runat="server" Value="0" />
                </td>
            </tr>
            <tr>
                <td>
                    <table class="tblDataArea b2b-form-layout" width="100%" align="center">
                        <tbody>
                            <tr>
                                <td style="background-color: #99ccff;" colspan="10" class="auto-style2">

                                    <strong>
                                        <asp:Label ID="Label14" runat="server" Text="[通路代碼]"></asp:Label>
                                    </strong>
                                    <asp:TextBox ID="txtCompanyid" runat="server" MaxLength="6" Width="50px" Enabled="false"></asp:TextBox>
                                    <strong>
                                        <asp:Label ID="Label15" runat="server" Text=" [帳號]"></asp:Label>
                                        <asp:TextBox ID="txtUserid" runat="server" MaxLength="10" Width="80px" Enabled="false"></asp:TextBox>
                                    </strong>
                                    <strong>
                                        <asp:Label ID="Label3" runat="server" Text="[登錄證字號]"></asp:Label>
                                    </strong>
                                    <asp:TextBox  ID="txtIRECPESN_TxtIn" runat="server" MaxLength="30" Width="100" Enabled="false"></asp:TextBox>&nbsp;
                                    <asp:CheckBox ID="chkAssignACC" runat="server" Checked ="True" AutoPostBack="true" onclick="if(!confirmUncheck(this)) return false;"  Text=""/>
                                    <strong>
                                        <Label for="chkAssignACC" ID="labChkMsg" runat="server" title="" >確認通路代碼與帳號</Label>&nbsp;
                                    </strong>
                                </td>
                            </tr>
                            <tr>
                                <td style="background-color: #99ccff;" colspan="10" class="auto-style2">
                                    <strong>
                                        <asp:Label ID="Label2" runat="server" Text="[經辦代號]"></asp:Label>
                                    </strong>
                                    <asp:Label ID="labIOFFICER" runat="server"></asp:Label>&nbsp;
                                    <asp:Label ID="labNIOfficer" runat="server" Text =""></asp:Label>&nbsp;&nbsp;
                                    <asp:Label ID="Label16" ForeColor="Red" runat="server" Font-Bold="True" Font-Size="Small" Text="「國內地區」僅受理投保人數2人(含)以上要保作業！"></asp:Label>
                                </td>
                            </tr>
                            <tr class="b2b-section-heading">
                                <td style="background-color: #99ccff; height: 30px" colspan="5">
                                    <p align="left">
                                        <b>[旅遊資訊]</b>
                                        <asp:TextBox Style="border-bottom-style: none; border-right-style: none; background-color: #99ccff; border-top-style: none; border-left-style: none" ID="CheckSNNN" runat="server" ReadOnly="True" BorderStyle="None" ForeColor="Red" BorderColor="Transparent"></asp:TextBox>
                                        </b>
                                    </p>
                                </td>
                                <td style="background-color: #99ccff; color: red; height: 30px" colspan="3"/>
                                <td style="background-color: #99ccff; height: 30px" colspan="2">
                                    <asp:Label runat="server" Style="display: none"> <a href="../HAS/img/ITP.pdf" target="_blank"><b>旅綜險空白要保書下載</b></a> <a href="../HAS/img/ITPlist.xls" target="_blank"><b>旅綜險空白名冊下載</b></a> <a href="../HAS/img/IT_KYC.pdf" target="_blank"><b>旅綜險空白業報書下載</b></a> </asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td style="background-color: #99ccff; color: red; height: 30px" colspan="10">
                                    <asp:Label ID="b2b_title" runat="server" Font-Bold="True" Font-Size="Large"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="tdDataCaption">保險期間</td>
                                <td class="TDInput" colspan="9">
                                    <uc1:ucRocDate id="UcRocDobjbgn" runat="server" ISROCDATE="TRUE"></uc1:ucRocDate>
                                    <asp:TextBox ID="txtTobjbgn" runat="server" MaxLength="2" Width="20"></asp:TextBox>時 
            起 ~ 至
                                    <uc1:ucRocDate id="UcRocDobjend" runat="server" ISROCDATE="TRUE"></uc1:ucRocDate>
                                    <asp:TextBox ID="txtTobjend" runat="server" MaxLength="2" Width="20"></asp:TextBox>時 
            <font color="red">(時間區間00~23)</font>共
                                    <asp:TextBox ID="txtIQPITP_Days" onfocus="this.blur()" runat="server" Width="30" MaxLength="3" BorderStyle="None" CssClass="tdDataField"></asp:TextBox>天(國內最長15日；國外最長60日) 
                                    <asp:CheckBox Style="display: none" ID="CK1" runat="server" Checked="True" AutoPostBack="True" Enabled="False" Text="確認後變更保險期間使用"></asp:CheckBox></td>
                            </tr>
                            <tr id="TRAVEL_hidden" runat="server">
                                <td class="tdDataCaption">旅遊地區</td>
                                <td class="tdDataField" colspan="3">
                                    <asp:DropDownList ID="TRAVEL_A" runat="server">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value="400">400 台灣地區及離島地區</asp:ListItem>
                                        <asp:ListItem Value="500">500 大陸地區(包含港、澳)</asp:ListItem>
                                        <asp:ListItem Value="600">600 美加地區</asp:ListItem>
                                        <asp:ListItem Value="710">710 東南亞、南亞地區</asp:ListItem>
                                        <asp:ListItem Value="720">720 日本地區</asp:ListItem>
                                        <asp:ListItem Value="725">725 韓國地區</asp:ListItem>
                                        <asp:ListItem Value="730">730 紐澳地區</asp:ListItem>
                                        <asp:ListItem Value="740">740 歐洲地區</asp:ListItem>
                                        <asp:ListItem Value="750">750 申根地區</asp:ListItem>
                                        <asp:ListItem Value="760">760 非洲地區</asp:ListItem>
                                        <asp:ListItem Value="770">770 中東地區</asp:ListItem>
                                        <asp:ListItem Value="800">800 全球各地</asp:ListItem>
                                        <asp:ListItem Value="810">810 東南亞地區</asp:ListItem>
                                    </asp:DropDownList></td>
                                <td class="tdDataCaption">
                                    <p align="center">旅遊地點</p>
                                </td>
                                <td class="tdDataField" colspan="5"><font
                                    color="#000099"><font color="#000000">啟程地</font></font>
                                    <asp:TextBox ID="TRAVEL_BEGIN" runat="server" MaxLength="10" Width="100"></asp:TextBox>目的地 
                                    <asp:TextBox ID="TRAVEL_END" runat="server" MaxLength="10" Width="100"></asp:TextBox></td>
                            </tr>
                            <tr id="ShowS1" runat="server">
                                <td class="auto-style10">旅遊國家</td>
                                <td class="auto-style11" colspan="2"><font face="新細明體">旅遊地點1</font>
                                    <br>
                                    <asp:DropDownList ID="ddlCountry_A1" runat="server" AutoPostBack="True">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value="1">亞洲</asp:ListItem>
                                        <asp:ListItem Value="2">歐洲</asp:ListItem>
                                        <asp:ListItem Value="3">非洲</asp:ListItem>
                                        <asp:ListItem Value="4">美洲</asp:ListItem>
                                        <asp:ListItem Value="5">大洋洲</asp:ListItem>
                                    </asp:DropDownList><asp:DropDownList ID="ddlCountry_A2" runat="server" Width="140" onchange="nook()"></asp:DropDownList>
                                    <asp:Button ID="btMore" runat="server" CssClass="cssUICommandButton" Text="更多"></asp:Button></FONT>
                                    <asp:Label ID="txtIiofficer" runat="server" Text="Label" Style="display: none"></asp:Label>
                                </td>
                                <td id="td_CountryB" class="tdDataField" colspan="2"><font face="新細明體">旅遊地點2</font>
                                    <br>
                                    <asp:DropDownList ID="ddlCountry_B1" runat="server" AutoPostBack="True">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value="1">亞洲</asp:ListItem>
                                        <asp:ListItem Value="2">歐洲</asp:ListItem>
                                        <asp:ListItem Value="3">非洲</asp:ListItem>
                                        <asp:ListItem Value="4">美洲</asp:ListItem>
                                        <asp:ListItem Value="5">大洋洲</asp:ListItem>
                                    </asp:DropDownList><asp:DropDownList ID="ddlCountry_B2" runat="server" Width="140" onchange="nook()"></asp:DropDownList></FONT></FONT>
                                </td>
                                <td id="td_CountryC" class="tdDataField" colspan="2"><font face="新細明體">旅遊地點3</font>
                                    <br>
                                    <asp:DropDownList ID="ddlCountry_C1" runat="server" AutoPostBack="True">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value="1">亞洲</asp:ListItem>
                                        <asp:ListItem Value="2">歐洲</asp:ListItem>
                                        <asp:ListItem Value="3">非洲</asp:ListItem>
                                        <asp:ListItem Value="4">美洲</asp:ListItem>
                                        <asp:ListItem Value="5">大洋洲</asp:ListItem>
                                    </asp:DropDownList><asp:DropDownList ID="ddlCountry_C2" runat="server" Width="140" onchange="nook()"></asp:DropDownList></FONT>
                                </td>
                                <td id="td_CountryD" class="tdDataField" colspan="2"><font face="新細明體">旅遊地點4</font>
                                    <br>
                                    <asp:DropDownList ID="ddlCountry_D1" runat="server" AutoPostBack="True">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value="1">亞洲</asp:ListItem>
                                        <asp:ListItem Value="2">歐洲</asp:ListItem>
                                        <asp:ListItem Value="3">非洲</asp:ListItem>
                                        <asp:ListItem Value="4">美洲</asp:ListItem>
                                        <asp:ListItem Value="5">大洋洲</asp:ListItem>
                                    </asp:DropDownList><asp:DropDownList ID="ddlCountry_D2" runat="server" Width="140" onchange="nook()"></asp:DropDownList></FONT>
                                </td>
                                <td id="td_CountryBlock" class="tdDataField" colspan ="2"/>
                            </tr>
                            <tr id="ShowS2" runat="server" visible ="false">
                                <td class="tdDataCaption"></td>
                                <td class="tdDataField" colspan="2"><font face="新細明體">旅遊地點5</font>
                                    <br>
                                    <asp:DropDownList ID="ddlCountry_E1" runat="server" AutoPostBack="True">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value="1">亞洲</asp:ListItem>
                                        <asp:ListItem Value="2">歐洲</asp:ListItem>
                                        <asp:ListItem Value="3">非洲</asp:ListItem>
                                        <asp:ListItem Value="4">美洲</asp:ListItem>
                                        <asp:ListItem Value="5">大洋洲</asp:ListItem>
                                    </asp:DropDownList><asp:DropDownList ID="ddlCountry_E2" runat="server" Width="140" onchange="nook()"></asp:DropDownList>
                                </td>                                
                                <td class="tdDataField" colspan="2"><font face="新細明體">旅遊地點6</font>
                                    <br>
                                    <asp:DropDownList ID="ddlCountry_F1" runat="server" AutoPostBack="True">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value="1">亞洲</asp:ListItem>
                                        <asp:ListItem Value="2">歐洲</asp:ListItem>
                                        <asp:ListItem Value="3">非洲</asp:ListItem>
                                        <asp:ListItem Value="4">美洲</asp:ListItem>
                                        <asp:ListItem Value="5">大洋洲</asp:ListItem>
                                    </asp:DropDownList><asp:DropDownList ID="ddlCountry_F2" runat="server" Width="140" onchange="nook()"></asp:DropDownList>
                                </td>                             
                                <td class="tdDataField" colspan="2"><font face="新細明體">旅遊地點7</font>
                                    <br>
                                    <asp:DropDownList ID="ddlCountry_G1" runat="server" AutoPostBack="True">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value="1">亞洲</asp:ListItem>
                                        <asp:ListItem Value="2">歐洲</asp:ListItem>
                                        <asp:ListItem Value="3">非洲</asp:ListItem>
                                        <asp:ListItem Value="4">美洲</asp:ListItem>
                                        <asp:ListItem Value="5">大洋洲</asp:ListItem>
                                    </asp:DropDownList><asp:DropDownList ID="ddlCountry_G2" runat="server" Width="140" onchange="nook()"></asp:DropDownList>
                                </td>                             
                                <td class="tdDataField" colspan="2"><font face="新細明體">旅遊地點8</font>
                                    <br>
                                    <asp:DropDownList ID="ddlCountry_H1" runat="server" AutoPostBack="True">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value="1">亞洲</asp:ListItem>
                                        <asp:ListItem Value="2">歐洲</asp:ListItem>
                                        <asp:ListItem Value="3">非洲</asp:ListItem>
                                        <asp:ListItem Value="4">美洲</asp:ListItem>
                                        <asp:ListItem Value="5">大洋洲</asp:ListItem>
                                    </asp:DropDownList><asp:DropDownList ID="ddlCountry_H2" runat="server" Width="140" onchange="nook()"></asp:DropDownList>
                                </td>
                                <td class="tdDataField" colspan ="2"/>
                            </tr>
                            <tr >
                                <td class="tdDataCaption"></td>
                                <td class="auto-style10">
                                    <div class="button">
                                        <span>日本地區資訊</span>
                                        <div class="hover">
                                            本島 : 東京,大阪,京都,名古屋...<br>
                                            九州 : 福岡,佐賀,熊本,宮崎...<br>
                                            四國 : 高知,德島,香川,愛媛...<br>
                                            沖繩 : 那霸...<br>
                                            北海道: 札幌,函館,千歲,小樽...
                                        </div>
                                    </div>
                                </td>
                                <td class="auto-style10">
                                    <p align="center">
                                        申根公約國<br>
                                        證明文件
                                    </p>
                                </td>
                                <td class="auto-style11">
                                    <asp:RadioButton ID="Checkbox2" onclick="nono()" runat="server" Text="是" GroupName="G750"></asp:RadioButton><font
                                        face="新細明體">&nbsp; </font>
                                    <asp:RadioButton ID="Checkbox4" onclick="nono2()" runat="server" Checked="True" Text="否" GroupName="G750"></asp:RadioButton>
                                </td>
                                <td class="auto-style10"></td>
                                <td class="auto-style11"  colspan="2">
                                    <asp:CheckBox ID="CHK5" onclick="nook()" runat="server" ForeColor="Red" Text="5人(含)以上適用" Font-Size="X-Small" Enabled="False" Style="display: none"></asp:CheckBox>
                                    <asp:Label ID="CHK5_G" runat="server" ForeColor="Red" Style="display: none"></asp:Label>
                                </td>
                                <td class="auto-style11" colspan="2">
                                    <asp:TextBox ID="txtFlightNum" runat="server" MaxLength="30" Width="100" Enabled="False" Visible="false"></asp:TextBox>
                                    <asp:Label ID="lblFlightNote" runat="server" Text="Label"><span style="color:red; font-size:x-small"></span></asp:Label>
                                </td>
                                <td class="auto-style11" colspan="2"></td>
                            </tr>



                            <tr class="b2b-section-heading">
                                <td style="background-color: #99ccff; height: 30px" colspan="9">
                                    <p align="left">
                                        <b><b>[</b>要保人資訊]</b><font color="red">(建議要保人須年滿18歲)
                                    </p>
                                    </FONT></td>
                                <td style="background-color: #99ccff">
                                    <asp:Label Style="display: none" ID="CheckS" runat="server" BackColor="Silver"></asp:Label></td>
                            </tr>
                            <tr>
                                <td style="height: 39px" class="tdDataCaption">要保人名稱</td>
                                <td style="height: 39px" class="tdDataField">
                                    <asp:TextBox onblur="TxtToupper('NAME_IINS')" ID="NAME_IINS" runat="server" MaxLength="50" Width="100"></asp:TextBox><br>
                                    <asp:CheckBox ID="CHK_P1" runat="server" ForeColor="Navy" Text="持外國護照請勾選" Font-Size="X-Small" onChange ="chkboxChange1()"></asp:CheckBox></td>
                                <td style="height: 39px" class="tdDataCaption">要保人ID</td>
                                <td style="height: 39px" class="tdDataField">
                                    <asp:TextBox onblur="ChangeMF0()" ID="txtIins_IINS2" runat="server" MaxLength="10" Width="100px" AutoPostBack="true"></asp:TextBox>
                                    <asp:Button ID="btnIINS2" runat="server" Text="法人帶入資料" CausesValidation="False"></asp:Button><br>
									<asp:dropdownlist id="ddlIINS_COUNTRY" runat="server"  style="display:none">
										<asp:ListItem value="" Selected="True">國別</asp:ListItem>
									</asp:dropdownlist>
                                </td>
                                <td style="height: 39px" class="tdDataCaption">生日</td>
                                <td style="height: 39px" class="tdDataField">
                                    <uc1:ucRocDate id="Birtdy_IINS" runat="server" ISROCDATE="TRUE" EnableViewState="False" ></uc1:ucRocDate>
                                    <%--<asp:textbox style="Z-INDEX: 0" id=txtBirtdy_IINS class=datepicker Runat="server" Width="100px"/>(CYYMMDD) --%>
                                    <font face="新細明體">(750101)</font></td>
                                <td style="height: 39px" class="tdDataCaption">電子保單</td>
                                <td style="height: 39px" class="tdDataField">
                                    <asp:DropDownList ID="ddlChkMailYN" runat="server" Enabled="False">
                                        <%--20260923 文彥郵件回饋，鎖死不可選，預設為"是"--%>
                                        <%--<asp:ListItem Value="N">否</asp:ListItem>--%>
                                        <asp:ListItem Value="Y" Selected="True">是</asp:ListItem>
                                    </asp:DropDownList>
                                <td style="height: 39px" class="tdDataCaption">電子郵件</td>
                                <td style="height: 39px" class="tdDataField">
                                    <asp:TextBox ID="IINSMAIL" runat="server" type="email" MaxLength="50" Width="200"></asp:TextBox></FONT></td>
                            </tr>
                            <tr>
                                <td class="tdDataCaption">聯絡地址</td>
                                <td class="tdDataField" colspan="3"><font face="新細明體">
                                    <uc1:ucCmnZipCodeNew id="ucCmnZipCodeNew1" runat="server" targetclientid="IINS_AREA;IINS_ADDR;IINS_ADDR;"></uc1:ucCmnZipCodeNew>
                                    <asp:TextBox ID="IINS_AREA" runat="server" MaxLength="3" Width="50"></asp:TextBox>
                                    <asp:TextBox ID="IINS_ADDR" runat="server" MaxLength="50" Width="250"></asp:TextBox></font>

                                </td>
                                <td class="tdDataCaption">負責人姓名</td>
                                <td class="tdDataField">
                                    <asp:TextBox ID="txtPrincipal" runat="server" MaxLength="20" Width="100"></asp:TextBox></td>
                                <td class="tdDataCaption">聯絡電話</td>
                                <td class="tdDataField">
                                    <asp:TextBox ID="IINS_TEL" runat="server" MaxLength="20" Width="100"></asp:TextBox></td>
                                <td class="tdDataCaption">行動電話</td>
                                <td class="tdDataField">
                                    <asp:TextBox ID="INSMOBILE" runat="server" MaxLength="20" Width="100"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="tdDataCaption" colspan="2"><font size="1"><strong>被保險人是否已投保其他商業實支實付型
					            		<br>
                                                傷害醫療保險或實支實付型醫療保險?</strong> </font>
                                </td>
					            <td class="tdDataField">
                                    <asp:RadioButton ID="rad_MARKX03_Y" runat="server" GroupName="FEP1" Text="是"></asp:RadioButton>
                                    <asp:RadioButton ID="rad_MARKX03_N" runat="server" GroupName="FEP1" Text="否"></asp:RadioButton>

                                </td>
                                <td class="tdDataCaption" colspan="2"><font size="1"><strong>被保險人目前是否<br>受有監護宣告?</strong></font>
                                </td>                                
					            <td class="tdDataField">
                                    <asp:RadioButton ID="rad_MARKX05_Y" runat="server" GroupName="FEP3" Text="是"></asp:RadioButton>
                                    <asp:RadioButton ID="rad_MARKX05_N" runat="server" GroupName="FEP3" Text="否" Checked="True"></asp:RadioButton>
                                    <asp:RadioButton Visible="false" ID="rad_MARKX04_Y" runat="server" GroupName="FEP2" Text="是"></asp:RadioButton>
                                    <asp:RadioButton Visible="false" ID="rad_MARKX04_N" runat="server" GroupName="FEP2" Text="否" Checked="True"></asp:RadioButton>
                                </td>
                                <td class="tdDataCaption" colspan="2"><font size="1"><strong>
                                    <asp:Label ID="flight_title" runat="server" Text="去程或回程航班(含轉機)是否搭乘長榮航空"></asp:Label></strong>
                                </td>
                                <td class="tdDataField" colspan="2">
                                    <asp:RadioButton ID="rad_flight_Y" runat="server" GroupName="FEP4" Text="是" AutoPostBack="true"></asp:RadioButton>
                                    <asp:RadioButton ID="rad_flight_N" runat="server" GroupName="FEP4" Text="否" AutoPostBack="true"></asp:RadioButton>
                                </td>
                            </tr>
                            <tr>
                                <td style="background-color: #99ccff; height: 30px" colspan="10">
                                    <p align="left">
                                        <b>[班機延誤快速理賠服務]</b>
                                        <asp:Button ID="btnOpenFlight" runat="server" Text="+" CssClass="cssUICommandButton"></asp:Button>
                                    </p>
                                </td>
                            </tr>
                            <tr id="Show5" runat="server">
                                <td class="tdDataCaption"></td>
                                <td class="tdDataField">
                                    <asp:Button ID="btnGetFlightInfo" runat="server" CssClass="cssUICommandButton" Text="新增航班資訊"></asp:Button>
                                </td>
                                <td class="tdDataCaption" style="text-align:left; color:red" colspan="6" >
                                    注意事項:<br/>
                                    &emsp;●當班機延誤達4小時以上時(不適用在班機抵達延誤、班機取消、機位超賣、改搭非原訂班機、其他查無航班等)，我們將主動寄發簡訊通知您，提供您快速理賠服務。<br/>
                                    &emsp;●如有轉機航班也可輸入。<br/>
                                    &emsp;●服務僅限於9位(含)以下被保險人。<br/>
                                    &emsp;●服務不適用「旅平險 & 國內海陸型」。<br/>
                                    &emsp;●快速理賠服務限年滿18歲之被保險人使用；未滿18歲者，請以紙本方式申請理賠，並由法定代理人於理賠申請書簽名或蓋章以示同意。  <br/>
                                    &emsp;●若觸發快速理賠時，將寄發簡訊至每位被保人所填之行動電話(不含未滿18歲之被保險人)。  <br/>
                                </td>
                            </tr>
                            <tr id="Show6" runat="server">
                                <td colspan="10">
                                    <uc1:datagridbanner id="Datagridbanner3" runat="server" NewItemEnable="false" NewItemCaption="" DataGridID="dgdFlight" Visible="False"></uc1:datagridbanner>
                                    <asp:DataGrid ID="dgdFlight" CssClass="cssDataGrid" AllowPaging="True" HorizontalAlign="Center"
                                        AllowSorting="True" runat="server" AutoGenerateColumns="False" Width="100%" Font-Names="新細明體">
                                        <FooterStyle CssClass="cssDataGridFooter"></FooterStyle>
                                        <SelectedItemStyle CssClass="cssDataGridSelectedItem"></SelectedItemStyle>
                                        <EditItemStyle CssClass="cssDataGridEditItem"></EditItemStyle>
                                        <AlternatingItemStyle CssClass="cssDataGridAlternatingItem"></AlternatingItemStyle>
                                        <ItemStyle CssClass="cssDataGridItem"></ItemStyle>
                                        <HeaderStyle ForeColor="#CCFFFF" CssClass="cssDataGridHeader"></HeaderStyle>
                                        <Columns>
                                            <asp:ButtonColumn Text="刪除" CommandName="DeleteFlight"></asp:ButtonColumn>
                                            <asp:TemplateColumn SortExpression="ISEQ" HeaderText="航程">
                                                <ItemTemplate>
                                                    <asp:Label ID="labISEQ" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ISEQ") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn SortExpression="Voyage" HeaderText="航程" Visible="false">
                                                <ItemTemplate>
                                                    <asp:Label ID="labVoyage" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Voyage") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn SortExpression="Dflight" HeaderText="日期">
                                                <ItemTemplate>
                                                    <asp:Label ID="labDflight" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Dflight", "{0:yyyy-MM-dd}") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn SortExpression="Carrier" HeaderText="航空公司" Visible="false">
                                                <ItemTemplate>
                                                    <asp:Label ID="labCarrier" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Carrier") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn SortExpression="Flight_IATA" HeaderText="航班代號">
                                                <ItemTemplate>
                                                    <asp:Label ID="labFlight_IATA" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Flight_IATA") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn> 
                                            <asp:TemplateColumn SortExpression="Flightnumber" HeaderText="航班編號">
                                                <ItemTemplate>
                                                    <asp:Label ID="labFlightnumber" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Flightnumber") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn SortExpression="Departureairport" HeaderText="起飛機場" Visible="false">
                                                <ItemTemplate>
                                                    <asp:Label ID="labDepartureairport" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Departureairport") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>  
                                            <asp:TemplateColumn SortExpression="DepartureAirport_IATA" HeaderText="出發機場">
                                                <ItemTemplate>
                                                    <asp:Label ID="labDepartureAirport_IATA" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.DepartureAirport_IATA") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn> 
                                            <asp:TemplateColumn SortExpression="Departuretime" HeaderText="起飛時間" Visible="false">
                                                <ItemTemplate>
                                                    <asp:Label ID="labDeparturetime" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Departuretime") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>    
                                            <asp:TemplateColumn SortExpression="Departuretime_EC" HeaderText="起飛時間(當地時間)">
                                                <ItemTemplate>
                                                    <asp:Label ID="labDeparturetime_EC" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Departuretime_EC", "{0:yyyy-MM-dd HH:mm}") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>   
                                            <asp:TemplateColumn SortExpression="Arrivalairport" HeaderText="抵達機場" Visible="false">
                                                <ItemTemplate>
                                                    <asp:Label ID="labArrivalairport" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Arrivalairport") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>    
                                            <asp:TemplateColumn SortExpression="ArrivalAirport_IATA" HeaderText="到達機場">
                                                <ItemTemplate>
                                                    <asp:Label ID="labArrivalAirport_IATA" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ArrivalAirport_IATA") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn> 
                                            <asp:TemplateColumn SortExpression="Arrivaltime" HeaderText="到站時間" Visible="false">
                                                <ItemTemplate>
                                                    <asp:Label ID="labArrivaltime" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Arrivaltime") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>      
                                            <asp:TemplateColumn SortExpression="Arrivaltime_EC" HeaderText="到站時間(當地時間)">
                                                <ItemTemplate>
                                                    <asp:Label ID="labArrivaltime_EC" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Arrivaltime_EC", "{0:yyyy-MM-dd HH:mm}") %>'>
                                                    </asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>                               
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>

                                </td>
                                <td colspan="1"/>
                            </tr>
                            <tr id="Show7" runat="server" height="20">
                                <td colspan="10" />
                            </tr>
                            <tr class="b2b-section-heading">
                                <td style="background-color: #99ccff" colspan="3"><b><b>[</b>被保人資訊]&nbsp;</b><input id="uploadM" value="名冊上傳" type="button">
                                    <a href="Insured_List_ITP.xls" target="_blank" download ><b>範例下載</b></a></td>
                                <td style="background-color: #99ccff" colspan="7">
                                    <p style="display: none" id="p1">
                                        <input id="UploadFile" size="50" type="file"
                                            name="UploadFile" runat="server"><asp:TextBox ID="InsertUP_BTN" runat="server" Width="1" MaxLength="1"></asp:TextBox>
                                        <asp:Button ID="UP_BTN" runat="server" Text="上傳" CausesValidation="False"></asp:Button>
                                    </p>
                                </td>
                                <tr>
                                    <td class="tdDataCaption">被保人名稱<br>
                                        <input style="width: 80px; height: 24px" id="Button0" onclick="ShowName()" value="同要保人" type="button" name="Button0" runat="server"></td>
                                    <td class="tdDataField">
                                        <asp:TextBox onblur="TxtToupper('NAME_IISSU')" ID="NAME_IISSU" runat="server" MaxLength="50" Width="100"></asp:TextBox>
                                        <asp:TextBox ID="txtISEQ" runat="server" ReadOnly="True" BorderStyle="None"></asp:TextBox>
                                        <asp:TextBox ID="txtIISSU_Ori" runat="server" ReadOnly="True" BorderStyle="None"></asp:TextBox><br>
                                        <asp:CheckBox ID="CHK_P2" runat="server" ForeColor="Navy" Text="持外國護照請勾選" Font-Size="X-Small"  onChange="chkboxChange2()"></asp:CheckBox></td>
                                    <td class="tdDataCaption">被保人ID<br>
                                    </td>
                                    <td class="tdDataField">
                                        <asp:TextBox onblur="ChangeMF()" ID="txtIins_ISSU2" runat="server" MaxLength="10" Width="100px"></asp:TextBox><br>
									    <asp:dropdownlist id="ddlIISSU_COUNTRY" runat="server" style="display:none">
									    	<asp:ListItem value="" Selected="True">國別</asp:ListItem>
									    </asp:dropdownlist>
                                    </td>
                                    <td class="tdDataCaption">生日</td>
                                    <td class="tdDataField"><font face="新細明體">
                                        <uc1:ucRocDate id="UcRocDbirth" runat="server" ISROCDATE="TRUE"></uc1:ucRocDate>(750101)<asp:TextBox ID="txtDBIRTH" runat="server" ReadOnly="True" BorderStyle="None" CssClass="tdDataField" Width="20"></asp:TextBox></font></td>
                                    <td style="height: 9px" class="tdDataCaption">被保險人關係<br />為要保人之</td>
                                    <td style="height: 9px" class="tdDataField">
                                        <asp:DropDownList ID="Relationship" runat="server">
                                            <asp:ListItem Selected="True"></asp:ListItem>
                                            <asp:ListItem Value="1">1 本人</asp:ListItem>
                                            <asp:ListItem Value="2">2 配偶</asp:ListItem>
                                            <asp:ListItem Value="3">3 子女</asp:ListItem>
                                            <asp:ListItem Value="4">4 父母</asp:ListItem>
                                            <asp:ListItem Value="5">5 兄弟姐妹</asp:ListItem>
                                            <asp:ListItem Value="6">6 祖父母</asp:ListItem>
                                            <asp:ListItem Value="7">7 孫子女</asp:ListItem>
                                            <asp:ListItem Value="8">8 雇員</asp:ListItem>
                                            <asp:ListItem Value="9">9 員工親屬</asp:ListItem>
                                            <asp:ListItem Value="0">0 其他</asp:ListItem>
                                        </asp:DropDownList><br />
                                        <asp:TextBox ID="txtRelationshipOther" runat="server" MaxLength="5" Width="100" placeholder="選[其他]請說明" Style="display: none;"></asp:TextBox>
                                    </td>
                                    <td style="height: 9px" class="tdDataCaption">電子郵件 </td>
                                    <td style="height: 9px" class="tdDataField"><font
                                        face="新細明體">
                                        <asp:TextBox ID="EMAIL" runat="server" type="email" MaxLength="50" Width="200"></asp:TextBox><asp:Label ID="txtMIISSU" runat="server" ForeColor="White" Font-Size="XX-Small"></asp:Label><asp:Label ID="txtMNIISSU" runat="server" ForeColor="White" Font-Size="XX-Small"></asp:Label></font></td>
                                </tr>
                            <tr>
                                <td style="height: 28px" class="tdDataCaption">聯絡地址<br>
                                    <input style="width: 80px; height: 24px" id="Button2" onclick="ShowName4()" value="同要保人" type="button" name="Button2" runat="server"></td>
                                <td style="height: 28px" class="tdDataField" colspan="3"><font face="新細明體">
                                    <uc1:ucCmnZipCodeNew id="ucCmnZipCodeNew2" runat="server" targetclientid="IISSU_AREA;IISSU_ADDR;IISSU_ADDR;"></uc1:ucCmnZipCodeNew>
                                    <asp:TextBox ID="IISSU_AREA" runat="server" MaxLength="3" Width="50"></asp:TextBox>
                                    <asp:TextBox ID="IISSU_ADDR" runat="server" MaxLength="50" Width="250"></asp:TextBox><span class="auto-style12"><strong></td>

                                <td class="tdDataCaption"><strong>被保險人<br />
                                    是否為孕婦?</strong></td>
                                <td style="height: 9px" class="tdDataField">
                                    <asp:RadioButton ID="isEMP_Y" runat="server" GroupName="FEP5" Text="是" />
                                    <asp:RadioButton ID="isEMP_N" runat="server" GroupName="FEP5" Text="否" Checked="true" />
                                </td>
                                <td style="height: 28px" class="tdDataCaption">性別</td>
                                <td style="height: 9px" class="tdDataField">
                                    <asp:DropDownList ID="ddlSex" runat="server">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value="1">1 男</asp:ListItem>
                                        <asp:ListItem Value="2">2 女</asp:ListItem>
                                    </asp:DropDownList></td>
                                <td style="height: 28px" class="tdDataCaption">行動電話</td>
                                <td style="height: 28px" class="tdDataField">
                                    <asp:TextBox ID="IISSU_TEL" runat="server" MaxLength="20" Width="100"></asp:TextBox></td>
                            </tr>
                            <tr id="eng_show">
                                <td class="tdDataCaption">英文姓名</td>
                                <td class="tdDataField">
                                    <asp:TextBox ID="PASSPORT_NAME" runat="server" Width="100"></asp:TextBox></FONT></td>
                                <td class="tdDataCaption">護照號碼</td>
                                <td class="tdDataField" colspan="5">
                                    <asp:TextBox ID="PASSPORT_NUMBER" runat="server" MaxLength="20" Width="100"></asp:TextBox></td>
                                <td class="tdDataCaption"></td>
                                <td class="tdDataField"><font face="新細明體"></font></td>
                            </tr>
                            <tr class="b2b-section-heading">
                                <td style="background-color: #99ccff; height: 30px" colspan="10"><b><b><b>[</b></b>受益人資訊]</b>
                                    <div style="display:none">
                                        <asp:DropDownList ID="ddl_IINS" runat="server" BackColor="White" Width="80">
                                            <asp:ListItem Selected="True"></asp:ListItem>
                                            <asp:ListItem Value="1">1.身份證號</asp:ListItem>
                                            <asp:ListItem Value="2">2.統一編號</asp:ListItem>
                                            <asp:ListItem Value="3">3.居留證號</asp:ListItem>
                                            <asp:ListItem Value="5">5.護照號碼</asp:ListItem>
                                        </asp:DropDownList><asp:DropDownList ID="ddl_IINS2" runat="server" Width="110" BackColor="White">
                                            <asp:ListItem></asp:ListItem>
                                            <asp:ListItem Value="1">1.自然人本國籍</asp:ListItem>
                                            <asp:ListItem Value="2">2.自然人外國籍</asp:ListItem>
                                            <asp:ListItem Value="3">3.法人、非法人團體機關</asp:ListItem>
                                            <asp:ListItem Value="7">7.自然人外國籍居留證號不符檢核</asp:ListItem>
                                            <asp:ListItem Value="8">8.自然人本國籍身份證號不符檢核</asp:ListItem>
                                            <asp:ListItem Value="9">9.法人、非法人團體機關不符檢核</asp:ListItem>
                                        </asp:DropDownList><asp:DropDownList ID="ddl_IISSU" runat="server" Width="80" BackColor="White">
                                            <asp:ListItem Selected="True"></asp:ListItem>
                                            <asp:ListItem Value="1">1.身份證號</asp:ListItem>
                                            <asp:ListItem Value="3">3.居留證號</asp:ListItem>
                                            <asp:ListItem Value="5">5.護照號碼</asp:ListItem>
                                        </asp:DropDownList><asp:DropDownList ID="ddl_IISSU2" runat="server" Width="100" BackColor="White">
                                            <asp:ListItem></asp:ListItem>
                                            <asp:ListItem Value="1">1.自然人本國籍</asp:ListItem>
                                            <asp:ListItem Value="2">2.自然人外國籍</asp:ListItem>
                                            <asp:ListItem Value="3">3.法人、非法人團體機關</asp:ListItem>
                                            <asp:ListItem Value="7">7.自然人外國籍居留證號不符檢核</asp:ListItem>
                                            <asp:ListItem Value="8">8.自然人本國籍身份證號不符檢核</asp:ListItem>
                                            <asp:ListItem Value="9">9.法人、非法人團體機關不符檢核</asp:ListItem>
                                        </asp:DropDownList>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td style="height: 14px" class="tdDataCaption">受益人名稱</td>
                                <td style="height: 14px" class="tdDataField">
                                    <asp:TextBox ID="Benefi" runat="server" MaxLength="10" Width="100"></asp:TextBox><font
                                        color="#000000"></font></FONT></td>
                                <td style="height: 14px" class="tdDataCaption">身分證號</td>
                                <td style="height: 14px" class="tdDataField" colspan="3">
                                    <asp:TextBox ID="Benefi_ID" runat="server" MaxLength="10" Width="100px"></asp:TextBox></td>
                                <td style="height: 14px" class="tdDataCaption">與被保人關係</td>
                                <td class="tdDataField" colspan="3">
                                    <asp:DropDownList ID="Relationship2" runat="server">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value='1'>1 配偶</asp:ListItem>
                                        <asp:ListItem Value='2'>2 子女</asp:ListItem>
                                        <asp:ListItem Value='3'>3 父母</asp:ListItem>
                                        <asp:ListItem Value='4'>4 兄弟姐妹</asp:ListItem>
                                        <asp:ListItem Value='9'>9 祖父母</asp:ListItem>
                                        <asp:ListItem Value='A'>A 孫子女</asp:ListItem>
                                    </asp:DropDownList>
                                    <asp:Button ID="btMore2" runat="server" CssClass="cssUICommandButton" Text="更多"></asp:Button>
                                </td>
                            </tr>
                            <tr id="ShowS22" runat="server">
                                <td style="height: 14px" class="tdDataCaption">受益人名稱2</td>
                                <td style="height: 14px" class="tdDataField">
                                    <asp:TextBox ID="Benefi2" runat="server" MaxLength="10" Width="100"></asp:TextBox><font
                                        color="#000000"></font></FONT></td>
                                <td style="height: 14px" class="tdDataCaption">身分證號2</td>
                                <td style="height: 14px" class="tdDataField" colspan="3">
                                    <asp:TextBox ID="Benefi_ID2" runat="server" MaxLength="10" Width="100px"></asp:TextBox></td>
                                <td style="height: 14px" class="tdDataCaption">與被保人關係2</td>
                                <td class="tdDataField">
                                    <asp:DropDownList ID="Relationship22" runat="server">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value='1'>1 配偶</asp:ListItem>
                                        <asp:ListItem Value='2'>2 子女</asp:ListItem>
                                        <asp:ListItem Value='3'>3 父母</asp:ListItem>
                                        <asp:ListItem Value='4'>4 兄弟姐妹</asp:ListItem>
                                        <asp:ListItem Value='9'>9 祖父母</asp:ListItem>
                                        <asp:ListItem Value='A'>A 孫子女</asp:ListItem>
                                    </asp:DropDownList></td>
                                <td style="height: 14px" class="tdDataCaption">給付方式</td>
                                <td class="tdDataField">
                                    <asp:DropDownList ID="ddlRELMOTHED" runat="server">
                                        <asp:ListItem Value='1'>1 均分</asp:ListItem>
                                        <asp:ListItem Value='2' Selected="True">2 順位</asp:ListItem>
                                    </asp:DropDownList></td>

                            </tr>
                            <tr id="ShowS23" runat="server" visible ="false">
                                <td style="height: 14px" class="tdDataCaption">受益人名稱3</td>
                                <td style="height: 14px" class="tdDataField">
                                    <asp:TextBox ID="Benefi3" runat="server" MaxLength="10" Width="100"></asp:TextBox><font
                                        color="#000000"></font></FONT></td>
                                <td style="height: 14px" class="tdDataCaption">身分證號3</td>
                                <td style="height: 14px" class="tdDataField" colspan="3">
                                    <asp:TextBox ID="Benefi_ID3" runat="server" MaxLength="10" Width="100px"></asp:TextBox></td>
                                <td style="height: 14px" class="tdDataCaption">與被保人關係3</td>
                                <td class="tdDataField">
                                    <asp:DropDownList ID="Relationship23" runat="server">
                                        <asp:ListItem Selected="True"></asp:ListItem>
                                        <asp:ListItem Value='1'>1 配偶</asp:ListItem>
                                        <asp:ListItem Value='2'>2 子女</asp:ListItem>
                                        <asp:ListItem Value='3'>3 父母</asp:ListItem>
                                        <asp:ListItem Value='4'>4 兄弟姐妹</asp:ListItem>
                                        <asp:ListItem Value='9'>9 祖父母</asp:ListItem>
                                        <asp:ListItem Value='A'>A 孫子女</asp:ListItem>
                                    </asp:DropDownList></td>

                            </tr>
                            <tr>
                                <td class="tdDataCaption">聯絡地址<br>
                                    <input style="width: 80px; height: 24px" id="Button3" onclick="ShowName3()" value="同要保人" type="button" name="Button3" runat="server"></td>
                                <td class="tdDataField" colspan="5"><font face="新細明體">
                                    <uc1:ucCmnZipCodeNew id="ucCmnZipCodeNew3" runat="server" targetclientid="txtZIPREL;txtAREL;txtAREL;"></uc1:ucCmnZipCodeNew>
                                    <asp:TextBox ID="txtZIPREL" runat="server" MaxLength="3" Width="50"></asp:TextBox>
                                    <asp:TextBox ID="txtAREL" runat="server" MaxLength="50" Width="250"></asp:TextBox></td>
                                <td class="tdDataCaption">聯絡電話</td>
                                <td class="tdDataField" colspan="3">
                                    <asp:TextBox ID="txtTELREL" runat="server" MaxLength="20" Width="100"></asp:TextBox>
                                    <asp:TextBox ID="Check750" runat="server" ReadOnly="True" Style="display: none"></asp:TextBox>
                                    <asp:TextBox ID="PassportShow" Style="display: none" runat="server" ReadOnly="True"></asp:TextBox>
                                </td>
                            </tr>


                        </tbody>
                    </table>
                    <table class="tblDataArea" width="100%">
                        <tbody>
                            <tr>
                                <td style="background-color: #99ccff; height: 30px" colspan="5">
                                    <p align="left">
                                        <b><b><b>[</b></b>旅平險保障內容]</b>
                                    </p>
                                </td>
                                <td style="background-color: #99ccff; height: 30px" colspan="10">
                                    <p align="left">
                                        <%--<b><b><b></b></b>本次旅遊海外突發疾病是否含傳染病?</b> --%>
                                        <%--                                        <input type="radio" id="rad_Disease_Y" runat="server" name="rad_Disease" value="optionY" onclick="radDiseaseClick(this.value)" />是
                                        <input type="radio" id="rad_Disease_N" runat="server" name="rad_Disease" value="optionN" onclick="radDiseaseClick(this.value)"  checked/>否--%>
                                        <asp:RadioButton ID="rad_Disease_Y" runat="server" GroupName="rad_Disease" Text="是" Style="Display:none" onclick="yesit63()" ></asp:RadioButton>
                                        <asp:RadioButton ID="rad_Disease_N" runat="server" GroupName="rad_Disease" Text="否" Style="Display:none" onclick="noit63()" Checked="True"></asp:RadioButton>
                                    </p>
                                </td>
                            </tr>
                            <tr>
                                <td class="tdDataCaption">
                                    <p align="center">
                                        <strong>旅行平安保險<br>
                                            <strong>承保項目</strong>
                                    </p>
                                    </STRONG>
                                </td>
                                <td class="tdDataCaption">
                                    <p align="center">
                                        意外身故<br>
                                        失能保險金
                                    </p>
                                </td>
                                <td class="tdDataField"><span class="style8"></span>
                                    <asp:DropDownList ID="ddlIT41" runat="server" autopostback="true">
                                        <asp:ListItem Value=""></asp:ListItem>
                                        <asp:ListItem Value="0">0</asp:ListItem>
                                        <asp:ListItem Value="50">50</asp:ListItem>
                                        <asp:ListItem Value="60">60</asp:ListItem>
                                        <asp:ListItem Value="100">100</asp:ListItem>
                                        <asp:ListItem Value="200">200</asp:ListItem>
                                        <asp:ListItem Value="300">300</asp:ListItem>
                                        <asp:ListItem Value="400">400</asp:ListItem>
                                        <asp:ListItem Value="500">500</asp:ListItem>
                                        <asp:ListItem Value="600">600</asp:ListItem>
                                        <asp:ListItem Value="800">800</asp:ListItem>
                                        <asp:ListItem Value="900">900</asp:ListItem>
                                        <asp:ListItem Value="1000">1000</asp:ListItem>
                                        <asp:ListItem Value="1200">1200</asp:ListItem>
                                        <asp:ListItem Value="1500">1500</asp:ListItem>
                                    </asp:DropDownList>萬
                                    <asp:TextBox  ID="txtIT41" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off" style="display:none"></asp:TextBox>
                                </td>
                                <td class="tdDataCaption">
                                    <p align="center">
                                        傷害醫療<br>
                                        保險金
                                    </p>
                                </td>
                                <td class="tdDataField">
                                    <asp:DropDownList ID="ddlIT42" runat="server">
                                        <asp:ListItem></asp:ListItem>
                                        <asp:ListItem Value="0">0</asp:ListItem>
                                        <asp:ListItem Value="1">1</asp:ListItem>
                                        <asp:ListItem Value="3">3</asp:ListItem>
                                        <asp:ListItem Value="5">5</asp:ListItem>
                                        <asp:ListItem Value="10">10</asp:ListItem>
                                        <asp:ListItem Value="20">20</asp:ListItem>
                                        <asp:ListItem Value="30">30</asp:ListItem>
                                        <asp:ListItem Value="40">40</asp:ListItem>
                                        <asp:ListItem Value="50">50</asp:ListItem>
                                        <asp:ListItem Value="60">60</asp:ListItem>
                                        <asp:ListItem Value="80">80</asp:ListItem>
                                        <asp:ListItem Value="90">90</asp:ListItem>
                                        <asp:ListItem Value="100">100</asp:ListItem>
                                        <asp:ListItem Value="120">120</asp:ListItem>
                                        <asp:ListItem Value="150">150</asp:ListItem>
                                    </asp:DropDownList>萬

                                    <asp:TextBox  ID="txtIT42" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off" style="display:none"></asp:TextBox>
                                </td>
                                <td class="tdDataCaption">
                                    <p align="center">
                                        海外突發疾病<br>
                                        保險金(不含傳染病)
                                    </p>
                                </td>
                                <td class="tdDataField">
                                    <asp:DropDownList ID="ddlIT61" runat="server">
                                        <asp:ListItem></asp:ListItem>
                                        <asp:ListItem Value="1">1</asp:ListItem>
                                        <asp:ListItem Value="3">3</asp:ListItem>
                                        <asp:ListItem Value="5">5</asp:ListItem>
                                        <asp:ListItem Value="10">10</asp:ListItem>
                                        <asp:ListItem Value="20">20</asp:ListItem>
                                        <asp:ListItem Value="30">30</asp:ListItem>
                                        <asp:ListItem Value="40">40</asp:ListItem>
                                        <asp:ListItem Value="50">50</asp:ListItem>
                                        <asp:ListItem Value="60">60</asp:ListItem>
                                        <asp:ListItem Value="80">80</asp:ListItem>
                                        <asp:ListItem Value="90">90</asp:ListItem>
                                        <asp:ListItem Value="100">100</asp:ListItem>
                                    </asp:DropDownList>萬
                                    <asp:TextBox  ID="txtIT61" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off" style="display:none"></asp:TextBox>
                                </td>
                                <td class="tdDataCaption">
                                    <p align="center">
                                        重大燒燙傷<br>
                                        保險金
                                    </p>
                                </td>
                                <td class="tdDataField"><span class="style8"></span>
                                    <asp:DropDownList ID="ddlIT44" runat="server">
                                        <asp:ListItem></asp:ListItem>
                                        <asp:ListItem Value="0">0</asp:ListItem>
                                        <asp:ListItem Value="1">1</asp:ListItem>
                                        <asp:ListItem Value="2">2</asp:ListItem>
                                        <asp:ListItem Value="3">3</asp:ListItem>
                                        <asp:ListItem Value="4">4</asp:ListItem>
                                        <asp:ListItem Value="5">5</asp:ListItem>
                                        <asp:ListItem Value="6">6</asp:ListItem>
                                        <asp:ListItem Value="7">7</asp:ListItem>
                                        <asp:ListItem Value="8">8</asp:ListItem>
                                        <asp:ListItem Value="9">9</asp:ListItem>
                                        <asp:ListItem Value="10">10</asp:ListItem>
                                        <asp:ListItem Value="11">11</asp:ListItem>
                                        <asp:ListItem Value="12">12</asp:ListItem>
                                        <asp:ListItem Value="13">13</asp:ListItem>
                                        <asp:ListItem Value="14">14</asp:ListItem>
                                        <asp:ListItem Value="15">15</asp:ListItem>
                                        <asp:ListItem Value="16">16</asp:ListItem>
                                        <asp:ListItem Value="17">17</asp:ListItem>
                                        <asp:ListItem Value="18">18</asp:ListItem>
                                        <asp:ListItem Value="19">19</asp:ListItem>
                                        <asp:ListItem Value="20">20</asp:ListItem>
                                        <asp:ListItem Value="25">25</asp:ListItem>
                                        <asp:ListItem Value="30">30</asp:ListItem>
                                        <asp:ListItem Value="40">40</asp:ListItem>
                                        <asp:ListItem Value="50">50</asp:ListItem>
                                        <asp:ListItem Value="60">60</asp:ListItem>
                                        <asp:ListItem Value="70">70</asp:ListItem>
                                        <asp:ListItem Value="75">75</asp:ListItem>
                                        <asp:ListItem Value="80">80</asp:ListItem>
                                        <asp:ListItem Value="90">90</asp:ListItem>
                                        <asp:ListItem Value="100">100</asp:ListItem>
                                        <asp:ListItem Value="125">125</asp:ListItem>
                                        <asp:ListItem Value="150">150</asp:ListItem>
                                        <asp:ListItem Value="175">175</asp:ListItem>
                                        <asp:ListItem Value="200">200</asp:ListItem>
                                        <asp:ListItem Value="225">225</asp:ListItem>
                                        <asp:ListItem Value="250">250</asp:ListItem>
                                    </asp:DropDownList>萬
                                    <asp:TextBox  ID="txtIT44" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off" style="display:none"></asp:TextBox>
                                </td>
                                <td class="tdDataCaption">
                                    <p align="center" id ="lblsum01" runat ="server">
                                        <strong><font size="2">保險費</font></strong>
                                    </p>
                                </td>
                            </tr>
                            <tr>
                                <td class="tdDataCaption">
                                    <p align="center">
                                        <font face="新細明體"></font>
                                    </p>
                                </td>
                                <td class="tdDataCaption">
                                    <p align="center">
                                        意外喪葬費用或<br>
                                        意外失能保險金<br>
                                        (未滿15足歲)
                                    </p>
                                </td>
                                <td class="tdDataField"><span class="style8">
                                    <asp:DropDownList ID="ddlIT4Y" runat="server"  autopostback="true">
                                        <asp:ListItem></asp:ListItem>
                                        <asp:ListItem Value="0">0</asp:ListItem>
                                        <asp:ListItem Value="10">10</asp:ListItem>
                                        <asp:ListItem Value="20">20</asp:ListItem>
                                        <asp:ListItem Value="30">30</asp:ListItem>
                                        <asp:ListItem Value="40">40</asp:ListItem>
                                        <asp:ListItem Value="50">50</asp:ListItem>
                                        <asp:ListItem Value="69">69</asp:ListItem>
                                    </asp:DropDownList>萬</span>
                                    <asp:TextBox  ID="txtIT4Y" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off" style="display:none"></asp:TextBox>
                                </td>

                                <asp:TextBox Style="display: none" ID="TextBox1" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField"></asp:TextBox>
                </td>
                <td class="tdDataCaption">
                    <p align="center">
                        搭乘大眾運輸意外<br />
                        保險金(身故20%)
                    </p>
                </td>
                <td class="tdDataField">
                    <!--ddlIT43下拉內容對應身故保險金的20%-->
                    <asp:DropDownList ID="ddlIT43" runat="server" Enabled ="false">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem Value="0">0</asp:ListItem>
                        <asp:ListItem Value="12">12</asp:ListItem>
                        <asp:ListItem Value="20">20</asp:ListItem>
                        <asp:ListItem Value="40">40</asp:ListItem>
                        <asp:ListItem Value="60">60</asp:ListItem>
                        <asp:ListItem Value="80">80</asp:ListItem>
                        <asp:ListItem Value="100">100</asp:ListItem>
                        <asp:ListItem Value="120">120</asp:ListItem>
                        <asp:ListItem Value="160">160</asp:ListItem>
                        <asp:ListItem Value="180">180</asp:ListItem>
                        <asp:ListItem Value="200">200</asp:ListItem>
                        <asp:ListItem Value="240">240</asp:ListItem>
                        <asp:ListItem Value="300">300</asp:ListItem>
                    </asp:DropDownList>萬
                                                        
                    <asp:TextBox  ID="txtIT43" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off" style="display:none"></asp:TextBox>
                </td>
                <td class="tdDataCaption" colspan="1">
                    <p align="center" id="titIT63" runat="server">
                        海外突發疾病<br>
                        保險金(含傳染病)
                    </p>
                    <p align="center" id="titIT64" runat="server">
                        海外突發疾病<br>
                        住院醫療保險金
                    </p>
                </td>
                <td class="tdDataField" colspan="1">
                    <asp:DropDownList ID="ddlIT63" runat="server" BackColor="#ffe5cc">
                        <asp:ListItem Value=""></asp:ListItem>
                        <asp:ListItem Value="1">1</asp:ListItem>
                        <asp:ListItem Value="3">3</asp:ListItem>
                        <asp:ListItem Value="5">5</asp:ListItem>
                        <asp:ListItem Value="10">10</asp:ListItem>
                        <asp:ListItem Value="20">20</asp:ListItem>
                        <asp:ListItem Value="30">30</asp:ListItem>
                        <asp:ListItem Value="40">40</asp:ListItem>
                        <asp:ListItem Value="50">50</asp:ListItem>
                        <asp:ListItem Value="60">60</asp:ListItem>
                        <asp:ListItem Value="80">80</asp:ListItem>
                        <asp:ListItem Value="90">90</asp:ListItem>
                        <asp:ListItem Value="100">100</asp:ListItem>
                        <asp:ListItem Value="120">120</asp:ListItem>
                        <asp:ListItem Value="150">150</asp:ListItem>
                    </asp:DropDownList>
                    
                    <!--ddlIT64內容與ddlIT63相同-->
                    <asp:DropDownList ID="ddlIT64" runat="server" BackColor="#ffffcc">
                        <asp:ListItem Value=""></asp:ListItem>
                        <asp:ListItem Value="1">1</asp:ListItem>
                        <asp:ListItem Value="3">3</asp:ListItem>
                        <asp:ListItem Value="5">5</asp:ListItem>
                        <asp:ListItem Value="10">10</asp:ListItem>
                        <asp:ListItem Value="20">20</asp:ListItem>
                        <asp:ListItem Value="30">30</asp:ListItem>
                        <asp:ListItem Value="40">40</asp:ListItem>
                        <asp:ListItem Value="50">50</asp:ListItem>
                        <asp:ListItem Value="60">60</asp:ListItem>
                        <asp:ListItem Value="80">80</asp:ListItem>
                        <asp:ListItem Value="90">90</asp:ListItem>
                        <asp:ListItem Value="100">100</asp:ListItem>
                        <asp:ListItem Value="120">120</asp:ListItem>
                        <asp:ListItem Value="150">150</asp:ListItem>
                    </asp:DropDownList>
                    <!--K077專屬 _ ddlIT6E內容與ddlIT63相同..-->
                    <asp:DropDownList ID="ddlIT6E" runat="server">
                        <asp:ListItem Value=""></asp:ListItem>
                        <asp:ListItem Value="1">1</asp:ListItem>
                        <asp:ListItem Value="3">3</asp:ListItem>
                        <asp:ListItem Value="5">5</asp:ListItem>
                        <asp:ListItem Value="10">10</asp:ListItem>
                        <asp:ListItem Value="20">20</asp:ListItem>
                        <asp:ListItem Value="30">30</asp:ListItem>
                        <asp:ListItem Value="40">40</asp:ListItem>
                        <asp:ListItem Value="50">50</asp:ListItem>
                        <asp:ListItem Value="60">60</asp:ListItem>
                        <asp:ListItem Value="80">80</asp:ListItem>
                        <asp:ListItem Value="90">90</asp:ListItem>
                        <asp:ListItem Value="100">100</asp:ListItem>
                        <asp:ListItem Value="120">120</asp:ListItem>
                        <asp:ListItem Value="150">150</asp:ListItem>
                    </asp:DropDownList>
                    萬<br/>

                    <a align="left" id="titIT65" runat="server">
                        門診醫療保險金：
                    </a>
                    <a align="left" id="titIT6G" runat="server" style="display:none">
                        門診醫療保險金：
                    </a>
                    <asp:label ID ="lblIT65" runat ="server"></asp:label>
                    <asp:label ID ="lblIT6G" runat ="server" style="display:none" ></asp:label>
                    <a align="left">萬</a><br/>

                    <a align="left" id="titIT66" runat="server">
                        急診醫療保險金：
                    </a>
                    <a align="left" id="titIT6F" runat="server" style="display:none" >
                        急診醫療保險金：
                    </a>
                    <asp:label ID ="lblIT66" runat ="server"></asp:label>
                    <asp:label ID ="lblIT6F" runat ="server" style="display:none" ></asp:label>
                    <a align="left">萬</a><br/>

                    <a align="left" id="titIT6D" runat="server">
                        返國繼續住院醫療保險金：
                    </a>
                    <asp:label ID ="lblIT6D" runat ="server"></asp:label><a align="left" id ="lblUnitIT6D" runat="server">萬</a>
                    
                    <asp:Label ID ="lblIT64" runat="server"></asp:Label>
                    <asp:TextBox  ID="txtIT63" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off"  style="display:none"></asp:TextBox>
                    <asp:TextBox  ID="txtIT64" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off"  style="display:none"></asp:TextBox>
                    <asp:TextBox  ID="txtIT65" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off"  style="display:none"></asp:TextBox>
                    <asp:TextBox  ID="txtIT66" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off"  style="display:none"></asp:TextBox>
                    <asp:TextBox  ID="txtIT6D" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off"  style="display:none"></asp:TextBox>
                    <asp:TextBox  ID="txtIT6E" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off"  style="display:none"></asp:TextBox>
                    <asp:TextBox  ID="txtIT6F" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off"  style="display:none"></asp:TextBox>
                    <asp:TextBox  ID="txtIT6G" runat="server" Width="30px" BorderStyle="None" CssClass="tdDataField" autocomplete="off"  style="display:none"></asp:TextBox>
                </td>
                <td class="tdDataField" colspan="2"></td>
                <td class="tdDataField"><font face="新細明體">
                    <p align="right">
                        <asp:TextBox ID="sum01" runat="server" Width="50" BorderStyle="None" ForeColor="Blue" CssClass="tdDataField" Font-Bold="True"></asp:TextBox>
                    </p>
                </font></td>
                </td>
            </tr>
        </table>
                    <table  id="divShow" runat="server" class="tblDataArea" width="100%">
                        <tbody>
                            <tr>
                                <td style="background-color: #99ccff; height: 20px" colspan="10">
                                    <p align="left">
                                        <b><b><b>[</b></b>旅綜險方案選擇]</b><font face="新細明體">&nbsp;&nbsp;</font><asp:CheckBox Style="display: none" ID="CK2" runat="server" Checked="false" AutoPostBack="True" Text=""></asp:CheckBox><strong>方案需與主被保險人相同，修改方案請選取主被保險人進行修改</strong>
                                    </p>
                                </td>
                            </tr>
                            <tr id="divfunc" runat="server">
                                <td style="background-color: #99ccff; height: 30px" colspan="10"
                                    align="center"><b></b>
                                    <asp:Button ID="BT1" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="A.國內海陸型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BT2" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Style="display: none" Text="B.國內飛翔型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BT3" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="C.海外基本型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BT4" Style="display: none" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="X.海外自駕型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BT5" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="D.海外超值型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BT6" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="J.海外標準型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BT7" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="K.海外豪華型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BT8" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="L.一般型或申根型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BT9" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="M. 海外安心型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BT10" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="N. 海外特選型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BT11" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="O. 海外小資型"></asp:Button><strong></strong>&nbsp;
                                    <asp:Button ID="BTIQP" runat="server" Width="195px" CssClass="cssUICommandButton" OnClientClick="return doBeforeSubmit()" Text="僅人身保險"></asp:Button>                  
                                &nbsp;&nbsp;</td>
                            </tr>
                            <tr id="divfuncinfo" runat="server">
                                <td class="tdDataCaption">
                                    <p align="center" class="auto-style7">
                                        <strong>個人旅行綜合險
                                    </p>
                                    </STRONG>
                                </td>
                                <td class="tdDataCaption">
                                    <p align="center">
                                        <font size="3"><strong><font size="3">[</font></strong><asp:TextBox ID="txtCase" runat="server" ReadOnly="True" BorderStyle="None" ForeColor="Blue" Font-Size="Small" BackColor="Transparent" Font-Bold="True" CssClass="auto-style7 txt150" Style="text-align:center"  Width="150px"></asp:TextBox><strong>]</strong></font>
                                    </p>
                                    </STRONG>
                                </td>
                                <td class="tdDataCaption">
                                    <p align="center" id="lblsum02" runat="server">
                                        <strong><span class="auto-style9">保險費</span></strong>
                                    </p>
                                </td>
                                <td class="tdDataCaptionCenter">
                                    <strong><font size="3">
                                        <asp:TextBox ID="sum02" runat="server" Width="50px" ForeColor="Blue" CssClass="auto-style8" Font-Bold="True"></asp:TextBox>
                                    </font></strong>
                                </td>
                                <td class="tdDataCaption">
                                    <p class="auto-style6">
                                        <strong>
                                            <span  id ="lblTotal" runat ="server" class="auto-style3">合計：</span></strong>
                                        <asp:TextBox ID="txtTotal" runat="server" ReadOnly="True" Width="50" BorderStyle="None" ForeColor="Blue" CssClass="tdDataField" Font-Size="Larger" Font-Bold="True"></asp:TextBox>
                                    </p>
                                    </font></strong></font></strong>
                                </td>
                            </tr>
                    </table>
        <table class="tblDataArea" width="100%"  style="display:none">
            <tr>

                <td class="tdDataCaption">
                    <p align="center">
                        IT11
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT11" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT3A
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT3A" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT38
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT38" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT2I
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT2I" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT27
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT27" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
            </tr>
            <tr>

                <td class="auto-style4">
                    <p align="center">
                        IT28
                    </p>
                </td>
                <td class="auto-style5">
                    <asp:TextBox ID="txtIT28" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="auto-style4">
                    <p align="center">
                        IT2E
                    </p>
                </td>
                <td class="auto-style5">
                    <asp:TextBox ID="txtIT2E" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="auto-style4">
                    <p align="center">
                        IT2G
                    </p>
                </td>
                <td class="auto-style5">
                    <asp:TextBox ID="txtIT2G" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>

                <td class="auto-style4">
                    <p align="center">
                        IT3B
                    </p>
                </td>
                <td class="auto-style5">
                    <asp:TextBox ID="txtIT3B" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="auto-style4">
                    <p align="center">
                        IT39
                    </p>
                </td>
                <td class="auto-style5">
                    <asp:TextBox ID="txtIT39" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>

            </tr>
            <tr>
                <td class="tdDataCaption">
                    <p align="center">
                        IT37
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT37" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT36
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT36" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT2M
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT2M" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT2N
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT2N" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>

                <td class="tdDataCaption">
                    <p align="center">
                        IT2D
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT2D" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="tdDataCaption">
                    <p align="center">
                        IT2R
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT2R" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT31
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT31" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox></td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT24
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT24" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT25
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT25" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>

                <td class="tdDataCaption">
                    <p align="center">
                        IT32
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT32" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="tdDataCaption">
                    <p align="center">
                        IT33
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT33" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT2P
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT2P" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT2S
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT2S" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT34
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT34" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
                
                <td class="tdDataCaption">
                    <p align="center">
                        IT35
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT35" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="tdDataCaption">
                    <p align="center">
                        IT2U
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT2U" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT2T
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT2T" runat="server" ReadOnly="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT3C
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT3C" runat="server" ReadOnly ="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
                <td class="tdDataCaption">
                    <p align="center">
                        IT3D
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT3D" runat="server" ReadOnly ="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
                
                <td class="tdDataCaption">
                    <p align="center">
                        IT3E
                    </p>
                </td>
                <td class="tdDataField">
                    <asp:TextBox ID="txtIT3E" runat="server" ReadOnly ="True" Width="30px" CssClass="tdDataField"></asp:TextBox>
                </td>
            </tr>
            </tbody>
        </table>
        <table class="tblDataArea" width="100%">
            <tbody>
                <tr>
                    <td style="background-color: #99ccff; height: 30px" colspan="10">
                        <p align="left">
                            <b>[通路內控資料]</b>
                            <asp:Button ID="Button5" runat="server" Text="+" CssClass="cssUICommandButton"></asp:Button>
                        </p>
                    </td>
                </tr>
                <tr id="Shows3" runat="server">
                    <td class="tdDataCaption">通路代碼:</td>
                    <td class="tdDataField">
                        <asp:TextBox ID="txtICHANNEL" runat="server" Width="80px" MaxLength="20"></asp:TextBox></td>
                    <td class="tdDataCaption">大宗分公司代號:</td>
                    <td class="tdDataField">
                        <asp:TextBox ID="txtIBIG_BR" runat="server" Width="80px" MaxLength="2"></asp:TextBox></td>
                    <td class="tdDataCaption">*通路單位: </td>
                    <td class="tdDataField">
                        <asp:DropDownList ID="dllGetIrec" CssClass="TXTHEIGHT" runat="server">
                            <asp:ListItem Value="">請選擇</asp:ListItem>
                        </asp:DropDownList>
                        <asp:TextBox ID="txtIrecunit" runat="server" Width="80px" MaxLength="18"></asp:TextBox></td>
                    <td class="tdDataCaption">通路單位名稱:</td>
                    <td class="tdDataField">
                        <asp:TextBox ID="txtNRECUNIT" runat="server" Width="80px" MaxLength="20"></asp:TextBox></td>
                    <td class="tdDataCaption">團號:</td>
                    <td class="tdDataField">
                        <asp:TextBox ID="txtTrlCod" runat="server" Width="80px" MaxLength="50"></asp:TextBox></td>

                </tr>
                <tr id="Shows4" runat="server">
                    <td class="tdDataCaption">通路人員代號:<br/>（登錄證字號）</td>
                    <td class="tdDataField">
                        <asp:TextBox ID="txtIRECPESN" runat="server" Width="80px" MaxLength="20"></asp:TextBox></td>
                    <td class="tdDataCaption">通路人員名稱:</td>
                    <td class="tdDataField">
                        <asp:TextBox ID="txtNrecpesn" runat="server" Width="80px" MaxLength="20"></asp:TextBox></td>
                    <td class="tdDataCaption">*收件號: </td>
                    <td class="tdDataField">
                        <asp:TextBox ID="txtIREC" runat="server" Width="80px" MaxLength="13"></asp:TextBox></td>
                    <td class="tdDataCaption">收件人: </td>
                    <td class="tdDataField">
                        <asp:TextBox ID="txtIREC_OFFICR" runat="server" Width="80px" MaxLength="10"></asp:TextBox></td>
                    <td class="tdDataCaption">電傳人:</td>
                    <td class="tdDataField">
                        <asp:TextBox ID="txtTrlCal" runat="server" Width="80px" MaxLength="20"></asp:TextBox><asp:TextBox ID="txtJOB" Style="display: none" runat="server" Width="80px"></asp:TextBox></td>
                </tr>
            </tbody>
        </table>
        <tr id="divIISSU_NnE" runat="server">
            <td class="tdDataField" colspan="8">
                <p align="center">
                    <asp:Button ID="IISSU_N" runat="server" ForeColor="Black" CssClass="auto-style13" Text="新增被保人" Font-Bold="True" Height="30px"></asp:Button>&nbsp;&nbsp;&nbsp; 
                    <asp:Button ID="IISSU_E" runat="server" ForeColor="Black" CssClass="cssUICommandButton" Text="修改被保人" Font-Bold="True" Height="30px"></asp:Button>&nbsp;&nbsp;&nbsp; 

                    <asp:Label ID="txtTOTPREM" runat="server"></asp:Label>
                </p>
            </td>
        </tr>
        <tr>
            <td>
                <marquee><font style="COLOR: red; FONT-WEIGHT: bold" size="4">※保費資訊請以執行【保費試算】後為準。※</font></marquee>
            </td>
        </tr>
        </TBODY></TABLE>
        <table class="tblDataArea" width="95%" align="center">
            <tr>
                <td>
                    <uc1:datagridbanner id="DataGridBanner1" runat="server" AllowNewItem="false" DataGridID="dgdMain"></uc1:datagridbanner></td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:DataGrid ID="dgdMain" Width="100%" CssClass="cssDataGrid" PageSize="50" AllowPaging="True" HorizontalAlign="Center" AllowSorting="True" AutoGenerateColumns="False" runat="server">
                        <FooterStyle CssClass="cssDataGridFooter"></FooterStyle>
                        <SelectedItemStyle CssClass="cssDataGridSelectedItem"></SelectedItemStyle>
                        <EditItemStyle CssClass="cssDataGridEditItem"></EditItemStyle>
                        <AlternatingItemStyle CssClass="cssDataGridAlternatingItem"></AlternatingItemStyle>
                        <ItemStyle CssClass="cssDataGridItem"></ItemStyle>
                        <HeaderStyle ForeColor="#CCFFFF" CssClass="cssDataGridHeader"></HeaderStyle>
                        <Columns>
                            <asp:ButtonColumn Text="選取" CommandName="Select"></asp:ButtonColumn>
                            <asp:EditCommandColumn Visible="False" ButtonType="LinkButton" UpdateText="更新" CancelText="取消" EditText=""></asp:EditCommandColumn>
                            <asp:ButtonColumn Text="刪除" CommandName="Delete"></asp:ButtonColumn>
                            <asp:TemplateColumn SortExpression="ISEQ" HeaderText="序號">
                                <ItemTemplate>
                                    <asp:Label ID="labISEQ" runat="server" Text='<%# dgdMain.CurrentPageIndex * dgdMain.PageSize + Container.ItemIndex + 1%>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="NISSU" HeaderText="被保人">
                                <ItemTemplate>
                                    <asp:Label ID="labNISSU" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.NISSU") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="IISSU" HeaderText="身份證號">
                                <ItemTemplate>
                                    <asp:Label ID="labIISSU" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.IISSU") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="IISSU_SEX" HeaderText="性別">
                                <ItemTemplate>
                                    <asp:Label ID="labIISSU_SEX" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.IISSU_SEX") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="NISSUBD" HeaderText="出生日期">
                                <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="ucNISSUBD" runat="server" Text='<%# TransROCDate(DataBinder.Eval(Container, "DataItem.NISSUBD")) %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="NINSREL" HeaderText="與要保人關係">
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="labNINSREL" runat="server" Text='<%# GetTYPEREL(DataBinder.Eval(Container, "DataItem.NINSREL")) %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
<%--                            <asp:TemplateColumn SortExpression="JOBOFFICE" HeaderText="其他補充說明">
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="labJOBOFFICE" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.JOBOFFICE") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>--%>
                            <asp:TemplateColumn SortExpression="isPre" HeaderText="懷孕否">
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="labisPre" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.EMP") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn Visible="False" SortExpression="NISSU_TEL" HeaderText="被保人電話">
                                <ItemTemplate>
                                    <asp:Label ID="labNISSU_TEL" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.NISSU_TEL") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn Visible="False" SortExpression="IISSU_ZIP" HeaderText="郵遞區號">
                                <ItemTemplate>
                                    <asp:Label ID="lavIISSU_ZIP" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.IISSU_ZIP") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn Visible="False" SortExpression="AISSU" HeaderText="地址">
                                <ItemTemplate>
                                    <asp:Label ID="labAISSU" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.AISSU") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn Visible="False" SortExpression="NISSUMAIL" HeaderText="MAIL">
                                <ItemTemplate>
                                    <asp:Label ID="labNISSUMAIL" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.NISSUMAIL") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="ISSU_1" HeaderText="意外身故、失能或<br>意外喪葬費用">
                                <ItemTemplate>
                                    <asp:Label ID="labISSU_1" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ISSU_1") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="ISSU_IT44" HeaderText="重大燒燙傷保險金">
                                <ItemTemplate>
                                    <asp:Label ID="labISSU_IT44" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ISSU_IT44") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="ISSU_2" HeaderText="傷害醫療保險給付">
                                <ItemTemplate>
                                    <asp:Label ID="labISSU_2" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ISSU_2") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="ISSU_3" HeaderText="海外突發疾病醫療保險(不含傳染病)">
                                <ItemTemplate>
                                    <asp:Label ID="labISSU_3" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ISSU_3") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="ISSU_3D" HeaderText="海外突發疾病醫療保險(含傳染病)(A)">
                                <ItemTemplate>
                                    <asp:Label ID="labISSU_3D" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ISSU_3D") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="ISSU_IT64" HeaderText="海外突發疾病住院醫療保險金(B)">
                                <ItemTemplate>
                                    <asp:Label ID="labISSU_IT64" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ISSU_IT64") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="ISSU_IT43" HeaderText="大眾運輸意外保險金">
                                <ItemTemplate>
                                    <asp:Label ID="labISSU_IT43" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ISSU_IT43") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="NAMEREL" HeaderText="受益人">
                                <ItemTemplate>
                                    <asp:Label ID="labNAMEREL" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.NAMEREL") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="IDREL" HeaderText="受益人ID">
                                <ItemTemplate>
                                    <asp:Label ID="labIDREL" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.IDREL") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="TYPEREL" HeaderText="受益人與被保人關係">
                                <ItemTemplate>
                                    <asp:Label ID="labTYPEREL" runat="server" Text='<%# GetTYPEREL2(DataBinder.Eval(Container, "DataItem.TYPEREL")) %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="NAMEREL2" HeaderText="受益人2">
                                <ItemTemplate>
                                    <asp:Label ID="labNAMEREL2" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.NAMEREL2") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="IDREL2" HeaderText="受益人ID2">
                                <ItemTemplate>
                                    <asp:Label ID="labIDREL2" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.IDREL2") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="TYPEREL2" HeaderText="受益人與被保人關係2">
                                <ItemTemplate>
                                    <asp:Label ID="labTYPEREL2" runat="server" Text='<%# GetTYPEREL2(DataBinder.Eval(Container, "DataItem.TYPEREL2")) %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="NAMEREL3" HeaderText="受益人3">
                                <ItemTemplate>
                                    <asp:Label ID="labNAMEREL3" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.NAMEREL3") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="IDREL3" HeaderText="受益人ID3">
                                <ItemTemplate>
                                    <asp:Label ID="labIDREL3" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.IDREL3") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="TYPEREL3" HeaderText="受益人與被保人關係3">
                                <ItemTemplate>
                                    <asp:Label ID="labTYPEREL3" runat="server" Text='<%# GetTYPEREL2(DataBinder.Eval(Container, "DataItem.TYPEREL3")) %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="RELMOTHED" HeaderText="給付方式">
                                <ItemTemplate>
                                    <asp:Label ID="labRELMOTHED" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.RELMOTHED") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn Visible="False" SortExpression="TELREL" HeaderText="受益人聯絡電話">
                                <ItemTemplate>
                                    <asp:Label ID="labTELREL" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.TELREL") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn Visible="False" SortExpression="ZIPREL" HeaderText="受益人連絡郵遞區號">
                                <ItemTemplate>
                                    <asp:Label ID="labZIPREL" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ZIPREL") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn Visible="False" SortExpression="AREL" HeaderText="受益人連絡地址">
                                <ItemTemplate>
                                    <asp:Label ID="LabAREL" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.AREL") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="iPlanNo" HeaderText="選擇方案">
                                <ItemTemplate>
                                    <asp:Label ID="labiPlanNo" runat="server" Text='<%# GetiPlanNo(DataBinder.Eval(Container, "DataItem.iPlanNo")) %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="ISSU_4" HeaderText="保費(NTD)"><%--後端有對此欄HeaderText做判斷，要更改需留意 --%>
                                <ItemTemplate>
                                    <asp:Label ID="labISSU_4" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.ISSU_4") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="IISSU_COUNTRY" HeaderText="國別" Visible="False">
                                <ItemTemplate>
                                    <asp:Label ID="labIISSU_COUNTRY" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.IISSU_COUNTRY") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="IISSU_COUNTRY" HeaderText="英文名稱">
                                <ItemTemplate>
                                    <asp:Label ID="Label12" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.PASSPORT_NAME") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="IISSU_COUNTRY" HeaderText="護照號碼">
                                <ItemTemplate>
                                    <asp:Label ID="Label13" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.PASSPORT_NUMBER") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="ERR" HeaderText="備註">
                                <ItemTemplate>
                                    <asp:Label ID="labERR" runat="server" ForeColor="red" Text='<%# DataBinder.Eval(Container, "DataItem.ERR") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="JOBCONTENT" HeaderText="">
                                <ItemTemplate>
                                    <asp:Label ID="labJOBCONTENT" runat="server" ForeColor="red" Text='<%# DataBinder.Eval(Container, "DataItem.JOBCONTENT") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid></td>
            </tr>
        </table>
        <table class="tblDataArea" width="95%" style="display: none;" align="center">
            <tr>
                <td>
                    <uc1:datagridbanner id="Datagridbanner2" runat="server" AllowNewItem="false" DataGridID="dgdMain2"></uc1:datagridbanner></td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:DataGrid ID="dgdMain2" Width="100%" CssClass="cssDataGrid" runat="server" AutoGenerateColumns="False" AllowSorting="True" HorizontalAlign="Center" AllowPaging="True" PageSize="5">
                        <FooterStyle CssClass="cssDataGridFooter"></FooterStyle>
                        <SelectedItemStyle CssClass="cssDataGridSelectedItem"></SelectedItemStyle>
                        <EditItemStyle CssClass="cssDataGridEditItem"></EditItemStyle>
                        <AlternatingItemStyle CssClass="cssDataGridAlternatingItem"></AlternatingItemStyle>
                        <ItemStyle CssClass="cssDataGridItem"></ItemStyle>
                        <HeaderStyle ForeColor="#CCFFFF" CssClass="cssDataGridHeader"></HeaderStyle>
                        <Columns>
                            <asp:EditCommandColumn ButtonType="LinkButton" UpdateText="更新" CancelText="取消" EditText=""></asp:EditCommandColumn>
                            <asp:ButtonColumn CommandName="Delete"></asp:ButtonColumn>
                            <asp:TemplateColumn SortExpression="IQPMainOID" HeaderText="OID">
                                <ItemTemplate>
                                    <asp:Label ID="labIQPMainOID" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.IQPMainOID") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="INSTyp" HeaderText="編號">
                                <ItemTemplate>
                                    <asp:Label ID="labINSTyp" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.INSTyp") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="IISSU" HeaderText="被保險人ID">
                                <ItemTemplate>
                                    <asp:Label ID="labIISSU2" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.IISSU") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn Visible="False" SortExpression="CONTENT" HeaderText="保險內容">
                                <ItemTemplate>
                                    <asp:Label ID="labCONTENT" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.CONTENT") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="INSAMT" HeaderText="保險金額">
                                <ItemTemplate>
                                    <asp:Label ID="labINSAMT" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.INSAMT") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="INSPREM" HeaderText="保險費">
                                <ItemTemplate>
                                    <asp:Label ID="labINSPREM" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.INSPREM") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn Visible="False" SortExpression="MDEDUCT" HeaderText="MDEDUCT">
                                <ItemTemplate>
                                    <asp:Label ID="labMDEDCUT" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.MDEDUCT") %>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid></td>
            </tr>
        </table>
        </TD></TR></TBODY></TABLE><asp:ValidationSummary ID="vdsCS" runat="server" ShowSummary="False" ShowMessageBox="True" BorderWidth="0px" HeaderText="無法儲存，欄位資料錯誤"></asp:ValidationSummary>
    </form>
    <!--jQuery函數庫,使用Google-->
    <script>
</script>
</body>
</html>

