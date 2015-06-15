Imports System.Web
Imports System.Data
Imports System

Public Class WebForm1
    Inherits System.Web.UI.Page
    Dim bUserLoggedIn As Boolean = False
    Dim bPilot As Boolean = False

    Private Sub Page_PreRender(sender As Object, e As EventArgs) Handles Me.PreRender

        If IsPostBack Then
            LanguageDropDown.SelectedValue = LanguageDropDown.Text
        Else
            LanguageDropDown.SelectedValue = System.Web.HttpContext.Current.Session("sv_Set_Language")
        End If


        If LANG = False Then

            svInit()

        End If
        '
        System.Web.HttpContext.Current.Session("sv_Set_Language") = LanguageDropDown.Text

        MOD_Setupliteral(System.Web.HttpContext.Current.Session("sv_Set_Language"))
        '

        '
        ' Common literals
        '
        l_PageHead.Text = System.Web.HttpContext.Current.Session("sv_l_pagehead")


        l_ForAssist.Text = System.Web.HttpContext.Current.Session("sv_l_ForAssist")
        l_Qlinks.Text = System.Web.HttpContext.Current.Session("sv_l_Qlinks")
        '
        hl_download.Text = System.Web.HttpContext.Current.Session("sv_hl_download")
        hl_upload.Text = System.Web.HttpContext.Current.Session("sv_hl_upload")
        hl_other.Text = System.Web.HttpContext.Current.Session("sv_hl_other")

        '
        ' Page specific literals
        '
        '
        lbl_UserName.Text = System.Web.HttpContext.Current.Session("sv_lbl_UserName")
        lbl_Password.Text = System.Web.HttpContext.Current.Session("sv_lbl_Password")
        LogOnButton.Text = System.Web.HttpContext.Current.Session("sv_but_LogOnButton")
        '
        ' Customer specific literals
        '

        System.Web.HttpContext.Current.Session("sv_Path") = "D:\SinglePointFiles\"

        navPanel.Visible = False

    End Sub
    '
    ':::::::::::::::::::::::::::::::::::::::::::::::
    ':: Page load                                 ::
    ':::::::::::::::::::::::::::::::::::::::::::::::
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
    ':::::::::::::::::::::::::::::::::::::::::::::::
    ':: Log In button pressed.                    ::
    ':::::::::::::::::::::::::::::::::::::::::::::::

    Protected Sub LogOnButton_Click(sender As Object, e As EventArgs) Handles LogOnButton.Click
        LoginMsg.Visible = False
        If UserName.Text = vbNullString Then
            LoginMsg.Text = System.Web.HttpContext.Current.Session("sv_emessage1")
            LoginMsg.Visible = True
        Else
            ReadClientDetails(UserName.Text, Password.Text)
            '
            If System.Web.HttpContext.Current.Session("sv_LoggedIn") = False Then
                LoginMsg.Text = System.Web.HttpContext.Current.Session("sv_emessage2")
                LoginMsg.Visible = True
            Else
                Dim FTPP1 As New Process
                Dim StrParmsFTPP1 As String
                Dim StrResFTPP1 As String
                '
                StrParmsFTPP1 = Session("sv_AuditNo") & " " & Session("sv_UserID") & " " & Session("sv_ClientPW")
                '
                FTPP1.StartInfo.FileName = "vbCreatePublicUserAccounts.exe"
                FTPP1.StartInfo.WorkingDirectory = System.Web.HttpContext.Current.Session("sv_Path")
                FTPP1.StartInfo.Arguments = StrParmsFTPP1
                StrResFTPP1 = FTPP1.Start()
                FTPP1.WaitForExit()
                '
            End If
        End If
    End Sub
    ':::::::::::::::::::::::::::::::::::::::::::::::
    ':: Read client details                       ::
    ':::::::::::::::::::::::::::::::::::::::::::::::
    Sub ReadClientDetails(ByVal UserID As String, ByVal Passw As String)

        Dim oConnection As New SqlClient.SqlConnection
        Dim oCommand As New SqlClient.SqlCommand
        Dim oReader As SqlClient.SqlDataReader
        Dim ListLoaded As Boolean = False
        Dim cmdtx As String
        '
        oConnection.ConnectionString = "Server=uksowlmsql1.infor.com;Database=HEATPrimary;User ID=heat;Password=heat; Trusted_Connection=False; PersistSecurityInfo=True"
        'oConnection.ConnectionString = "Server=10.8.32.123;Database=HEATPrimary;User ID=heat;Password=heat; Trusted_Connection=False; PersistSecurityInfo=True"
        '
        ' Check server available 
        '

        Try
            oConnection.Open()
        Catch ex As Exception
            System.Web.HttpContext.Current.Session("sv_Ec") = "001"
            Response.Redirect("ErrorPage.aspx")
            '	
        End Try
        '
        cmdtx = ""
        cmdtx = cmdtx & "SELECT "
        cmdtx = cmdtx & "Subset.CallID,"
        cmdtx = cmdtx & "Subset.CustID,"
        cmdtx = cmdtx & "Subset.MainProduct,"
        cmdtx = cmdtx & "Subset.Platform,"
        cmdtx = cmdtx & "Subset.Attributes,"
        cmdtx = cmdtx & "Subset.CountryCode,"
        cmdtx = cmdtx & "Subset.ClientUserID,"
        cmdtx = cmdtx & "Subset.ClientPassword,"
        cmdtx = cmdtx & "Subset.Company_Name,"
        cmdtx = cmdtx & "Profile.CityOrTown "
        cmdtx = cmdtx & "FROM Profile,Subset WHERE Subset.CustID = Profile.CustID And Subset.ClientUserID = '" & UserID & "'"

        oCommand.CommandText = cmdtx

        oCommand.Connection = oConnection

        oReader = oCommand.ExecuteReader

        If oReader.HasRows = True Then
            oReader.Read()
            '
            System.Web.HttpContext.Current.Session("sv_ClientPW") = oReader.Item("ClientPassword")
            System.Web.HttpContext.Current.Session("sv_ClientUserID") = oReader.Item("ClientUserID")

            '
            If System.Web.HttpContext.Current.Session("sv_ClientPW") = Passw And System.Web.HttpContext.Current.Session("sv_ClientUserID") = UserID Then
                '
                bUserLoggedIn = True
                System.Web.HttpContext.Current.Session("sv_AuditNo") = oReader.Item("CallID")
                System.Web.HttpContext.Current.Session("sv_MainProduct") = oReader.Item("MainProduct")
                System.Web.HttpContext.Current.Session("sv_Platform") = oReader.Item("Platform")
                System.Web.HttpContext.Current.Session("sv_Attributes") = oReader.Item("Attributes")
                System.Web.HttpContext.Current.Session("sv_CompanyName") = oReader.Item("Company_Name")
                '
                If DBNull.Value.Equals(oReader.Item("CityOrTown")) Then
                    'If oReader.Item("CityOrTown") = "" Then
                    System.Web.HttpContext.Current.Session("sv_CityOrTown") = "Head Office"
                Else
                    System.Web.HttpContext.Current.Session("sv_CityOrTown") = oReader.Item("CityOrTown")
                End If
                '
                System.Web.HttpContext.Current.Session("sv_LoggedIn") = bUserLoggedIn
                System.Web.HttpContext.Current.Session("sv_UserID") = UserID

                '
                Dim dd As String
                dd = System.Web.HttpContext.Current.Session("sv_AuditNo")
                '
                Response.Redirect("Welcome.aspx")
            Else
                LoginMsg.Text = System.Web.HttpContext.Current.Session("sv_LogInMsg1")
                LoginMsg.Visible = True
                '
                bUserLoggedIn = False
                System.Web.HttpContext.Current.Session("sv_LoggedIn") = bUserLoggedIn
            End If
            '
        Else
            LoginMsg.Text = System.Web.HttpContext.Current.Session("sv_LogInMsg1")
            LoginMsg.Visible = True
            '
            bUserLoggedIn = False
            System.Web.HttpContext.Current.Session("sv_LoggedIn") = bUserLoggedIn
        End If

        oConnection.Close()
    End Sub

    Public Sub svInit()

        '
        System.Web.HttpContext.Current.Session.Add("sv_Path", "")

        System.Web.HttpContext.Current.Session.Add("sv_lbl_Password", "")
        System.Web.HttpContext.Current.Session.Add("sv_but_LogOnButton", "")
        '
        System.Web.HttpContext.Current.Session.Add("sv_lbl_AuditToolDownload", "")
        System.Web.HttpContext.Current.Session.Add("sv_lbl_InstallInstruct", "")
        System.Web.HttpContext.Current.Session.Add("sv_lbl_TechSpec", "")
        System.Web.HttpContext.Current.Session.Add("sv_lbl_HelpGuide", "")
        System.Web.HttpContext.Current.Session.Add("sv_Dwn_Button1", "")
        System.Web.HttpContext.Current.Session.Add("sv_Dwn_Button2", "")
        System.Web.HttpContext.Current.Session.Add("sv_Dwn_Button3", "")
        System.Web.HttpContext.Current.Session.Add("sv_Dwn_Button4", "")
        '
        System.Web.HttpContext.Current.Session.Add("sv_l_welcome1", "")
        System.Web.HttpContext.Current.Session.Add("sv_l_welcome2", "")
        System.Web.HttpContext.Current.Session.Add("sv_l_welcome3", "")
        System.Web.HttpContext.Current.Session.Add("sv_l_welcome4", "")
        '
        System.Web.HttpContext.Current.Session.Add("sv_hl_download", "")
        System.Web.HttpContext.Current.Session.Add("sv_hl_upload", "")
        System.Web.HttpContext.Current.Session.Add("sv_hl_other", "")
        '
        System.Web.HttpContext.Current.Session.Add("sv_Set_Language", "")
        System.Web.HttpContext.Current.Session.Add("sv_Dwn_Other_File", "")
        LANG = True
    End Sub

End Class