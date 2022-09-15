<#	
	.NOTES
	===========================================================================
     Version:       1.0.6
	 Updated on:   	9/14/2022
	 Created by:   	/u/TheLazyAdministrator
     Contributors:  /u/ascIVV, /u/jmn_lab, /u/nothingpersonalbro
	===========================================================================

        ExchangeOnlineManagement Module is required
            Install-Module -Name AzureAD
            https://www.powershellgallery.com/packages/ExchangeOnlineManagement/
        AzureAD  Module is required
            Install-Module -Name AzureAD
            https://www.powershellgallery.com/packages/azuread/
        ReportHTML Moduile is required
            Install-Module -Name ReportHTML
            https://www.powershellgallery.com/packages/ReportHTML/

        UPDATES
        1.0.5
            /u/ascIVV: Added the following:
                - Admin Tab
                    - Privileged Role Administrators
                    - Exchange Administrators
                    - User Account Administrators
                    - Tech Account Restricted Exchange Admin Role
                    - SharePoint Administrators
                    - Skype Administrators
                    - CRM Service Administrators
                    - Power BI Administrators
                    - Service Support Administrators
                    - Billing Administrators
            /u/TheLazyAdministrator
                - Cleaned up formatting
                - Error Handling for $Null obj
                - Console status
                - Windows Defender ATP SKU
        

	.DESCRIPTION
		Generate an interactive HTML report on your Office 365 tenant. Report on Users, Tenant information, Groups, Policies, Contacts, Mail Users, Licenses and more!
    
    .Link
        Original: http://thelazyadministrator.com/2018/06/22/create-an-interactive-html-report-for-office-365-with-powershell/
#>
#########################################
#                                       #
#            VARIABLES                  #
#                                       #
#########################################

#Company logo that will be displayed on the left, can be URL or UNC
$CompanyLogo = "http://thelazyadministrator.com/wp-content/uploads/2018/06/logo-2-e1529684959389.png"

#Logo that will be on the right side, UNC or URL
$RightLogo = "http://thelazyadministrator.com/wp-content/uploads/2018/06/amd.png"

#Location the report will be saved to
$ReportSavePath = "C:\Automation\"

#Variable to filter licenses out, in current state will only get licenses with a count less than 9,000 this will help filter free/trial licenses
$LicenseFilter = "9000"

#Set to $True if your global admin requires 2FA
$2FA = $True

########################################


If ($2FA -eq $False)
{
    $credential = Get-Credential -Message "Please enter your Office 365 credentials"
    Import-Module AzureAD
    Connect-AzureAD -Credential $credential
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/"  -Authentication "Basic" -AllowRedirection -Credential $credential
    Import-PSSession $exchangeSession -AllowClobber
}
Else
{
    $Modules = dir $Env:LOCALAPPDATA\Apps\2.0\*\CreateExoPSSession.ps1 -Recurse | Select-Object -ExpandProperty Target -First 1
    foreach ($Module in $Modules)
    {
     Import-Module "$Module"
    }
    Write-Host "Credential prompt to connect to Azure Graph" -ForegroundColor Yellow
    #Connect to Azure Graph w/2FA
    Connect-AzureAD

    Write-Host "Credential prompt to connect to Azure" -ForegroundColor Yellow
	#Connect to Azure w/ 2FA
    Connect-MSOLService

    Write-Host "Credential prompt to connect to Exchange Online" -ForegroundColor Yellow
    #Connect to Exchange Online w/ 2FA
    Connect-ExchangeOnline
}


$Table = New-Object 'System.Collections.Generic.List[System.Object]'
$LicenseTable = New-Object 'System.Collections.Generic.List[System.Object]'
$UserTable = New-Object 'System.Collections.Generic.List[System.Object]'
$SharedMailboxTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GroupTypetable = New-Object 'System.Collections.Generic.List[System.Object]'
$IsLicensedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ContactTable = New-Object 'System.Collections.Generic.List[System.Object]'
$MailUser = New-Object 'System.Collections.Generic.List[System.Object]'
$ContactMailUserTable = New-Object 'System.Collections.Generic.List[System.Object]'
$RoomTable = New-Object 'System.Collections.Generic.List[System.Object]'
$EquipTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GlobalAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ExchangeAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$PrivAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$UserAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$TechExchAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$SharePointAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$SkypeAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$CRMAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$PowerBIAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ServiceAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$BillingAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
$StrongPasswordTable = New-Object 'System.Collections.Generic.List[System.Object]'
$CompanyInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DomainTable = New-Object 'System.Collections.Generic.List[System.Object]'

$Sku = @{
	"ADV_COMMS"							     = "Advanced Communications"
	"CDSAICAPACITY"							     = "AI Builder Capacity add-on"
	"SPZA_IW"							     = "App Connect IW"
	"MCOMEETADV"							     = "Microsoft 365 Audio Conferencing"
	"AAD_BASIC"							     = "Azure Active Directory Basic"
	"AAD_PREMIUM"							     = "Azure Active Directory Premium P1"
	"AAD_PREMIUM_P2"						     = "Azure Active Directory Premium P2"
	"RIGHTSMANAGEMENT"						     = "Azure Information Protection Plan 1"
	"SMB_APPS"							     = "Business Apps (free)"
	"MCOCAP"							     = "Common Area Phone"
	"MCOCAP_GOV"							     = "Common Area Phone for GCC"
	"CDS_DB_CAPACITY"						     = "Common Data Service Database Capacity"
	"CDS_DB_CAPACITY_GOV"						     = "Common Data Service Database Capacity for Government"
	"CDS_LOG_CAPACITY"						     = "Common Data Service Log Capacity"
	"MCOPSTNC"							     = "Communications Credtis"
	"CMPA_addon_GCC"						     = "Compliance Manager Premium Assessment Add-On for GCC"
	"CRMSTORAGE"							     = "Dynamics 365 - Additional Database Storage (Qualified Offer)"
	"CRMINSTANCE"							     = "Dynamics 365 - Additional Production Instance (Qualified Offer)"
	"CRMTESTINSTANCE"						     = "Dynamics 365 - Additional Non-Production Instance (Qualified Offer)"
	"SOCIAL_ENGAGEMENT_APP_USER "					     = "Dynamics 365 AI for Market Insights (Preview)"
	"DYN365_ASSETMANAGEMENT"					     = "Dynamics 365 Asset Management Addl Assets"
	"DYN365_BUSCENTRAL_ADD_ENV_ADDON"				     = "Dynamics 365 Business Central Additional Environment Addon"
	"DYN365_BUSCENTRAL_DB_CAPACITY"					     = "Dynamics 365 Business Central Database Capacity"
	"DYN365_BUSCENTRAL_ESSENTIAL"					     = "Dynamics 365 Business Central Essentials"
	"DYN365_FINANCIALS_ACCOUNTANT_SKU"				     = "Dynamics 365 Business Central External Accountant"
	"PROJECT_MADEIRA_PREVIEW_IW_SKU"				     = "Dynamics 365 Business Central for IWs"
	"DYN365_BUSCENTRAL_PREMIUM"					     = "Dynamics 365 Business Central Premium"
	"DYN365_BUSCENTRAL_TEAM_MEMBER "				     = "Dynamics 365 Business Central Team Members "
	"DYN365_ENTERPRISE_PLAN1"					     = "Dynamics 365 Customer Engagement Plan"
	"DYN365_CUSTOMER_INSIGHTS_VIRAL"				     = "Dynamics 365 Customer Insights vTrial "
	"Dynamics_365_Customer_Service_Enterprise_viral_trial "		     = "Dynamics 365 Customer Service Enterprise Viral Trial "
	"DYN365_AI_SERVICE_INSIGHTS"					     = "Dynamics 365 Customer Service Insights Trial"
	"FORMS_PRO"							     = "Dynamics 365 Customer Voice Trial"
	"DYN365_CUSTOMER_SERVICE_PRO"					     = "Dynamics 365 Customer Service Professional"
	"DYN365_CUSTOMER_VOICE_BASE"					     = "Dynamics 365 Customer Voice"
	"Forms_Pro_AddOn"						     = "Dynamics 365 Customer Voice Additional Responses"
	"Forms_Pro_USL"							     = "Dynamics 365 Customer Voice USL"
	"CRM_ONLINE_PORTAL"						     = "Dynamics 365 Enterprise Edition - Additional Portal (Qualified Offer)"
	"Dynamics_365_Field_Service_Enterprise_viral_trial "		     = "Dynamics 365 Field Service Viral Trial "
	"DYN365_FINANCE"						     = "Dynamics 365 Finance"
	"DYN365_ENTERPRISE_CUSTOMER_SERVICE"				     = "Dynamics 365 for Customer Service Enterprise Edition"
	"D365_FIELD_SERVICE_ATTACH"					     = "Dynamics 365 for Field Service Attach to Qualifying Dynamics 365 Base Offer"
	"DYN365_ENTERPRISE_FIELD_SERVICE "				     = "Dynamics 365 for Field Service Enterprise Edition "
	"DYN365_FINANCIALS_BUSINESS_SKU"				     = "Dynamics 365 for Financials Business Edition"
	"D365_MARKETING_USER"						     = "Dynamics 365 for Marketing USL"
	"DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE"			     = "Dynamics 365 for Sales and Customer Service Enterprise Edition"
	"DYN365_ENTERPRISE_SALES"					     = "Dynamics 365 for Sales Enterprise Edition"
	"DYN365_BUSINESS_MARKETING "					     = "Dynamics 365 for Marketing Business Edition"
	"DYN365_REGULATORY_SERVICE "					     = "Dynamics 365 Regulatory Service - Enterprise Edition Trial "
	"Dynamics_365_Sales_Premium_Viral_Trial "			     = "Dynamics 365 Sales Premium Viral Trial "
	"D365_SALES_PRO "						     = "Dynamics 365 For Sales Professional"
	"D365_SALES_PRO_IW "						     = "Dynamics 365 For Sales Professional Trial "
	"D365_SALES_PRO_ATTACH"						     = "Dynamics 365 Sales Professional Attach to Qualifying Dynamics 365 Base Offer"
	"DYN365_SCM"							     = "Dynamics 365 for Supply Chain Management"
	"SKU_Dynamics_365_for_HCM_Trial"				     = "Dynamics 365 for Talent"
	"Dynamics_365_Hiring_SKU "					     = "Dynamics 365 Talent: Attract "
	"DYN365_ENTERPRISE_TEAM_MEMBERS"				     = "Dynamics 365 for Team Members Enterprise Edition"
	"GUIDES_USER"							     = "Dynamics 365 Guides"
	"Dynamics_365_for_Operations_Devices"				     = "Dynamics 365 Operations - Device"
	"Dynamics_365_for_Operations_Sandbox_Tier2_SKU"			     = "Dynamics 365 Operations - Sandbox Tier 2:Standard Acceptance Testing"
	"Dynamics_365_for_Operations_Sandbox_Tier4_SKU"			     = "Dynamics 365 Operations - Sandbox Tier 4:Standard Performance Testing"
	"DYN365_ENTERPRISE_P1_IW"					     = "Dynamics 365 P1 Trial for Information Workers"
	"MICROSOFT_REMOTE_ASSIST"					     = "Dynamics 365 Remote Assist"
	"MICROSOFT_REMOTE_ASSIST_HOLOLENS"				     = "Dynamics 365 Remote Assist HoloLens"
	"D365_SALES_ENT_ATTACH"						     = "Dynamics 365 Sales Enterprise Attach to Qualifying Dynamics 365 Base Offer"
	"DYNAMICS_365_ONBOARDING_SKU"					     = "Dynamics 365 Talent: Onboard"
	"DYN365_TEAM_MEMBERS"						     = "Dynamics 365 Team Members"
	"Dynamics_365_for_Operations"					     = "Dynamics 365 UNF OPS Plan ENT Edition"
	"EMS_EDU_FACULTY"						     = "Enterprise Mobility + Security A3 for Faculty "
	"EMS"							     	     = "Enterprise Mobility + Security E3"
	"EMSPREMIUM"							     = "Enterprise Mobility + Security E5"
	"EMS_GOV"							     = "Enterprise Mobility + Security G3 GCC"
	"EOP_ENTERPRISE_PREMIUM"					     = "Exchange Enterprise CAL Services (EOP DLP)"
	"EXCHANGESTANDARD"						     = "Exchange Online (Plan 1)"
	"EXCHANGESTANDARD_GOV"						     = "Exchange Online (Plan 1) for GCC"
	"EXCHANGEENTERPRISE"						     = "Exchange Online (Plan 2)"
	"EXCHANGEARCHIVE_ADDON"						     = "Exchange Online Archiving for Exchange Online"
	"EXCHANGEARCHIVE"						     = "Exchange Online Archiving for Exchange Server"
	"EXCHANGEESSENTIALS"						     = "Exchange Online Essentials (ExO P1 Based)"
	"EXCHANGE_S_ESSENTIALS"						     = "Exchange Online Essentials"
	"EXCHANGEDESKLESS"						     = "Exchange Online Kiosk"
	"EXCHANGETELCO"							     = "Exchange Online POP"
	"EOP_ENTERPRISE"						     = "Exchange Online Protection"
	"INTUNE_A"							     = "Intune"
	"AX7_USER_TRIAL"						     = "Microsoft Dynamics AX7 User Trial"
	"MFA_STANDALONE"						     = "Microsoft Azure Multi-Factor Authentication"
	"THREAT_INTELLIGENCE"						     = "Microsoft Defender for Office 365 (Plan 2)"
	"M365EDU_A1"							     = "Microsoft 365 A1"
	"M365EDU_A3_FACULTY"						     = "Microsoft 365 A3 for Faculty"
	"M365EDU_A3_STUDENT"						     = "Microsoft 365 A3 for Students"
	"M365EDU_A3_STUUSEBNFT"						     = "Microsoft 365 A3 for students use benefit"
	"M365EDU_A3_STUUSEBNFT_RPA1"					     = "Microsoft 365 A3 - Unattended License for students use benefit"
	"M365EDU_A5_FACULTY"						     = "Microsoft 365 A5 for Faculty"
	"M365EDU_A5_STUDENT"						     = "Microsoft 365 A5 for Students"
	"M365EDU_A5_STUUSEBNFT"						     = "Microsoft 365 A5 for students use benefit"
	"M365EDU_A5_NOPSTNCONF_STUUSEBNFT"				     = "Microsoft 365 A5 without Audio Conferencing for students use benefit"
	"O365_BUSINESS"							     = "Microsoft 365 Apps for Business"
	"SMB_BUSINESS"							     = "Microsoft 365 Apps for Business"
	"OFFICESUBSCRIPTION"						     = "Microsoft 365 Apps for Enterprise"
	"OFFICE_PROPLUS_DEVICE1"					     = "Microsoft 365 Apps for enterprise (device)"
	"OFFICESUBSCRIPTION_STUDENT"					     = "Microsoft 365 Apps for Students"
	"OFFICESUBSCRIPTION_FACULTY"					     = "Microsoft 365 Apps for Faculty"
	"MCOMEETADV_GOC"						     = "Microsoft 365 Audio Conferencing for GCC"
	"SMB_BUSINESS_ESSENTIALS"					     = "Microsoft 365 Business Basic"
	"O365_BUSINESS_PREMIUM"						     = "Microsoft 365 Business Standard"
	"SMB_BUSINESS_PREMIUM"						     = "Microsoft 365 Business Standard - Prepaid Legacy"
	"SPB"							     	     = "Microsoft 365 Business Premium"
	"BUSINESS_VOICE_MED2"						     = "Microsoft 365 Business Voice"
	"BUSINESS_VOICE_MED2_TELCO"					     = "Microsoft 365 Business Voice (US)"
	"BUSINESS_VOICE_DIRECTROUTING"					     = "Microsoft 365 Business Voice (without calling plan) "
	"BUSINESS_VOICE_DIRECTROUTING_MED"				     = "Microsoft 365 Business Voice (without Calling Plan) for US"
	"MCOPSTN_5"							     = "MICROSOFT 365 DOMESTIC CALLING PLAN (120 Minutes)"
	"MCOPSTN_1_GOV"							     = "Microsoft 365 Domestic Calling Plan for GCC"
	"SPE_E3"							     = "Microsoft 365 E3"
	"SPE_E3_RPA1"							     = "Microsoft 365 E3 - Unattended License"
	"SPE_E3_USGOV_DOD"						     = "Microsoft 365 E3_USGOV_DOD"
	"SPE_E3_USGOV_GCCHIGH"						     = "Microsoft 365 E3_USGOV_GCCHIGH"
	"SPE_E5"							     = "Microsoft 365 E5"
	"DEVELOPERPACK_E5"						     = "Microsoft 365 E5 Developer (without Windows and Audio Conferencing)"
	"INFORMATION_PROTECTION_COMPLIANCE"				     = "Microsoft 365 E5 Compliance"
	"IDENTITY_THREAT_PROTECTION"					     = "Microsoft 365 E5 Security"
	"IDENTITY_THREAT_PROTECTION_FOR_EMS_E5"				     = "Microsoft 365 E5 Security for EMS E5"
	"SPE_E5_NOPSTNCONF"						     = "Microsoft 365 E5 without Audio Conferencing "
	"M365_F1"							     = "Microsoft 365 F1"
	"SPE_F1"							     = "Microsoft 365 F3"
	"M365_F1_GOV"							     = "Microsoft 365 F3 GCC"
	"SPE_F5_SEC"							     = "Microsoft 365 F5 Security Add-on"
	"SPE_F5_SECCOMP "						     = "Microsoft 365 F5 Security + Compliance Add-on "
	"M365_G5_GCC"							     = "Microsoft 365 GCC G5"
	"FLOW_FREE"							     = "Microsoft Flow Free"
	"MCOMEETADV_GOV"						     = "Microsoft 365 Audio Conferencing for GCC"
	"M365_E5_SUITE_COMPONENTS"					     = "Microsoft 365 E5 Suite features"
	"M365_F1_COMM"							     = "Microsoft 365 F1"
	"M365_G3_GOV"							     = "Microsoft 365 G3 GCC"
	"MCOEV"								     = "Microsoft 365 Phone System"
	"MCOEV_DOD"							     = "Microsoft 365 Phone System for DOD"
	"MCOEV_FACULTY"							     = "Microsoft 365 Phone System for Faculty"
	"MCOEV_GOV"							     = "Microsoft 365 Phone System for GCC"
	"MCOEV_GCCHIGH"							     = "Microsoft 365 Phone System for GCCHIGH"
	"MCOEVSMB_1"							     = "Microsoft 365 Phone System for Small and Medium Business"
	"MCOEV_STUDENT"							     = "Microsoft 365 Phone System for Students"
	"MCOEV_TELSTRA"							     = "Microsoft 365 Phone System for TELSTRA"
	"MCOEV_USGOV_DOD"						     = "Microsoft 365 Phone System_USGOV_DOD"
	"MCOEV_USGOV_GCCHIGH"						     = "Microsoft 365 Phone System_USGOV_GCCHIGH"
	"PHONESYSTEM_VIRTUALUSER"					     = "Microsoft 365 Phone System - Virtual User"
	"PHONESYSTEM_VIRTUALUSER_GOV"					     = "Microsoft 365 Phone System - Virtual User for GCC"
	"M365_SECURITY_COMPLIANCE_FOR_FLW"				     = "Microsoft 365 Security and Compliance for Firstline Workers"
	"MICROSOFT_BUSINESS_CENTER"					     = "Microsoft Business Center"
	"ADALLOM_STANDALONE"						     = "Microsoft Cloud App Security"
	"WIN_DEF_ATP"							     = "Microsoft Defender for Endpoint"
	"DEFENDER_ENDPOINT_P1"						     = "Microsoft Defender for Endpoint P1"
	"MDATP_XPLAT"							     = "Microsoft Defender for Endpoint P2_XPLAT"
	"MDATP_Server"							     = "Microsoft Defender for Endpoint Server"
	"CRMPLAN2"							     = "Microsoft Dynamics CRM Online Basic"
	"ATA"							  	     = "Microsoft Defender for Identity"
	"ATP_ENTERPRISE_FACULTY"					     = "Microsoft Defender for Office 365 (Plan 1) Faculty"
	"ATP_ENTERPRISE_GOV"						     = "Microsoft Defender for Office 365 (Plan 1) GCC "
	"CRMSTANDARD"							     = "Microsoft Dynamics CRM Online"
	"IT_ACADEMY_AD"							     = "Microsoft Imagine Academy"
	"INTUNE_A_D"							     = "Microsoft Intune Device"
	"INTUNE_A_D_GOV"						     = "Microsoft Intune Device for Government"
	"POWERAPPS_DEV "						     = "Microsoft Power Apps for Developer "
	"POWERAPPS_VIRAL"						     = "Microsoft Power Apps Plan 2 Trial"
	"FLOW_P2"							     = "Microsoft Power Automate Plan 2"
	"INTUNE_SMB"							     = "Microsoft Intune SMB"
	"POWERFLOW_P2"							     = "Microsoft Power Apps Plan 2 (Qualified Offer)"
	"STREAM"							     = "Microsoft Stream"
	"STREAM_P2"							     = "Microsoft Stream Plan 2"
	"STREAM_STORAGE"						     = "Microsoft Stream Storage Add-On (500 GB)"
	"Microsoft_Teams_Audio_Conferencing_select_dial_out"		     = "Microsoft Teams Audio Conferencing select dial-out"
	"TEAMS_FREE"							     = "Microsoft Teams (Free)"
	"TEAMS_EXPLORATORY"						     = "Microsoft Teams Exploratory"
	"MEETING_ROOM"							     = "Microsoft Teams Rooms Standard"
	"MEETING_ROOM_NOAUDIOCONF"					     = "Microsoft Teams Rooms Standard without Audio Conferencing"
	"MS_TEAMS_IW"							     = "Microsoft Teams Trial"
	"EXPERTS_ON_DEMAND"						     = "Microsoft Threat Experts - Experts on Demand"
	"OFFICE365_MULTIGEO"						     = "Multi-Geo Capabilities in Office 365"
	"NONPROFIT_PORTAL"						     = "Nonprofit Portal"
	"STANDARDWOFFPACK_FACULTY"					     = "Office 365 A1 for faculty"
	"STANDARDWOFFPACK_IW_FACULTY"					     = "Office 365 A1 Plus for faculty"
	"STANDARDWOFFPACK_STUDENT "					     = "Office 365 A1 for students "
	"STANDARDWOFFPACK_IW_STUDENT"					     = "Office 365 A1 Plus for students"
	"ENTERPRISEPACKPLUS_FACULTY"					     = "Office 365 A3 for faculty"
	"ENTERPRISEPACKPLUS_STUDENT"					     = "Office 365 A3 for students"
	"ENTERPRISEPREMIUM_FACULTY"					     = "Office 365 A5 for faculty"
	"ENTERPRISEPREMIUM_STUDENT"					     = "Office 365 A5 for students"
	"EQUIVIO_ANALYTICS"						     = "Office 365 Advanced Compliance"
	"EQUIVIO_ANALYTICS_GOV"						     = "Office 365 Advanced Compliance for GCC"
	"ATP_ENTERPRISE"						     = "Microsoft Defender for Office 365 (Plan 1)"
	"SHAREPOINTSTORAGE_GOV"						     = "Office 365 Extra File Storage for GCC"
	"TEAMS_COMMERCIAL_TRIAL"					     = "Microsoft Teams Commercial Cloud"
	"ADALLOM_O365"							     = "Office 365 Cloud App Security"
	"SHAREPOINTSTORAGE"						     = "Office 365 Extra File Storage"
	"STANDARDPACK"							     = "Office 365 E1"
	"STANDARDWOFFPACK"						     = "OFFICE 365 E2"
	"ENTERPRISEPACK"						     = "Office 365 E3"
	"DEVELOPERPACK"							     = "Office 365 E3 Developer"
	"ENTERPRISEPACK_USGOV_DOD"					     = "Office 365 E3_USGOV_DOD"
	"ENTERPRISEPACK_USGOV_GCCHIGH"					     = "Office 365 E3_USGOV_GCCHIGH"
	"ENTERPRISEWITHSCAL"						     = "Office 365 E4"
	"ENTERPRISEPREMIUM"						     = "Office 365 E5"
	"ENTERPRISEPREMIUM_NOPSTNCONF"					     = "Office 365 E5 Without Audio Conferencing"
	"DESKLESSPACK"							     = "OFFICE 365 F3"
	"STANDARDPACK_GOV"						     = "Office 365 G1 GCC"
	"ENTERPRISEPACK_GOV"						     = "OFFICE 365 G3 GCC"
	"ENTERPRISEPREMIUM_GOV"						     = "Office 365 G5 GCC"
	"MIDSIZEPACK"							     = "Office 365 Midsize Business"
	"LITEPACK"							     = "Office 365 Small Business"
	"LITEPACK_P2"							     = "Office 365 Small Business Premium"
	"WACONEDRIVESTANDARD"						     = "OneDrive for Business (Plan 1)"
	"WACONEDRIVEENTERPRISE"						     = "OneDrive for Business (Plan 2)"
	"POWERAPPS_INDIVIDUAL_USER"					     = "Power Apps and Logic Flows"
	"POWERAPPS_PER_APP_IW"						     = "PowerApps per app baseline access"
	"POWERAPPS_PER_APP"						     = "Power Apps per app plan"
	"POWERAPPS_PER_APP_NEW "					     = "Power Apps per app plan (1 app or portal)"
	"POWERAPPS_PER_USER"						     = "Power Apps per user plan"
	"POWERAPPS_PER_USER_GCC"					     = "Power Apps per user plan for Government"
	"POWERAPPS_P1_GOV"						     = "PowerApps Plan 1 for Government"
	"POWERAPPS_PORTALS_LOGIN_T2"					     = "Power Apps Portals login capacity add-on Tier 2 (10 unit min)"
	"POWERAPPS_PORTALS_LOGIN_T2_GCC"				     = "Power Apps Portals login capacity add-on Tier 2 (10 unit min) for Government"
	"POWERAPPS_PORTALS_PAGEVIEW_GCC"				     = "Power Apps Portals page view capacity add-on for Government"
	"FLOW_BUSINESS_PROCESS"						     = "Power Automate per flow plan"
	"FLOW_PER_USER"							     = "Power Automate per user plan"
	"FLOW_PER_USER_DEPT"						     = "Power Automate per user plan dept"
	"FLOW_PER_USER_GCC"						     = "Power Automate per user plan for Government"
	"POWERAUTOMATE_ATTENDED_RPA"					     = "Power Automate per user with attended RPA plan"
	"POWERAUTOMATE_UNATTENDED_RPA"					     = "Power Automate unattended RPA add-on"
	"POWER_BI_INDIVIDUAL_USER"					     = "Power BI"
	"POWER_BI_STANDARD"						     = "Power BI (free)"
	"POWER_BI_ADDON"						     = "Power BI for Office 365 Add-On"
	"PBI_PREMIUM_P1_ADDON"						     = "Power BI Premium P1"
	"PBI_PREMIUM_PER_USER"						     = "Power BI Premium Per User"
	"PBI_PREMIUM_PER_USER_ADDON"					     = "Power BI Premium Per User Add-On"
	"PBI_PREMIUM_PER_USER_DEPT"					     = "Power BI Premium Per User Dept"
	"POWER_BI_PRO"							     = "Power BI Pro"
	"POWER_BI_PRO_CE"						     = "Power BI Pro CE"
	"POWER_BI_PRO_DEPT"						     = "Power BI Pro Dept"
	"POWERBI_PRO_GOV"						     = "Power BI Pro for GCC"
	"VIRTUAL_AGENT_BASE"						     = "Power Virtual Agent"
	"CCIBOTS_PRIVPREV_VIRAL"					     = "Power Virtual Agents Viral Trial"
	"PROJECTCLIENT"							     = "Project for Office 365"
	"PROJECTESSENTIALS"						     = "Project Online Essentials"
	"PROJECTESSENTIALS_GOV"						     = "Project Online Essentials for GCC"
	"PROJECTPREMIUM"						     = "Project Online Premium"
	"PROJECTONLINE_PLAN_1"						     = "Project Online Premium Without Project Client"
	"PROJECTONLINE_PLAN_2"						     = "Project Online With Project for Office 365"
	"PROJECT_P1"							     = "Project Plan 1"
	"PROJECT_PLAN1_DEPT"						     = "Project Plan 1 (for Department)"
	"PROJECTPROFESSIONAL"						     = "Project Plan 3"
	"PROJECT_PLAN3_DEPT"						     = "Project Plan 3 (for Department)"
	"PROJECTPROFESSIONAL_GOV"					     = "Project Plan 3 for GCC"
	"PROJECTPREMIUM_GOV"						     = "Project Plan 5 for GCC"
	"RIGHTSMANAGEMENT_ADHOC"					     = "Rights Management Adhoc"
	"RMSBASIC"							     = "Rights Management Service Basic Content Protection"
	"DYN365_IOT_INTELLIGENCE_ADDL_MACHINES"				     = "Sensor Data Intelligence Additional Machines Add-in for Dynamics 365 Supply Chain Management"
	"DYN365_IOT_INTELLIGENCE_SCENARIO"				     = "Sensor Data Intelligence Scenario Add-in for Dynamics 365 Supply Chain Management"
	"SHAREPOINTSTANDARD"						     = "SharePoint Online (Plan 1)"
	"SHAREPOINTENTERPRISE"						     = "SharePoint Online (Plan 2)"
	"Intelligent_Content_Services"					     = "SharePoint Syntex"
	"MCOIMP"							     = "Skype for Business Online (Plan 1)"
	"MCOSTANDARD"							     = "Skype for Business Online (Plan 2)"
	"MCOPSTN2"							     = "Skype for Business PSTN Domestic and International Calling"
	"MCOPSTN1"							     = "Skype for Business PSTN Domestic Calling"
	"MCOPSTN5"							     = "Skype for Business PSTN Domestic Calling (120 Minutes)"
	"MCOPSTNPP"							     = "Skype for Business PSTN Usage Calling Plan"
	"MCOTEAMS_ESSENTIALS"						     = "Teams Phone with Calling Plan"
	"MTR_PREM"							     = "Teams Rooms Premium"
	"MCOPSTNEAU2"							     = "TELSTRA Calling for O365"
	"UNIVERSAL_PRINT"						     = "Universal Print"
	"VISIO_PLAN1_DEPT"						     = "Visio Plan 1"
	"VISIO_PLAN2_DEPT"						     = "Visio Plan 2"
	"VISIOONLINE_PLAN1"						     = "Visio Online Plan 1"
	"VISIOCLIENT"							     = "Visio Online Plan 2"
	"VISIOCLIENT_GOV"						     = "Visio Plan 2 for GCC"
	"TOPIC_EXPERIENCES"						     = "Viva Topics"
	"WIN_ENT_E5"							     = "Windows 10/11 Enterprise E5 (Original)"
	"WIN10_ENT_A3_FAC"						     = "Windows 10 Enterprise A3 for faculty"
	"WIN10_ENT_A3_STU"						     = "Windows 10 Enterprise A3 for students"
	"WIN10_PRO_ENT_SUB"						     = "Windows 10 Enterprise E3"
	"WIN10_VDA_E3"							     = "Windows 10 Enterprise E3"
	"WIN10_VDA_E5"							     = "Windows 10 Enterprise E5"
	"WINE5_GCC_COMPAT"						     = "Windows 10 Enterprise E5 Commercial (GCC Compatible)"
	"E3_VDA_only"							     = "Windows 10/11 Enterprise E3 VDA"
	"CPC_B_2C_4RAM_64GB"						     = "Windows 365 Business 2 vCPU 4 GB 64 GB"
	"CPC_B_4C_16RAM_128GB_WHB"					     = "Windows 365 Business 4 vCPU 16 GB 128 GB (with Windows Hybrid Benefit)"
	"CPC_E_2C_4GB_64GB"						     = "Windows 365 Enterprise 2 vCPU 4 GB 64 GB"
	"CPC_E_2C_8GB_128GB "						     = "Windows 365 Enterprise 2 vCPU, 8 GB, 128 GB "
	"CPC_LVL_2 "							     = "Windows 365 Enterprise 2 vCPU, 8 GB, 128 GB (Preview) "
	"CPC_LVL_3"							     = "Windows 365 Enterprise 4 vCPU, 16 GB, 256 GB (Preview) "
	"WINDOWS_STORE"							     = "Windows Store for Business"
	"WSFB_EDU_FACULTY "						     = "Windows Store for Business EDU Faculty"
	"WORKPLACE_ANALYTICS"						     = "Microsoft Workplace Analytics"
}
# Get all users right away. Instead of doing several lookups, we will use this object to look up all the information needed.
$AllUsers = get-azureaduser -All:$true -ErrorAction SilentlyContinue

Write-Host "Gathering Company Information..." -ForegroundColor Yellow
#Company Information
$CompanyInfo = Get-AzureADTenantDetail -ErrorAction SilentlyContinue

$CompanyName = $CompanyInfo.DisplayName
$TechEmail = $CompanyInfo.TechnicalNotificationMails | Out-String
$DirSync = $CompanyInfo.DirSyncEnabled
$LastDirSync = $CompanyInfo.CompanyLastDirSyncTime


If ($DirSync -eq $Null)
{
	$LastDirSync = "Not Available"
	$DirSync = "Disabled"
}
If ($PasswordSync -eq $Null)
{
	$LastPasswordSync = "Not Available"
}

$obj = [PSCustomObject]@{
	'Name'					   = $CompanyName
	'Technical E-mail'		   = $TechEmail
	'Directory Sync'		   = $DirSync
	'Last Directory Sync'	   = $LastDirSync
}

$CompanyInfoTable.add($obj)

Write-Host "Gathering Admin Roles and Members..." -ForegroundColor Yellow

Write-Host "Getting Tenant Global Admins" -ForegroundColor white
#Get Tenant Global Admins
$role = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -match "Company Administrator" } -ErrorAction SilentlyContinue
If ($null -ne $role)
{
	$Admins = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -ne "CloudConsoleGrapApi" }
	Foreach ($Admin in $Admins)
	{
		
		$MFAS = ((Get-MsolUser -objectid $Admin.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State
		
		if ($Null -ne $MFAS)
		{
			$MFASTATUS = "Enabled"
		}
		else
		{
			$MFASTATUS = "Disabled"
		}
		
		$Name = $Admin.DisplayName
		$EmailAddress = $Admin.Mail
		if (($admin.assignedlicenses.SkuID) -ne $Null)
		{
			$Licensed = $True
		}
		else
		{
			$Licensed = $False
		}
		$obj = [PSCustomObject]@{
			'Name'		     = $Name
			'MFA Status'	 = $MFAStatus
			'Is Licensed'    = $Licensed
			'E-Mail Address' = $EmailAddress
		}
		
		$GlobalAdminTable.add($obj)
	}
}



Write-Host "Getting Tenant Exchange Admins" -ForegroundColor white
#Get Tenant Exchange Admins
$exchrole = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -match "Exchange Service Administrator" } -ErrorAction SilentlyContinue
If ($Null -ne $exchrole)
{
	$ExchAdmins = Get-AzureADDirectoryRoleMember -ObjectId $exchrole.ObjectId -ErrorAction SilentlyContinue
	Foreach ($ExchAdmin in $ExchAdmins)
	{
		$MFAS = ((Get-MsolUser -objectid $ExchAdmin.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State
		
		if ($Null -ne $MFAS)
		{
			$MFASTATUS = "Enabled"
		}
		else
		{
			$MFASTATUS = "Disabled"
		}
		$Name = $ExchAdmin.DisplayName
		$EmailAddress = $ExchAdmin.Mail
		if (($Exchadmin.assignedlicenses.SkuID) -ne $Null)
		{
			$Licensed = $True
		}
		else
		{
			$Licensed = $False
		}
		
		$obj = [PSCustomObject]@{
			'Name'		     = $Name
			'MFA Status'	 = $MFAStatus
			'Is Licensed'    = $Licensed
			'E-Mail Address' = $EmailAddress
		}
		
		$ExchangeAdminTable.add($obj)
		
	}
}
If (($ExchangeAdminTable).count -eq 0)
{
	$ExchangeAdminTable = [PSCustomObject]@{
		'Information' = 'Information: No Users with the Exchange Administrator role were found, refer to the Global Administrators list.'
	}
}

Write-Host "Getting Tenant Privileged Admins" -ForegroundColor white
#Get Tenant Privileged Admins
$privadminrole = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -match "Privileged Role Administrator" } -ErrorAction SilentlyContinue
If ($Null -ne $privadminrole)
{
	$PrivAdmins = Get-AzureADDirectoryRoleMember -ObjectId $privadminrole.ObjectId -ErrorAction SilentlyContinue -ErrorVariable SilentlyContinue
	Foreach ($PrivAdmin in $PrivAdmins)
	{
		$MFAS = ((Get-MsolUser -objectid $PrivAdmin.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State
		
		if ($Null -ne $MFAS)
		{
			$MFASTATUS = "Enabled"
		}
		else
		{
			$MFASTATUS = "Disabled"
		}
		
		$Name = $PrivAdmin.DisplayName
		$EmailAddress = $PrivAdmin.Mail
		if (($admin.assignedlicenses.SkuID) -ne $Null)
		{
			$Licensed = $True
		}
		else
		{
			$Licensed = $False
		}
		
		$obj = [PSCustomObject]@{
			'Name'		     = $Name
			'MFA Status'	 = $MFAStatus
			'Is Licensed'    = $Licensed
			'E-Mail Address' = $EmailAddress
		}
		
		$PrivAdminTable.add($obj)
		
	}
}
If (($PrivAdminTable).count -eq 0)
{
	$PrivAdminTable = [PSCustomObject]@{
		'Information' = 'Information: No Users with the Privileged Administrator role were found, refer to the Global Administrators list.'
	}
}

Write-Host "Getting Tenant User Account Admins" -ForegroundColor white
#Get Tenant User Account Admins
$userrole = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -match "User Account Administrator" } -ErrorAction SilentlyContinue
If ($Null -ne $userrole)
{
	$userAdmins = Get-AzureADDirectoryRoleMember -ObjectId $userrole.ObjectId -ErrorAction SilentlyContinue
	Foreach ($userAdmin in $userAdmins)
	{
		$MFAS = ((Get-MsolUser -objectid $userAdmin.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State
		
		if ($Null -ne $MFAS)
		{
			$MFASTATUS = "Enabled"
		}
		else
		{
			$MFASTATUS = "Disabled"
		}
		$Name = $userAdmin.DisplayName
		$EmailAddress = $userAdmin.Mail
		if (($useradmin.assignedlicenses.SkuID) -ne $Null)
		{
			$Licensed = $True
		}
		else
		{
			$Licensed = $False
		}
		
		$obj = [PSCustomObject]@{
			'Name'		     = $Name
			'MFA Status'	 = $MFAStatus
			'Is Licensed'    = $Licensed
			'E-Mail Address' = $EmailAddress
		}
		
		$UserAdminTable.add($obj)
		
	}
}
If (($UserAdminTable).count -eq 0)
{
	$UserAdminTable = [PSCustomObject]@{
		'Information' = 'Information: No Users with the User Account Administrator role were found, refer to the Global Administrators list.'
	}
}

Write-Host "Getting Helpdesk Admins" -ForegroundColor white
#Get Tenant Tech Account Exchange Admins
$TechExchAdmins = Get-RoleGroupMember -Identity "Helpdesk Administrator" -ErrorAction SilentlyContinue
Foreach ($TechExchAdmin in $TechExchAdmins)
{
	$AccountInfo = Get-MsolUser -searchstring $TechExchAdmin.Name -ErrorAction SilentlyContinue
	$Name = $AccountInfo.DisplayName

    $MFAS = ((Get-MsolUser -objectid $AccountInfo.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State

			if ($Null -ne $MFAS)
			{
				$MFASTATUS = "Enabled"
			}
			else
			{
				$MFASTATUS = "Disabled"
			}
	$EmailAddress = $AccountInfo.UserPrincipalName
	if (($AccountInfo.assignedlicenses.SkuID) -ne $Null)
	{
		$Licensed = $True
	}
	else
	{
		$Licensed = $False
	}
	
	$obj = [PSCustomObject]@{
		'Name'			      = $Name
		'MFA Status'		  = $MFAStatus
		'Is Licensed'		  = $Licensed
		'E-Mail Address'	  = $EmailAddress
	}
	
	$TechExchAdminTable.add($obj)
	
}
If (($TechExchAdminTable).count -eq 0)
{
	$TechExchAdminTable = [PSCustomObject]@{
		'Information'  = 'Information: No Users with the Helpdesk Administrator role were found, refer to the Global Administrators list.'
	}
}

Write-Host "Getting Tenant SharePoint Admins" -ForegroundColor white
#Get Tenant SharePoint Admins
$sprole = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -match "SharePoint Service Administrator" } -ErrorAction SilentlyContinue
If ($Null -ne $sprole)
{
	$SPAdmins = Get-AzureADDirectoryRoleMember -ObjectId $sprole.ObjectId -ErrorAction SilentlyContinue
	Foreach ($SPAdmin in $SPAdmins)
	{
		$MFAS = ((Get-MsolUser -objectid $SPAdmin.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State
		
		if ($Null -ne $MFAS)
		{
			$MFASTATUS = "Enabled"
		}
		else
		{
			$MFASTATUS = "Disabled"
		}
		$Name = $SPAdmin.DisplayName
		$EmailAddress = $SPAdmin.Mail
		if (($SPadmin.assignedlicenses.SkuID) -ne $Null)
		{
			$Licensed = $True
		}
		else
		{
			$Licensed = $False
		}
		
		$obj = [PSCustomObject]@{
			'Name'		     = $Name
			'MFA Status'	 = $MFAStatus
			'Is Licensed'    = $Licensed
			'E-Mail Address' = $EmailAddress
		}
		
		$SharePointAdminTable.add($obj)
		
	}
}
If (($SharePointAdminTable).count -eq 0)
{
	$SharePointAdminTable = [PSCustomObject]@{
		'Information' = 'Information: No Users with the SharePoint Service Administrator role were found, refer to the Global Administrators list.'
	}
}

Write-Host "Getting Tenant Skype Admins" -ForegroundColor white
#Get Tenant Skype Admins
$skyperole = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -match "Lync Service Administrator" } -ErrorAction SilentlyContinue
If ($Null -ne $skyperole)
{
	$skypeAdmins = Get-AzureADDirectoryRoleMember -ObjectId $skyperole.ObjectId -ErrorAction SilentlyContinue
	Foreach ($skypeAdmin in $skypeAdmins)
	{
		$MFAS = ((Get-MsolUser -objectid $skypeAdmin.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State
		
		if ($Null -ne $MFAS)
		{
			$MFASTATUS = "Enabled"
		}
		else
		{
			$MFASTATUS = "Disabled"
		}
		$Name = $skypeAdmin.DisplayName
		$EmailAddress = $skypeAdmin.Mail
		if (($skypeadmin.assignedlicenses.SkuID) -ne $Null)
		{
			$Licensed = $True
		}
		else
		{
			$Licensed = $False
		}
		
		$obj = [PSCustomObject]@{
			'Name'		     = $Name
			'MFA Status'	 = $MFAStatus
			'Is Licensed'    = $Licensed
			'E-Mail Address' = $EmailAddress
		}
		
		$SkypeAdminTable.add($obj)
		
	}
}
If (($skypeAdminTable).count -eq 0)
{
	$skypeAdminTable = [PSCustomObject]@{
		'Information' = 'Information: No Users with the Lync Service Administrator role were found, refer to the Global Administrators list.'
	}
}

Write-Host "Getting Tenant CRM Admins" -ForegroundColor white
#Get Tenant CRM Admins
$crmrole = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -match "CRM Service Administrator" } -ErrorAction SilentlyContinue
If ($Null -ne $crmrole)
{
	$crmAdmins = Get-AzureADDirectoryRoleMember -ObjectId $crmrole.ObjectId -ErrorAction SilentlyContinue
	Foreach ($crmAdmin in $crmAdmins)
	{
		$MFAS = ((Get-MsolUser -objectid $crmAdmin.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State
		
		if ($Null -ne $MFAS)
		{
			$MFASTATUS = "Enabled"
		}
		else
		{
			$MFASTATUS = "Disabled"
		}
		$Name = $crmAdmin.DisplayName
		$EmailAddress = $crmAdmin.Mail
		if (($crmadmin.assignedlicenses.SkuID) -ne $Null)
		{
			$Licensed = $True
		}
		else
		{
			$Licensed = $False
		}
		
		$obj = [PSCustomObject]@{
			'Name'		     = $Name
			'MFA Status'	 = $MFAStatus
			'Is Licensed'    = $Licensed
			'E-Mail Address' = $EmailAddress
		}
		
		$CRMAdminTable.add($obj)
		
	}
}
If (($CRMAdminTable).count -eq 0)
{
	$CRMAdminTable = [PSCustomObject]@{
		'Information' = 'Information: No Users with the CRM Service Administrator role were found, refer to the Global Administrators list.'
	}
}

Write-Host "Getting Tenant Power BI Admins" -ForegroundColor white
#Get Tenant Power BI Admins
$birole = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -match "Power BI Service Administrator" } -ErrorAction SilentlyContinue
If ($null -ne $birole)
{
	$biAdmins = Get-AzureADDirectoryRoleMember -ObjectId $birole.ObjectId -ErrorAction SilentlyContinue
	
	Foreach ($biAdmin in $biAdmins)
	{
		$MFAS = ((Get-MsolUser -objectid $biAdmin.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State
		
		if ($Null -ne $MFAS)
		{
			$MFASTATUS = "Enabled"
		}
		else
		{
			$MFASTATUS = "Disabled"
		}
		$Name = $biAdmin.DisplayName
		$EmailAddress = $biAdmin.Mail
		if (($biadmin.assignedlicenses.SkuID) -ne $Null)
		{
			$Licensed = $True
		}
		else
		{
			$Licensed = $False
		}
		
		$obj = [PSCustomObject]@{
			'Name'		     = $Name
			'MFA Status'	 = $MFAStatus
			'Is Licensed'    = $Licensed
			'E-Mail Address' = $EmailAddress
		}
		
		$PowerBIAdminTable.add($obj)
		
	}
}
If (($PowerBIAdminTable).count -eq 0)
{
	$PowerBIAdminTable = [PSCustomObject]@{
		'Information' = 'Information: No Users with the Power BI Administrator role were found, refer to the Global Administrators list.'
	}
}

Write-Host "Getting Tenant Service Support Admins" -ForegroundColor white
#Get Tenant Service Support Admins
$servicerole = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -match "Service Support Administrator" } -ErrorAction SilentlyContinue
If ($Null -ne $servicerole)
{
	$serviceAdmins = Get-AzureADDirectoryRoleMember -ObjectId $servicerole.ObjectId -ErrorAction SilentlyContinue
	Foreach ($serviceAdmin in $serviceAdmins)
	{
		$MFAS = ((Get-MsolUser -objectid $serviceAdmin.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State
		
		if ($Null -ne $MFAS)
		{
			$MFASTATUS = "Enabled"
		}
		else
		{
			$MFASTATUS = "Disabled"
		}
		$Name = $serviceAdmin.DisplayName
		$EmailAddress = $serviceAdmin.Mail
		if (($serviceadmin.assignedlicenses.SkuID) -ne $Null)
		{
			$Licensed = $True
		}
		else
		{
			$Licensed = $False
		}
		
		$obj = [PSCustomObject]@{
			'Name'		     = $Name
			'MFA Status'	 = $MFAStatus
			'Is Licensed'    = $Licensed
			'E-Mail Address' = $EmailAddress
		}
		
		$ServiceAdminTable.add($obj)
		
	}
}
If (($serviceAdminTable).count -eq 0)
{
	$serviceAdminTable = [PSCustomObject]@{
		'Information' = 'Information: No Users with the Service Support Administrator role were found, refer to the Global Administrators list.'
	}
}

Write-Host "Getting Tenant Billing Admins" -ForegroundColor white
#Get Tenant Billing Admins
$billingrole = Get-AzureADDirectoryRole | Where-Object { $_.DisplayName -match "Billing Administrator" } -ErrorAction SilentlyContinue
If ($Null -ne $billingrole)
{
	$billingAdmins = Get-AzureADDirectoryRoleMember -ObjectId $billingrole.ObjectId -ErrorAction SilentlyContinue
	Foreach ($billingAdmin in $billingAdmins)
	{
		$MFAS = ((Get-MsolUser -objectid $billingAdmin.ObjectID -ErrorAction SilentlyContinue).StrongAuthenticationRequirements).State
		
		if ($Null -ne $MFAS)
		{
			$MFASTATUS = "Enabled"
		}
		else
		{
			$MFASTATUS = "Disabled"
		}
		$Name = $billingAdmin.DisplayName
		$EmailAddress = $billingAdmin.Mail
		if (($billingadmin.assignedlicenses.SkuID) -ne $Null)
		{
			$Licensed = $True
		}
		else
		{
			$Licensed = $False
		}
		
		$obj = [PSCustomObject]@{
			'Name'		     = $Name
			'MFA Status'	 = $MFAStatus
			'Is Licensed'    = $Licensed
			'E-Mail Address' = $EmailAddress
		}
		
		$BillingAdminTable.add($obj)
		
	}
}
If (($billingAdminTable).count -eq 0)
{
	$billingAdminTable = [PSCustomObject]@{
		'Information' = 'Information: No Users with the Billing Administrator role were found, refer to the Global Administrators list.'
	}
}

Write-Host "Getting Users with Strong Password Disabled..." -ForegroundColor Yellow
#Users with Strong Password Disabled
$LooseUsers = $AllUsers | Where-Object { $_.PasswordPolicies -eq "DisableStrongPassword" }
Foreach ($LooseUser in $LooseUsers)
{
	$NameLoose = $LooseUser.DisplayName
	$UPNLoose = $LooseUser.UserPrincipalName
	$StrongPasswordLoose = "False"
	if (($LooseUser.assignedlicenses.SkuID) -ne $Null)
	{
		$LicensedLoose = $true
	}
	else
	{
		$LicensedLoose = $false
	}
	
	$obj = [PSCustomObject]@{
		'Name'						    = $NameLoose
		'UserPrincipalName'			    = $UPNLoose
		'Is Licensed'				    = $LicensedLoose
		'Strong Password Required'	    = $StrongPasswordLoose
	}
	
	
	$StrongPasswordTable.add($obj)
}
If (($StrongPasswordTable).count -eq 0)
{
	$StrongPasswordTable = [PSCustomObject]@{
		'Information'  = 'Information: No Users were found with Strong Password Enforcement disabled'
	}
}

Write-Host "Getting Tenant Domains..." -ForegroundColor Yellow
#Tenant Domain
$Domains = Get-AzureAdDomain
foreach ($Domain in $Domains)
{
	$DomainName = $Domain.Name
	$Verified = $Domain.IsVerified
	$DefaultStatus = $Domain.IsDefault
	
	$obj = [PSCustomObject]@{
		'Domain Name'				  = $DomainName
		'Verification Status'		  = $Verified
		'Default'				      = $DefaultStatus
	}
	
	$DomainTable.add($obj)
}

Write-Host "Getting Groups..." -ForegroundColor Yellow
#Get groups and sort in alphabetical order
$Groups = Get-AzureAdGroup -All $True | Sort-Object DisplayName
$365GroupCount = ($Groups | Where-Object { $_.MailEnabled -eq $true -and $_.DirSyncEnabled -eq $null -and $_.SecurityEnabled -eq $false }).Count
$obj1 = [PSCustomObject]@{
	'Name'					      = 'Office 365 Group'
	'Count'					      = $365GroupCount
}

$GroupTypetable.add($obj1)

Write-Host "Getting Distribution Groups..." -ForegroundColor White
$DistroCount = ($Groups | Where-Object { $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $false }).Count
$obj1 = [PSCustomObject]@{
	'Name'					      = 'Distribution List'
	'Count'					      = $DistroCount
}

$GroupTypetable.add($obj1)

Write-Host "Getting Security Groups..." -ForegroundColor White
$SecurityCount = ($Groups | Where-Object { $_.MailEnabled -eq $false -and $_.SecurityEnabled -eq $true }).Count
$obj1 = [PSCustomObject]@{
	'Name'					      = 'Security Group'
	'Count'					      = $SecurityCount
}

$GroupTypetable.add($obj1)

Write-Host "Getting Mail-Enabled Security Groups..." -ForegroundColor White
$SecurityMailEnabledCount = ($Groups | Where-Object { $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $true }).Count
$obj1 = [PSCustomObject]@{
	'Name'					      = 'Mail Enabled Security Group'
	'Count'					      = $SecurityMailEnabledCount
}

$GroupTypetable.add($obj1)

Foreach ($Group in $Groups)
{
	$Type = New-Object 'System.Collections.Generic.List[System.Object]'
	
	if ($group.MailEnabled -eq $True -and $group.DirSyncEnabled -eq $null -and $group.SecurityEnabled -eq $False)
	{
		$Type = "Office 365 Group"
	}
	if ($group.MailEnabled -eq $True -and $group.SecurityEnabled -eq $False)
	{
		$Type = "Distribution List"
	}
	if ($group.MailEnabled -eq $False -and $group.SecurityEnabled -eq $True)
	{
		$Type = "Security Group"
	}
	if ($group.MailEnabled -eq $True -and $group.SecurityEnabled -eq $True)
	{
		$Type = "Mail Enabled Security Group"
	}
	
	$Users = (Get-AzureADGroupMember -ObjectId $Group.ObjectID | Sort-Object DisplayName | Select-Object -ExpandProperty DisplayName) -join ", "
	$GName = $Group.DisplayName
	
	$hash = New-Object PSObject -property @{ Name = "$GName"; Type = "$Type"; Members = "$Users" }
	$GEmail = $Group.Mail
	
	
	$obj = [PSCustomObject]@{
		'Name'				   = $GName
		'Type'				   = $Type
		'Members'			   = $users
		'E-mail Address'	   = $GEmail
	}
	
	$table.add($obj)
}
If (($table).count -eq 0)
{
	$table = [PSCustomObject]@{
		'Information'  = 'Information: No Groups were found in the tenant'
	}
}


Write-Host "Getting Licenses..." -ForegroundColor Yellow
#Get all licenses
$Licenses = Get-AzureADSubscribedSku
#Split licenses at colon
Foreach ($License in $Licenses)
{
	$TextLic = $null
	
	$ASku = ($License).SkuPartNumber
	$TextLic = $Sku.Item("$ASku")
	If (!($TextLic))
	{
		$OLicense = $License.SkuPartNumber
	}
	Else
	{
		$OLicense = $TextLic
	}
	
	$TotalAmount = $License.PrepaidUnits.enabled
	$Assigned = $License.ConsumedUnits
	$Unassigned = ($TotalAmount - $Assigned)

	If ($TotalAmount -lt $LicenseFilter)
	{
		$obj = [PSCustomObject]@{
			'Name'					    = $Olicense
			'Total Amount'			    = $TotalAmount
			'Assigned Licenses'		    = $Assigned
			'Unassigned Licenses'	    = $Unassigned
		}
		
		$licensetable.add($obj)
	}
}
If (($licensetable).count -eq 0)
{
	$licensetable = [PSCustomObject]@{
		'Information'  = 'Information: No Licenses were found in the tenant'
	}
}


$IsLicensed = ($AllUsers | Where-Object { $_.assignedlicenses.count -gt 0 }).Count
$objULic = [PSCustomObject]@{
	'Name'	   = 'Users Licensed'
	'Count'    = $IsLicensed
}

$IsLicensedUsersTable.add($objULic)

$ISNotLicensed = ($AllUsers | Where-Object { $_.assignedlicenses.count -eq 0 }).Count
$objULic = [PSCustomObject]@{
	'Name'	   = 'Users Not Licensed'
	'Count'    = $IsNotLicensed
}

$IsLicensedUsersTable.add($objULic)
If (($IsLicensedUsersTable).count -eq 0)
{
	$IsLicensedUsersTable = [PSCustomObject]@{
		'Information'  = 'Information: No Licenses were found in the tenant'
	}
}

Write-Host "Getting Users..." -ForegroundColor Yellow
Foreach ($User in $AllUsers)
{
	$ProxyA = New-Object 'System.Collections.Generic.List[System.Object]'
	$NewObject02 = New-Object 'System.Collections.Generic.List[System.Object]'
	$NewObject01 = New-Object 'System.Collections.Generic.List[System.Object]'
    $UserLicenses = ($user | Select -ExpandProperty AssignedLicenses).SkuID
	If (($UserLicenses).count -gt 1)
	{
	$LastLogon = Get-MailboxStatistics $User.DisplayName | Select-Object -ExpandProperty LastLogonTime
		Foreach ($UserLicense in $UserLicenses)
		{
            $UserLicense = ($licenses | Where-Object { $_.skuid -match $UserLicense }).SkuPartNumber
			$TextLic = $Sku.Item("$UserLicense")
			If (!($TextLic))
			{
				$NewObject01 = [PSCustomObject]@{
					'Licenses'	   = $UserLicense
				}
				$NewObject02.add($NewObject01)
			}
			Else
			{
				$NewObject01 = [PSCustomObject]@{
					'Licenses'	   = $textlic
				}
				
				$NewObject02.add($NewObject01)
			}
		}
	}
	Elseif (($UserLicenses).count -eq 1)
	{
	$LastLogon = Get-MailboxStatistics $User.DisplayName | Select-Object -ExpandProperty LastLogonTime
		$lic = ($licenses | Where-Object { $_.skuid -match $UserLicenses}).SkuPartNumber
		$TextLic = $Sku.Item("$lic")
		If (!($TextLic))
		{
			$NewObject01 = [PSCustomObject]@{
				'Licenses'	   = $lic
			}
			$NewObject02.add($NewObject01)
		}
		Else
		{
			$NewObject01 = [PSCustomObject]@{
				'Licenses'	   = $textlic
			}
			$NewObject02.add($NewObject01)
		}
	}
	Else
	{
	$LastLogon = $Null
		$NewObject01 = [PSCustomObject]@{
			'Licenses'	   = $Null
		}
		$NewObject02.add($NewObject01)
	}
	
	$ProxyAddresses = ($User | Select-Object -ExpandProperty ProxyAddresses)
	If ($ProxyAddresses -ne $Null)
	{
		Foreach ($Proxy in $ProxyAddresses)
		{
			$ProxyB = $Proxy -split ":" | Select-Object -Last 1
			$ProxyA.add($ProxyB)
			
		}
		$ProxyC = $ProxyA -join ", "
	}
	Else
	{
		$ProxyC = $Null
	}
	
	$Name = $User.DisplayName
	$UPN = $User.UserPrincipalName
	$UserLicenses = ($NewObject02 | Select-Object -ExpandProperty Licenses) -join ", "
	$Enabled = $User.AccountEnabled
	$ResetPW = Get-User $User.DisplayName | Select-Object -ExpandProperty ResetPasswordOnNextLogon 
	
 $obj = [PSCustomObject]@{
		    'Name'				                   = $Name
		    'UserPrincipalName'	                   = $UPN
		    'Licenses'			                   = $UserLicenses
            'Last Mailbox Logon'                   = $LastLogon
		    'Reset Password at Next Logon'         = $ResetPW
		    'Enabled'			                   = $Enabled
		    'E-mail Addresses'	                   = $ProxyC
	    }
	
	$usertable.add($obj)
}
If (($usertable).count -eq 0)
{
	$usertable = [PSCustomObject]@{
		'Information'  = 'Information: No Users were found in the tenant'
	}
}

Write-Host "Getting Shared Mailboxes..." -ForegroundColor Yellow
#Get all Shared Mailboxes
$SharedMailboxes = Get-Recipient -Resultsize unlimited | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }
Foreach ($SharedMailbox in $SharedMailboxes)
{
	$ProxyA = New-Object 'System.Collections.Generic.List[System.Object]'
	$Name = $SharedMailbox.Name
	$PrimEmail = $SharedMailbox.PrimarySmtpAddress
	$ProxyAddresses = ($SharedMailbox | Where-Object { $_.EmailAddresses -notlike "*$PrimEmail*" } | Select-Object -ExpandProperty EmailAddresses)
	If ($ProxyAddresses -ne $Null)
	{
		Foreach ($ProxyAddress in $ProxyAddresses)
		{
			$ProxyB = $ProxyAddress -split ":" | Select-Object -Last 1
			If ($ProxyB -eq $PrimEmail)
			{
				$ProxyB = $Null
			}
			$ProxyA.add($ProxyB)
			$ProxyC = $ProxyA
		}
	}
	Else
	{
		$ProxyC = $Null
	}
	
	$ProxyF = ($ProxyC -join ", ").TrimEnd(", ")
	
	$obj = [PSCustomObject]@{
		'Name'				   = $Name
		'Primary E-Mail'	   = $PrimEmail
		'E-mail Addresses'	   = $ProxyF
	}
	
	
	
	$SharedMailboxTable.add($obj)
	
}
If (($SharedMailboxTable).count -eq 0)
{
	$SharedMailboxTable = [PSCustomObject]@{
		'Information'  = 'Information: No Shared Mailboxes were found in the tenant'
	}
}

Write-Host "Getting Contacts..." -ForegroundColor Yellow
#Get all Contacts
$Contacts = Get-MailContact
#Split licenses at colon
Foreach ($Contact in $Contacts)
{
	
	$ContactName = $Contact.DisplayName
	$ContactPrimEmail = $Contact.PrimarySmtpAddress
	
	$objContact = [PSCustomObject]@{
		'Name'			     = $ContactName
		'E-mail Address'	 = $ContactPrimEmail
	}
	
	$ContactTable.add($objContact)
	
}
If (($ContactTable).count -eq 0)
{
	$ContactTable = [PSCustomObject]@{
		'Information'  = 'Information: No Contacts were found in the tenant'
	}
}

Write-Host "Getting Mail Users..." -ForegroundColor Yellow
#Get all Mail Users
$MailUsers = Get-MailUser
foreach ($MailUser in $mailUsers)
{
	$MailArray = New-Object 'System.Collections.Generic.List[System.Object]'
	$MailPrimEmail = $MailUser.PrimarySmtpAddress
	$MailName = $MailUser.DisplayName
	$MailEmailAddresses = ($MailUser.EmailAddresses | Where-Object { $_ -cnotmatch '^SMTP' })
	foreach ($MailEmailAddress in $MailEmailAddresses)
	{
		$MailEmailAddressSplit = $MailEmailAddress -split ":" | Select-Object -Last 1
		$MailArray.add($MailEmailAddressSplit)
		
		
	}
	
	$UserEmails = $MailArray -join ", "
	
	$obj = [PSCustomObject]@{
		'Name'				   = $MailName
		'Primary E-Mail'	   = $MailPrimEmail
		'E-mail Addresses'	   = $UserEmails
	}
	
	$ContactMailUserTable.add($obj)
}
If (($ContactMailUserTable).count -eq 0)
{
	$ContactMailUserTable = [PSCustomObject]@{
		'Information'  = 'Information: No Mail Users were found in the tenant'
	}
}

Write-Host "Getting Room Mailboxes..." -ForegroundColor Yellow
$Rooms = Get-Mailbox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "RoomMailBox")'
Foreach ($Room in $Rooms)
{
	$RoomArray = New-Object 'System.Collections.Generic.List[System.Object]'
	
	$RoomName = $Room.DisplayName
	$RoomPrimEmail = $Room.PrimarySmtpAddress
	$RoomEmails = ($Room.EmailAddresses | Where-Object { $_ -cnotmatch '^SMTP' })
	foreach ($RoomEmail in $RoomEmails)
	{
		$RoomEmailSplit = $RoomEmail -split ":" | Select-Object -Last 1
		$RoomArray.add($RoomEmailSplit)
	}
	$RoomEMailsF = $RoomArray -join ", "
	
	
	$obj = [PSCustomObject]@{
		'Name'				   = $RoomName
		'Primary E-Mail'	   = $RoomPrimEmail
		'E-mail Addresses'	   = $RoomEmailsF
	}
	
	$RoomTable.add($obj)
}
If (($RoomTable).count -eq 0)
{
	$RoomTable = [PSCustomObject]@{
		'Information'  = 'Information: No Room Mailboxes were found in the tenant'
	}
}

Write-Host "Getting Equipment Mailboxes..." -ForegroundColor Yellow
$EquipMailboxes = Get-Mailbox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "EquipmentMailBox")'
Foreach ($EquipMailbox in $EquipMailboxes)
{
	$EquipArray = New-Object 'System.Collections.Generic.List[System.Object]'
	
	$EquipName = $EquipMailbox.DisplayName
	$EquipPrimEmail = $EquipMailbox.PrimarySmtpAddress
	$EquipEmails = ($EquipMailbox.EmailAddresses | Where-Object { $_ -cnotmatch '^SMTP' })
	foreach ($EquipEmail in $EquipEmails)
	{
		$EquipEmailSplit = $EquipEmail -split ":" | Select-Object -Last 1
		$EquipArray.add($EquipEmailSplit)
	}
	$EquipEMailsF = $EquipArray -join ", "
	
	$obj = [PSCustomObject]@{
		'Name'				   = $EquipName
		'Primary E-Mail'	   = $EquipPrimEmail
		'E-mail Addresses'	   = $EquipEmailsF
	}
	
	
	$EquipTable.add($obj)
}
If (($EquipTable).count -eq 0)
{
	$EquipTable = [PSCustomObject]@{
		'Information'  = 'Information: No Equipment Mailboxes were found in the tenant'
	}
}

Write-Host "Generating HTML Report..." -ForegroundColor Yellow

$tabarray = @('Dashboard', 'Admins', 'Users', 'Groups', 'Licenses', 'Shared Mailboxes', 'Contacts', 'Resources')

#basic Properties 
$PieObject2 = Get-HTMLPieChartObject
$PieObject2.Title = "Office 365 Total Licenses"
$PieObject2.Size.Height = 500
$PieObject2.Size.width = 500
$PieObject2.ChartStyle.ChartType = 'doughnut'

#These file exist in the module directoy, There are 4 schemes by default
$PieObject2.ChartStyle.ColorSchemeName = "ColorScheme4"
#There are 8 generated schemes, randomly generated at runtime 
$PieObject2.ChartStyle.ColorSchemeName = "Generated7"
#you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObject2.ChartStyle.ColorSchemeName = 'Random'

#Data defintion you can reference any column from name and value from the  dataset.  
#Name and Count are the default to work with the Group function.
$PieObject2.DataDefinition.DataNameColumnName = 'Name'
$PieObject2.DataDefinition.DataValueColumnName = 'Total Amount'

#basic Properties 
$PieObject3 = Get-HTMLPieChartObject
$PieObject3.Title = "Office 365 Assigned Licenses"
$PieObject3.Size.Height = 500
$PieObject3.Size.width = 500
$PieObject3.ChartStyle.ChartType = 'doughnut'

#These file exist in the module directoy, There are 4 schemes by default
$PieObject3.ChartStyle.ColorSchemeName = "ColorScheme4"
#There are 8 generated schemes, randomly generated at runtime 
$PieObject3.ChartStyle.ColorSchemeName = "Generated5"
#you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObject3.ChartStyle.ColorSchemeName = 'Random'

#Data defintion you can reference any column from name and value from the  dataset.  
#Name and Count are the default to work with the Group function.
$PieObject3.DataDefinition.DataNameColumnName = 'Name'
$PieObject3.DataDefinition.DataValueColumnName = 'Assigned Licenses'

#basic Properties 
$PieObject4 = Get-HTMLPieChartObject
$PieObject4.Title = "Office 365 Unassigned Licenses"
$PieObject4.Size.Height = 250
$PieObject4.Size.width = 250
$PieObject4.ChartStyle.ChartType = 'doughnut'

#These file exist in the module directoy, There are 4 schemes by default
$PieObject4.ChartStyle.ColorSchemeName = "ColorScheme4"
#There are 8 generated schemes, randomly generated at runtime 
$PieObject4.ChartStyle.ColorSchemeName = "Generated4"
#you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObject4.ChartStyle.ColorSchemeName = 'Random'

#Data defintion you can reference any column from name and value from the  dataset.  
#Name and Count are the default to work with the Group function.
$PieObject4.DataDefinition.DataNameColumnName = 'Name'
$PieObject4.DataDefinition.DataValueColumnName = 'Unassigned Licenses'

#basic Properties 
$PieObjectGroupType = Get-HTMLPieChartObject
$PieObjectGroupType.Title = "Office 365 Groups"
$PieObjectGroupType.Size.Height = 250
$PieObjectGroupType.Size.width = 250
$PieObjectGroupType.ChartStyle.ChartType = 'doughnut'

#These file exist in the module directoy, There are 4 schemes by default
$PieObjectGroupType.ChartStyle.ColorSchemeName = "ColorScheme4"
#There are 8 generated schemes, randomly generated at runtime 
$PieObjectGroupType.ChartStyle.ColorSchemeName = "Generated8"
#you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectGroupType.ChartStyle.ColorSchemeName = 'Random'

#Data defintion you can reference any column from name and value from the  dataset.  
#Name and Count are the default to work with the Group function.
$PieObjectGroupType.DataDefinition.DataNameColumnName = 'Name'
$PieObjectGroupType.DataDefinition.DataValueColumnName = 'Count'

##--LICENSED AND UNLICENSED USERS PIE CHART--##
#basic Properties 
$PieObjectULicense = Get-HTMLPieChartObject
$PieObjectULicense.Title = "License Status"
$PieObjectULicense.Size.Height = 250
$PieObjectULicense.Size.width = 250
$PieObjectULicense.ChartStyle.ChartType = 'doughnut'

#These file exist in the module directoy, There are 4 schemes by default
$PieObjectULicense.ChartStyle.ColorSchemeName = "ColorScheme3"
#There are 8 generated schemes, randomly generated at runtime 
$PieObjectULicense.ChartStyle.ColorSchemeName = "Generated3"
#you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectULicense.ChartStyle.ColorSchemeName = 'Random'

#Data defintion you can reference any column from name and value from the  dataset.  
#Name and Count are the default to work with the Group function.
$PieObjectULicense.DataDefinition.DataNameColumnName = 'Name'
$PieObjectULicense.DataDefinition.DataValueColumnName = 'Count'

$rpt = New-Object 'System.Collections.Generic.List[System.Object]'
$rpt += get-htmlopenpage -TitleText 'Office 365 Tenant Report' -LeftLogoString $CompanyLogo 

$rpt += Get-HTMLTabHeader -TabNames $tabarray 
    $rpt += get-htmltabcontentopen -TabName $tabarray[0] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
        $rpt+= Get-HtmlContentOpen -HeaderText "Office 365 Dashboard"
          $rpt += Get-HTMLContentOpen -HeaderText "Company Information"
            $rpt += Get-HtmlContentTable $CompanyInfoTable 
          $rpt += Get-HTMLContentClose

	        $rpt+= get-HtmlColumn1of2
		        $rpt+= Get-HtmlContentOpen -BackgroundShade 1 -HeaderText 'Global Administrators'
			        $rpt+= get-htmlcontentdatatable  $GlobalAdminTable -HideFooter
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose
	            $rpt+= get-htmlColumn2of2
		            $rpt+= Get-HtmlContentOpen -HeaderText 'Users With Strong Password Enforcement Disabled'
			            $rpt+= get-htmlcontentdatatable $StrongPasswordTable -HideFooter 
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose

          $rpt += Get-HTMLContentOpen -HeaderText "Domains"
            $rpt += Get-HtmlContentTable $DomainTable 
          $rpt += Get-HTMLContentClose

        $rpt+= Get-HtmlContentClose 
    $rpt += get-htmltabcontentclose
	
	    $rpt += get-htmltabcontentopen -TabName $tabarray[1] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
        $rpt+= Get-HtmlContentOpen -HeaderText "Role Assignments"
	       
		   	$rpt+= get-HtmlColumn1of2
		        $rpt+= Get-HtmlContentOpen -BackgroundShade 1 -HeaderText 'Privileged Role Administrators'
			        $rpt+= get-htmlcontentdatatable  $PrivAdminTable -HideFooter
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose
	            $rpt+= get-htmlColumn2of2
		            $rpt+= Get-HtmlContentOpen -HeaderText 'Exchange Administrators'
			            $rpt+= get-htmlcontentdatatable $ExchangeAdminTable -HideFooter 
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose
			
		   $rpt+= get-HtmlColumn1of2
		        $rpt+= Get-HtmlContentOpen -BackgroundShade 1 -HeaderText 'User Account Administrators'
			        $rpt+= get-htmlcontentdatatable  $UserAdminTable -HideFooter
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose
	            $rpt+= get-htmlColumn2of2
		            $rpt+= Get-HtmlContentOpen -HeaderText 'Helpdesk Administrators'
			            $rpt+= get-htmlcontentdatatable $TechExchAdminTable -HideFooter 
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose
			
		   $rpt+= get-HtmlColumn1of2
		        $rpt+= Get-HtmlContentOpen -BackgroundShade 1 -HeaderText 'SharePoint Administrators'
			        $rpt+= get-htmlcontentdatatable  $SharePointAdminTable -HideFooter
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose
	            $rpt+= get-htmlColumn2of2
		            $rpt+= Get-HtmlContentOpen -HeaderText 'Skype Administrators'
			            $rpt+= get-htmlcontentdatatable $SkypeAdminTable -HideFooter 
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose

		   $rpt+= get-HtmlColumn1of2
		        $rpt+= Get-HtmlContentOpen -BackgroundShade 1 -HeaderText 'CRM Service Administrators'
			        $rpt+= get-htmlcontentdatatable  $CRMAdminTable -HideFooter
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose
	            $rpt+= get-htmlColumn2of2
		            $rpt+= Get-HtmlContentOpen -HeaderText 'Power BI Administrators'
			            $rpt+= get-htmlcontentdatatable $PowerBIAdminTable -HideFooter 
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose
			
		   $rpt+= get-HtmlColumn1of2
		        $rpt+= Get-HtmlContentOpen -BackgroundShade 1 -HeaderText 'Service Support Administrators'
			        $rpt+= get-htmlcontentdatatable  $ServiceAdminTable -HideFooter
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose
	            $rpt+= get-htmlColumn2of2
		            $rpt+= Get-HtmlContentOpen -HeaderText 'Billing Administrators'
			            $rpt+= get-htmlcontentdatatable $BillingAdminTable -HideFooter 
		        $rpt+= Get-HtmlContentClose
	        $rpt+= get-htmlColumnClose
        $rpt+= Get-HtmlContentClose 
    $rpt += get-htmltabcontentclose
	
	    $rpt += get-htmltabcontentopen -TabName $tabarray[2] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
        $rpt += Get-HTMLContentOpen -HeaderText "Office 365 Users"
            $rpt += get-htmlcontentdatatable $UserTable -HideFooter
        $rpt += Get-HTMLContentClose
        $rpt += Get-HTMLContentOpen -HeaderText "Licensed & Unlicensed Users Chart"
		    $rpt += Get-HTMLPieChart -ChartObject $PieObjectULicense -DataSet $IsLicensedUsersTable
	    $rpt += Get-HTMLContentClose
    $rpt += get-htmltabcontentclose
	
    $rpt += get-htmltabcontentopen -TabName $tabarray[3] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
        $rpt += Get-HTMLContentOpen -HeaderText "Office 365 Groups"
            $rpt += get-htmlcontentdatatable $Table -HideFooter
        $rpt += Get-HTMLContentClose
        $rpt += Get-HTMLContentOpen -HeaderText "Office 365 Groups Chart"
		    $rpt += Get-HTMLPieChart -ChartObject $PieObjectGroupType -DataSet $GroupTypetable
	    $rpt += Get-HTMLContentClose
    $rpt += get-htmltabcontentclose
	
    $rpt += get-htmltabcontentopen -TabName $tabarray[4]  -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
        $rpt += Get-HTMLContentOpen -HeaderText "Office 365 Licenses"
            $rpt += get-htmlcontentdatatable $LicenseTable -HideFooter
        $rpt += Get-HTMLContentClose
	$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Licensing Charts"
	    $rpt += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2
	        $rpt += Get-HTMLPieChart -ChartObject $PieObject2 -DataSet $licensetable
	    $rpt += Get-HTMLColumnClose
	    $rpt += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2
	        $rpt += Get-HTMLPieChart -ChartObject $PieObject3 -DataSet $licensetable
	    $rpt += Get-HTMLColumnClose
    $rpt += Get-HTMLContentclose
    $rpt += get-htmltabcontentclose

    $rpt += get-htmltabcontentopen -TabName $tabarray[5] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy)) 
        $rpt += Get-HTMLContentOpen -HeaderText "Office 365 Shared Mailboxes"
        $rpt += get-htmlcontentdatatable $SharedMailboxTable -HideFooter
        $rpt += Get-HTMLContentClose
    $rpt += get-htmltabcontentclose
	
        $rpt += get-htmltabcontentopen -TabName $tabarray[6] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy)) 
        $rpt += Get-HTMLContentOpen -HeaderText "Office 365 Contacts"
            $rpt += get-htmlcontentdatatable $ContactTable -HideFooter
        $rpt += Get-HTMLContentClose
        $rpt += Get-HTMLContentOpen -HeaderText "Office 365 Mail Users"
            $rpt += get-htmlcontentdatatable $ContactMailUserTable -HideFooter
        $rpt += Get-HTMLContentClose
    $rpt += get-htmltabcontentclose
	
    $rpt += get-htmltabcontentopen -TabName $tabarray[7] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy)) 
        $rpt += Get-HTMLContentOpen -HeaderText "Office 365 Room Mailboxes"
            $rpt += get-htmlcontentdatatable $RoomTable -HideFooter
        $rpt += Get-HTMLContentClose
        $rpt += Get-HTMLContentOpen -HeaderText "Office 365 Equipment Mailboxes"
            $rpt += get-htmlcontentdatatable $EquipTable -HideFooter
        $rpt += Get-HTMLContentClose
    $rpt += get-htmltabcontentclose
	

$rpt += Get-HTMLClosePage

$Day = (Get-Date).Day
$Month = (Get-Date).Month
$Year = (Get-Date).Year
$ReportName = ( "$Month" + "-" + "$Day" + "-" + "$Year" + "-" + "O365 Tenant Report")
Save-HTMLReport -ReportContent $rpt -ShowReport -ReportName $ReportName -ReportPath $ReportSavePath
