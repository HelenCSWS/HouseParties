<!-- #include file="adovbs.inc"-->
<SCRIPT LANGUAGE=VBScript RUNAT=Server>


'height
public const g_topFramHT=58 'top frame height
public const g_mdFramHT=560 'middle frame height
public const g_btFramHT=0 'bottome frame height

''quote
public const g_mixquotes=630
public const g_msg_mixquotes= "The quote can not be longer than 530 characters."
public const g_quote_counts = 660

''monthly wine description
public const g_mixDesc=330
public const g_msg_mixDesc= "The description of this wine can not be longer than 360 characters."
'You can add special event handlers in this file that will get run automatically when
'special Active Server Pages events occur. To create these handlers, just create a
'subroutine with a name from the list below that corresponds to the event you want to
'use. For example, to create an event handler for Session_OnStart, you would put the
'following code into this file (without the comments):

'Sub Session_OnStart
'**Put your code here **
'End Sub

'EventName              Description
'Session_OnStart        Runs the first time a user runs any page in your application
'Session_OnEnd          Runs when a user's session times out or quits your application
'Application_OnStart    Runs once when the first page of your application is run for the first time by any user
'Application_OnEnd      Runs once when the web server shuts down

public const g_ImagePathWine = "/images/wineimgs"
public const g_ImagePathSp = "/images/SpLogo"
public const g_ImagePathSpPhoto = "/images/SpPhotos"
public const g_webname = "Tiptop"
public const g_web_email = "webmaster@christopherstewart.com"

'===================== DSN CONNECTION
'DSN=date;DB=datin;SERVER=user1-dat;UID=root;PORT=;FLAG=0;"
'public const g_strCon = "DSN=CSWS;DB=CSWS;server=jack;UID=root;PASSWORD=;PORT=;FLAG=0"nissan

'=========DSN LESS ================================
'server
'public const g_strCon = "Driver={MySQL ODBC 3.51 Driver};server=www.christopherstewart.com;UID=nissan;PASSWORD=gw!2a7v;DATABASE=christopherstewart_db"    '--- on server
'local
public const g_strCon = "Driver={MySQL ODBC 3.51 Driver};server=localhost;port=3306;Option=16387;UID=root;PASSWORD=;DATABASE=christopherstewart_db"    ' on local macthine

public function GetConnection() 
	dim strCon
	dim strSQL
	
	set GetConnection=Server.CreateObject("ADODB.connection")		
	GetConnection.open g_strCon
end function

public function SpToNb(strText)

	SpToNb = replace(strText," ","&nbsp;")
end function

public function NbToSp(strText)

	NbToSp = replace(strText,"&nbsp;"," ")
end function

public function strsqltoDB(strsql)
	
	strsqltoDB = replace(strsql,"\","\\")
	strsqltoDB = replace(strsqltoDB,"'","\'")
'	strsqltoDB = replace(strsqltoDB,"%","\%")
'	strsqltoDB = replace(strsqltoDB,"_","\_")
	
end function

public sub delWinePics(rsWine)
'rsWine has country_name and tbs_wine.*
	if not rsWine.EOF then
		dim countrypath
		dim imagepath
		dim fso
	
		set fso = server.CreateObject("Scripting.FileSystemObject")

		countrypath = Server.MapPath(g_ImagePathWine)
		countrypath = countrypath & "\" & rsWine("country_name")
		if fso.FolderExists(countrypath) then
			do while not rsWine.EOF
				imagepath = countrypath & "\" & rsWine("wine_id")
				if fso.FolderExists(imagepath) then
					fso.DeleteFolder imagepath, true
				end if
				rsWine.MoveNext
			loop
		end if
		set fso = nothing
	end if
end sub

'add by wenling
public sub delSpPic(pic_name,ral_id)
'rsWine has country_name and tbs_wine.*
		dim imagepath
		dim fso
	
		set fso = server.CreateObject("Scripting.FileSystemObject")

		imagepath = Server.MapPath(g_ImagePathSpPhoto)
		imagepath =imagepath& "\" & ral_id & "\" & pic_name

		if fso.FileExists(imagepath) then
			fso.DeleteFile imagepath, true
		end if
		set fso = nothing
end sub


public sub delSPLogo(rsSP)
'rsSP has country_name and tbs_suppliers.*
	if not rsSP.EOF then
		dim countrypath
		dim imagepath
		dim fso
	
		set fso = server.CreateObject("Scripting.FileSystemObject")

		countrypath = Server.MapPath(g_ImagePathSp)
		countrypath = countrypath & "\" & rsSP("country_name")
		if fso.FolderExists(countrypath) then
			do while not rsSP.EOF
				imagepath = countrypath & "\" & rsSP("supplier_id")
				if fso.FolderExists(imagepath) then
					fso.DeleteFolder imagepath, true
				end if
				rsSP.MoveNext
			loop
		end if
		set fso = nothing
	end if
end sub

public function getProvince(proID) 
	select case Cstr(proId)
		case "1"
			getProvince = "British Columbia"
		case "2"
			getProvince = "Alberta"
		case "9"
			getProvince = "Saskatchewan"
		case "3"
			getProvince = "Manitoba"
		case "7"
			getProvince = "Ontario"
		case "8"
			getProvince = "Quebec"
		case "4"
			getProvince = "Newfoundland"
		case "5"
			getProvince = "Northwest Territories"
		case "6"
			getProvince = "Nunavut"
		case "10"
			getProvince = "Yukon"
	end select
	
end function

public function getMonth(mID) 
	select case Cstr(mID)
		case "1"
			getMonth = "January"
		case "2"
			getMonth = "February"
		case "3"
			getMonth = "March"
		case "4"
			getMonth = "April"
		case "5"
			getMonth = "May"
		case "6"
			getMonth = "June"
		case "7"
			getMonth = "July"
		case "8"
			getMonth = "August"
		case "9"
			getMonth = "September"
		case "10"
			getMonth = "October"
		case "11"
			getMonth = "November"
		case "12"
			getMonth = "December"
	end select
	
end function

public function getSupplierName(spid)
		dim objRec1
		dim strSQL
		
		getSupplierName =""
		strSQL="Select supplier_name from tbs_suppliers where supplier_id = " & spid
		set objRec1= server.CreateObject("ADODB.Recordset")
		objRec1.Open strSQL,g_strCon,adOpenKeyset
		
		
		if not objRec1.EOF then
			getSupplierName = replaceName4APS(objRec1("supplier_name"))
		end if

end function

public function replaceName4APS(name)
	
	replaceName4APS = replace(name,"'","&#39;" )

end function

public function replaceName2APS(name)
	
	replaceName4APS = replace(name,"&#39;","'" )

end function

public function replaceName4SP(name,sReplace)
	
	replaceName4APS = replace(name," ","sReplace" )

end function

public function c(name,sReplace)
	
	replaceName4APS = replace(name,"sReplace;","'" )

end function


Public function getNews(strURL)
	' Start out declaring our variables.
	' You are using Option Explicit aren't you?
	Dim objWinHttp
	

	Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
	

	objWinHttp.Open "GET", strURL
	objWinHttp.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 4.01; Windows 95)"

	on error resume next
	objWinHttp.Send

	getNews = objWinHttp.ResponseText

	set objWinHttp=nothing
end function

Public function getWebID(web_id)
	
		dim nIndex
		dim nCnts
		dim sHref

		getWebID = ""
		Select case web_id	
			case 1
				getwebID = "Wine business online"
			case 2
				getwebID = "The advertiser"
			case 3 
				getwebID = "Wine news only "
			case 4
				getwebID = "Topix.net"
		end select

end function

public function setempname2z(sval)

	setempname2z = "z~"
	if trim(sval)<>"" then
		setempname2z = sval
	end if
end function

public function setz2empname(sval)

	setz2empname = sval
	if trim(sval)="z~" then
		setz2empname = ""
	end if
end function

public function getCountryLeftPic(country_id)

	getCountryLeftPic="no graphic"
	if Cstr(country_id) = "1" then
			getCountryRightPic = "images/francebg1.jpg"
	elseif Cstr(country_id) = "2" then
			getCountryRightPic = "images/spain_rightbanner1.gif"
	elseif Cstr(country_id) = "3" then
			getCountryRightPic = "images/canadabg1.jpg"
	elseif Cstr(country_id) = "4" then
			getCountryRightPic = "images/italybg1.jpg"
	elseif Cstr(country_id) = "5" then
			getCountryRightPic = "images/moroccobg1.jpg"
	elseif Cstr(country_id) = "6" then
			getCountryRightPic = "images/algeriabg1.jpg"
	elseif Cstr(country_id) = "11" then
			getCountryRightPic = "images/australiabg1.jpg"
	elseif Cstr(country_id) = "12" then
			getCountryRightPic = "images/usabg2.gif"
	elseif Cstr(country_id) = "13" then
			getCountryRightPic = "images/argentinabg2.gif"
	elseif Cstr(country_id) = "14" then
			getCountryRightPic = "images/chilebg2.gif"
	elseif Cstr(country_id) = "15" then
			getCountryRightPic = "images/hungarybg2.gif"
	elseif Cstr(country_id) = "16" then
			getCountryRightPic = "images/indiabg2.gif"
	elseif Cstr(country_id) = "17" then
			getCountryRightPic = "images/uruguaybg2.gif"
	end if
end function

public function getCountryRightPic(country_id)

	getCountryRightPic="no graphic"
	if Cstr(country_id) = "1" then
		getCountryRightPic = "images/france_rightbanner3.gif"
		'strctrygrd = "images/fr2.jpg"
	elseif Cstr(country_id) = "2" then
		getCountryRightPic = "images/spain_rightbanner3.gif"
	elseif Cstr(country_id) = "3" then
		getCountryRightPic = "images/canada_rightbanner3.gif"
	elseif Cstr(country_id) = "4" then 'italy
		getCountryRightPic = "images/italy_rightbanner3.gif"
	elseif Cstr(country_id) = "5" then 
		getCountryRightPic = "images/morocco_rightbanner3.gif"
	elseif Cstr(country_id) = "6" then 'italy
		getCountryRightPic = "images/algeria_rightbanner3.gif"
	elseif Cstr(country_id) = "11" then 'italy
		getCountryRightPic = "images/australia_rightbanner3.gif"
	elseif Cstr(country_id) = "12" then 'USA
		getCountryRightPic = "images/usabg.gif"
	elseif Cstr(country_id) = "13" then 'Argentina
		getCountryRightPic = "images/argentina_rightbanner.gif"
	elseif Cstr(country_id) = "14" then 'Chile
		getCountryRightPic = "images/chile_rightbanner.gif"
	elseif Cstr(country_id) = "15" then 'Hungary
		getCountryRightPic = "images/hungary_rightbanner.gif"
	elseif Cstr(country_id) = "16" then 'India
		getCountryRightPic = "images/india_rightbanner.gif"
	elseif Cstr(country_id) = "17" then 'Uruguay
		getCountryRightPic = "images/uruguay_rightbanner.gif"
	end if
end function

public function getCountryMonthlyPic(strcnid)
	getCountryMonthlyPic = "no grapic"
	if Cstr(strcnid) = "1" then
			getCountryMonthlyPic = "images/french2.gif"
		elseif Cstr(strcnid) = "2" then'spain
			getCountryMonthlyPic = "images/Spain_WineOfTheMonth1.gif"
		elseif Cstr(strcnid) = "3" then'ca
			getCountryMonthlyPic = "images/Canada_wineofthemonth1.gif"
		elseif Cstr(strcnid) = "4" then 'italy
			getCountryMonthlyPic = "images/Italy_wineofthemonth1.gif"
		elseif Cstr(strcnid) = "5" then 'morocco
			getCountryMonthlyPic = "images/Morocco_wineofthemonth1.gif"
		elseif Cstr(strcnid) = "6" then 'Algeria
			getCountryMonthlyPic = "images/Algeria_wineofthemonth1.gif"
		elseif Cstr(strcnid) = "11" then 'Aus
			getCountryMonthlyPic = "images/Australia_wineofthemonth1.gif"
		elseif Cstr(strcnid) = "12" then 'USA
			getCountryMonthlyPic = "images/USA_wineofthemonth1.gif"
		elseif Cstr(strcnid) = "13" then 'Argentian
			getCountryMonthlyPic = "images/argentina.gif"
		elseif Cstr(strcnid) = "14" then 'chile
			getCountryMonthlyPic = "images/chile.gif"
		elseif Cstr(strcnid) = "15" then 'hungary
			getCountryMonthlyPic = "images/hungary.gif"
		elseif Cstr(strcnid) = "16" then 'india
			getCountryMonthlyPic = "images/india.gif"
		elseif Cstr(strcnid) = "17" then 'Uruguay
			getCountryMonthlyPic = "images/uruguay.gif"
	end if

end function


function RTESafe(strText)
	'returns safe code for preloading in the RTE
	dim tmpString
	
	tmpString = trim(strText)
	
	'convert all types of single quotes
	tmpString = replace(tmpString, chr(145), chr(39))
	tmpString = replace(tmpString, chr(146), chr(39))
	tmpString = replace(tmpString, "'", "&#39;")
	
	'convert all types of double quotes
	tmpString = replace(tmpString, chr(147), chr(34))
	tmpString = replace(tmpString, chr(148), chr(34))
'	tmpString = replace(tmpString, """", "\""")
	
	'replace carriage returns & line feeds
	tmpString = replace(tmpString, chr(10), " ")
	tmpString = replace(tmpString, chr(13), " ")
	
	RTESafe = tmpString
end function

</SCRIPT>
