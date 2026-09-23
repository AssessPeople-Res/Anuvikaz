<%

set con = Server.CreateObject("ADODB.Connection")
con.open ("provider=sqloledb;data source=10.20.10.1;uid=ciadmin;pwd=ciadmin123;DATABASE=ci_newsletter")

FrientName_1		= Request.Form("txtFrientName_1")
FrientName_2		= Request.Form("txtFrientName_2")
FrientName_3		= Request.Form("txtFrientName_3")
FrientName_4		= Request.Form("txtFrientName_4")
FrientName_5		= Request.Form("txtFrientName_5")

FrientName_1	= Replace(FrientName_1,"'","''") 
FrientName_2	= Replace(FrientName_2,"'","''") 
FrientName_3	= Replace(FrientName_3,"'","''") 
FrientName_4	= Replace(FrientName_4,"'","''") 
FrientName_5	= Replace(FrientName_5,"'","''") 

FrientEmailID_1	= Request.Form("txtFrientEmailID_1")
FrientEmailID_2	= Request.Form("txtFrientEmailID_2")
FrientEmailID_3	= Request.Form("txtFrientEmailID_3")
FrientEmailID_4	= Request.Form("txtFrientEmailID_4")
FrientEmailID_5	= Request.Form("txtFrientEmailID_5")

FrientEmailID_1	= Replace(FrientEmailID_1,"'","''") 
FrientEmailID_2	= Replace(FrientEmailID_2,"'","''") 
FrientEmailID_3	= Replace(FrientEmailID_3,"'","''") 
FrientEmailID_4	= Replace(FrientEmailID_4,"'","''") 
FrientEmailID_5	= Replace(FrientEmailID_5,"'","''") 

FullName			= Request.Form("txtFullName")
EmailSubject		= Request.Form("txtEmailSubject")
EmailContent		= Request.Form("txtEmailContent") 

FullName		= Replace(FullName,"'","''") 
EmailSubject	= Replace(EmailSubject,"'","''") 
EmailContent	= Replace(EmailContent,"'","''") 

Sender = FullName&"<noreply@anuvikaz.org>"

EmailContent = FullName&" has refered you to visit <a href=www.anuvikaz.org>www.anuvikaz.org</a> <br><br>"&EmailContent

If NOT FullName = "" then 
	Con.Execute("insert into TbAnuvikaz_Friend_Reference(GuestName,EmailSubject,EmailContent,RowDate) values('"&FullName&"','"&EmailSubject&"','"&EmailContent&"',getDate())")

	set LastRowIDRs = Server.CreateObject("ADODB.Recordset") 
	LastRowIDRs.Open "SELECT IDENT_CURRENT('TbAnuvikaz_Friend_Reference') LastRowID",con  
	LastRowID = LastRowIDRs("LastRowID")
	LastRowIDRs.Close 
End If 


If NOT FrientEmailID_1 = "" then 
	Con.Execute("insert into TbAnuvikaz_Friend_Reference_Detail(RowID,FriendName,EmailID) values('"&LastRowID&"','"&FrientName_1&"','"&FrientEmailID_1&"')")	

	Set Mail 	= Server.CreateObject("CDONTS.NewMail") 
	Mail.From 	= Sender 
	Mail.To		= FrientEmailID_1
	Mail.Subject  	=  EmailSubject 
	Mail.Body 	= EmailContent
	Mail.bodyformat	= 0
	Mail.mailformat = 0 
	Mail.Send 
End If 

If NOT FrientEmailID_2 = "" then 
	Con.Execute("insert into TbAnuvikaz_Friend_Reference_Detail(RowID,FriendName,EmailID) values('"&LastRowID&"','"&FrientName_2&"','"&FrientEmailID_2&"')")	

	Set Mail 	= Server.CreateObject("CDONTS.NewMail") 
	Mail.From 	= Sender 
	Mail.To		= FrientEmailID_2
	Mail.Subject  	=  EmailSubject 
	Mail.Body 	= EmailContent
	Mail.bodyformat	= 0
	Mail.mailformat = 0 
	Mail.Send 
End If 

If NOT FrientEmailID_3 = "" then 
	Con.Execute("insert into TbAnuvikaz_Friend_Reference_Detail(RowID,FriendName,EmailID) values('"&LastRowID&"','"&FrientName_3&"','"&FrientEmailID_3&"')")	

	Set Mail 	= Server.CreateObject("CDONTS.NewMail") 
	Mail.From 	= Sender 
	Mail.To		= FrientEmailID_3
	Mail.Subject  	=  EmailSubject 
	Mail.Body 	= EmailContent
	Mail.bodyformat	= 0
	Mail.mailformat = 0 
	Mail.Send 
End If 

If NOT FrientEmailID_4 = "" then 
	Con.Execute("insert into TbAnuvikaz_Friend_Reference_Detail(RowID,FriendName,EmailID) values('"&LastRowID&"','"&FrientName_4&"','"&FrientEmailID_4&"')")	

	Set Mail 	= Server.CreateObject("CDONTS.NewMail") 
	Mail.From 	= Sender 
	Mail.To		= FrientEmailID_4 
	Mail.Subject  	=  EmailSubject 
	Mail.Body 	= EmailContent 
	Mail.bodyformat	= 0
	Mail.mailformat = 0 
	Mail.Send 
End If 

If NOT FrientEmailID_5 = "" then 
	Con.Execute("insert into TbAnuvikaz_Friend_Reference_Detail(RowID,FriendName,EmailID) values('"&LastRowID&"','"&FrientName_5&"','"&FrientEmailID_5&"')")	

	Set Mail 	= Server.CreateObject("CDONTS.NewMail") 
	Mail.From 	= Sender 
	Mail.To		= FrientEmailID_5 
	Mail.Subject  	=  EmailSubject 
	Mail.Body 	= EmailContent 
	Mail.bodyformat	= 0
	Mail.mailformat = 0 
	Mail.Send 
End If 

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Anuvikaz</title>
<link rel="stylesheet" href="css/style.css" media="screen" type="text/css" />
<style type="text/css">
<!--
.style1 {
	font-size: 14px;
	font-weight: bold;
}
.style2 {color: #990000}
-->
</style>
</head>

<body>
<div class="container">
	<div class="header">
    	<h1><a href="index.html"><img src="images/logo.jpg" alt="Anuvikaz" title="Anuvikaz" /></a></h1>
        <p>
        	<span>
            	<a href="#"><img src="images/facebook.png" alt="" /></a>
                <a href="#"><img src="images/twitter.png" alt="" /></a>
            </span>
        </p>
		<ul>
        	<li><a href="index.html">Home</a></li>
            <li><a href="#">Projects</a>
            	<ul>
                    <li><a href="health-care.html">Health Care</a></li>
                    <li><a href="scholarship-education.html">Scholarship &amp; Education</a></li>
                    <li><a href="employability-programmes.html">Employability Programmes</a></li>
                    <li><a href="vocational-guidance-programmes.html">Vocational Guidance</a></li>
                    <li><a href="mother-earth-programmes.html">Mother Earth Programmes</a></li>
				</ul>
            </li>
            <li><a href="csr.html">CSR</a></li>
            <li><a href="#">Gallery</a>
            	<ul>
                	<li><a href="photo-gallery.html">Photo Gallery</a></li>
                    <li><a href="video-gallery.html">Video Gallery</a></li>
                </ul>
            </li>
            <li><a href="contactus.html">Contact Us</a></li>
        </ul>
        <div class="clr"></div>
    </div>
    
    <div class="inner-banner">
    	<div class="don fl">
            <h3>Bring Smiles</h3>
            <ul>
                <li><a href="donate.html">Donate</a></li>
                <li><a href="volunteer.html">Volunteer</a></li>
                <li><a href="#">Spread the Word</a></li>
        	</ul>
        </div>
        <div class="add fr">
        	<img src="images/add-banner.jpg" alt=""  height="98" />
        </div>
        <div class="clr"></div>
        <div class="contantArea">
        	<h2>Spread the Good Work</h2>
            <p><br />
              <br />
                  <span class="style1"><span class="style2">Thank you for refering Anuvikaz to your friends. <br />
                  <br />
                  We look forward to your continued support.<br />
                  <br />
                  <br />
                  </span><br />
            </span></p>
      </div>
        <div class="clr"></div>
    </div>
    <div class="footer">
    	<p>Copyright &copy; anuvikaz.org</p>
    </div>
</div>
</body>
</html>
