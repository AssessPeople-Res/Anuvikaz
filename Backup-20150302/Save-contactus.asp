<%

set con = Server.CreateObject("ADODB.Connection")
con.open ("provider=sqloledb;data source=10.20.10.1;uid=ciadmin;pwd=ciadmin123;DATABASE=ci_newsletter")

FullName	= Request.Form("txtFullName")
EmailID		= Request.Form("txtEmailID")
ContactNo	= Request.Form("txtContactNo")
Subject		= Request.Form("txtSubject")
Comments	= Request.Form("txtComments")

FullName	= Replace(FullName,"'","''") 
ContactNo	= Replace(ContactNo,"'","''") 
Subject		= Replace(Subject,"'","''") 
Comments	= Replace(Comments,"'","''") 

If NOT EmailID = "" then 

	Con.Execute("insert into TbAnuvikaz_Feedback(FullName,EmailID,ContactNo,Subject,Comments) values('"&FullName&"','"&EmailID&"','"&ContactNo&"','"&Subject&"','"&Comments&"')") 

	MailContent = "" 
	MailContent = MailContent& "<table width='70%' border='1' align='center' cellpadding='1' cellspacing='1' bordercolor='#CCCCCC'>" 
	MailContent = MailContent& "<tr><td><font face='Verdana' size='2'>One of the visitor contacted through Anuvikaz website </font></td></tr>" 
	MailContent = MailContent& "</table><br>" 
	MailContent = MailContent& "<table width='70%' border='1' align='center' cellpadding='1' cellspacing='1' bordercolor='#CCCCCC'>" 
	MailContent = MailContent& "<tr><td width='30%'><font face='Verdana' size='2'>Name</font></td><td><font face='Verdana' size='2'>"&FullName&"</font></td></tr>" 
	MailContent = MailContent& "<tr><td><font face='Verdana' size='2'>Email ID </font></td><td><font face='Verdana' size='2'>"&EmailID&"</font></td></tr>" 
	MailContent = MailContent& "<tr><td><font face='Verdana' size='2'>Contact No.</font></td><td><font face='Verdana' size='2'>"&ContactNo&"</font></td></tr>" 
	MailContent = MailContent& "<tr><td><font face='Verdana' size='2'>Subject</font></td><td><font face='Verdana' size='2'>"&Subject&"</font></td></tr>" 
	MailContent = MailContent& "<tr><td><font face='Verdana' size='2'>Comments</font></td><td><font face='Verdana' size='2'>"&Comments&"</font></td></tr>" 
	MailContent = MailContent& "</table>" 

	Set Mail 	= Server.CreateObject("CDONTS.NewMail") 
	Mail.From 	= "visitor@anuvikaz.org.in" 
	Mail.To		= "ayyalu@assesspeople.com"
	Mail.Subject  	=  "Comments from Anuvikaz Website" 
	Mail.Body 	= MailContent
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
	color: #990000;
	font-weight: bold;
	font-size: 14px;
}
-->
</style>
</head>

<body>
<div class="container">
	<div class="header">
    	<h1><a href="index.html"><img src="images/logo.jpg" alt="Anuvikaz" title="Anuvikaz" /></a></h1>
        <p>
        	<span>
            	<a href="https://www.facebook.com/Anuvikaz" target="_blank"><img src="images/facebook.png" alt="" border="0" /></a>
                <a href="#"><img src="images/twitter.png" alt="" /></a>            </span>        </p>
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
            <li><a href="contactus.html" class="current">Contact Us</a></li>
        </ul>
        <div class="clr"></div>
    </div>
    
    <div class="inner-banner">
    	<div class="don fl">
            <h3>Bring Smiles</h3>
            <ul>
                <li><a href="donate.html">Donate</a></li>
                <li><a href="volunteer.html">Volunteer</a></li>
                <li><a href="Spread-the-Word.html">Spread the Word</a></li>
        	</ul>
        </div>
        <div class="add fr">
        	<img src="images/add-banner.jpg" alt=""  height="98" />
        </div>
        <div class="clr"></div>
        <div class="contantArea">
        	<h2>Contact Us</h2>
            <p><br />
              <span class="style1">              Thank you for contacting us. We will get in touch with you at the earliest.</span> <br />
            </p>
          <p>
            	<label>&nbsp;</label>
            </p>
        </div>
        <div class="clr"></div>
    </div>
    <div class="footer">
    	<p>Copyright &copy; anuvikaz.org</p>
    </div>
</div>
</body>
</html>
