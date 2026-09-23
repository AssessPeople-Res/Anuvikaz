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

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Anuvikaz</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<link href="css/res-menu.css" rel="stylesheet" type="text/css" />

<link href="css/media.css" rel="stylesheet" type="text/css" />
<link href="css/slider.css" rel="stylesheet" type="text/css" />

</head>

<body>
<div class="main_wrapper">
	<div class="header">
    	<h1><a href="index.html"><img src="images/logo.png" alt="" /></a></h1>
  		<p>
        	<a href="https://www.facebook.com/Anuvikaz" target="_blank"><img src="images/facebook.jpg" alt="" /></a>
            <a href="#"><img src="images/twitter.jpg" alt="" /></a>
        </p>
        <div class="clr"></div>
    </div>
</div>
<div class="menu_outer">
	<div class="menu_inner">
    	<div class="main_wrapper">
        	<div class="animenu">
            	<button class="animenu__toggle">
                    <span class="animenu__toggle__bar"></span>
                    <span class="animenu__toggle__bar"></span>
                    <span class="animenu__toggle__bar"></span>
                </button>
            	<ul class="animenu__nav">
                	<li><a href="index.html" class="active">Home</a></li>
                    <li><a href="#">Projects</a>
                    	<ul class="animenu__nav__child">
                            <li><a href="english-express.html">English Express</a></li>
                            <li><a href="vocational-guide.html">Vocational Guidance Programs</a></li>
                            <li><a href="employability.html">Employability Programs</a></li>
                            <li><a href="mother-earth.html">Mother Earth Programs</a></li>
                            <li><a href="reading-room.html">Reading Rooms</a></li>
                 <!--            <li><a href="age-care-centre.html">Age Care Centres</a></li> -->
                        </ul>
                    </li>
                    <li><a href="#">CSR</a>
					<ul class="animenu__nav__child">
					 <li><a href="ketti.html">Ketti,Nilgiris - Project</a></li> 
                     <li><a href="Vagapannai.html">Vagapanai,Nilgiris - Project</a></li>
					 </ul>
					 </li>	
                    <li><a href="donate.html">Donate</a></li>
                    <li><a href="volunteer.html">Volunteer</a></li>
                    <li><a href="spread-word.html" class="active">Spread the Word</a></li>
                    <li><a href="#">Gallery</a>
                    	<ul class="animenu__nav__child">
                            <li><a href="video-gallery.html">Video Gallery</a></li>
                            <li><a href="photo-gallery.html">Photo Gallery</a></li>
                        </ul>
                    </li>
                    <li><a href="contactus.html">Contact Us</a></li>
                </ul>
                <div class="clr"></div>
            </div>
        </div>
    </div>
</div>
<div class="main_wrapper">
	<div class="banner">
    	<div class="slider-container" id="caption-slide"> 
			<div class="slider">
            	<div><img src="images/slide03.jpg" alt="">
                	<span class="caption">Employability Programs</span>
                </div>
                <div><img src="images/slide.jpg" alt=""></div>
            	<div><img src="images/slide01.jpg" alt="">
                	<span class="caption">English Express</span>
                </div>
                <div><img src="images/slide.jpg" alt=""></div>                
                <div><img src="images/slide04.jpg" alt="">
                	<span class="caption">Mother Earth Programs</span>
                </div>
                <div><img src="images/slide.jpg" alt=""></div>
                <div><img src="images/slide07.jpg" alt=""></div>
                <div><img src="images/slide.jpg" alt=""></div>
                <div><img src="images/slide02.jpg" alt="">
                	<span class="caption">Career Guidance Programs</span>
                </div>
                <div><img src="images/slide.jpg" alt=""></div>
                <div><img src="images/slide06.jpg" alt="">
                	<span class="caption">Age Care Centres</span>
                </div>
                <div><img src="images/slide.jpg" alt=""></div>
                <div><img src="images/slide05.jpg" alt="">
                	<span class="caption">Reading Rooms</span>
                </div>
                <div><img src="images/slide.jpg" alt=""></div>
                <div><img src="images/slide08.jpg" alt="">
                	<span class="caption">Vision Care Supports</span>
                </div>
           	</div>
            <!--<div class="switch" id="prev"><span></span></div>
			<div class="switch" id="next"><span></span></div> -->
		 </div>
         <div  class="clr"></div>
    </div>
    <div class="about">
    	<div class="about_left">
        	<h3>Spread the Word</h3>
            <h6>Thank you for refering Anuvikaz to your friends. <br> We look forward to your continued support.</h6>
        </div>
        
        
        <div class="contant_right">
        	<h3>Photo Gallery</h3>
            <ul>
            	<li><img src="images/photo_img01.jpg" alt="" /></li>
                <li><img src="images/photo_img02.jpg" alt="" /></li>
                <li><img src="images/photo_img03.jpg" alt="" /></li>
                <li>
                	<p>
                    	<a href="photo-gallery.html">More photos</a>
                    </p>
                </li>
                <div class="clr"></div>
            </ul>
        </div>
        <div class="clr"></div>
    </div>
</div>
<div class="footer_outer">
	<div class="footer_inner">
    	<div class="footerbg">
        	<div class="main_wrapper">
            	<div class="footer">
                	<p>Copyright &copy; 2015 - Anuvikaz - All Rights Reserved - Privacy Policy</p>
                </div>
            </div>
        </div>
    </div>
</div>


<script type="text/javascript" src="js/jquery-1.11.1.min.js"></script>
<script type="text/javascript" src="js/slider.js"></script>
<script type="text/javascript">
	$("#slider-container").sliderUi({
		speed: 700,
		cssEasing: "cubic-bezier(0.285, 1.015, 0.165, 1.000)"
	});
	$("#caption-slide").sliderUi({
		caption: true
	});
</script>
<script type="text/javascript" src="js/animenu.js"></script>
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-42593673-1']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
</body>
</html>
