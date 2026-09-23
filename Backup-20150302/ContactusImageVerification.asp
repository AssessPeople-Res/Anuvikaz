<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
  <title>BotDetect CAPTCHA Basic ASP Sample</title>
  <link type="text/css" rel="Stylesheet" href="StyleSheet.css" />
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  <!-- #include file ="BotDetect.asp" -->
  
<script language="javascript">

function check()
{
	var Frm = parent.document.forms[0]; 

	if(Frm.txtFullName.value == "")
	{
		Frm.txtFullName.focus(); 
		alert("Enter your Name")
		return false;
	}

	if(Frm.txtEmailID.value == "")
	{
		Frm.txtEmailID.focus(); 
		alert("Enter your Email ID")
		return false;
	}
	else
	{
	   var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;  
	   if(emailPattern.test(Frm.txtEmailID.value) == false)
	   {
			Frm.txtEmailID.focus(); 
			alert("Email ID is not valid")
			return false;
	   }
	}

	if(Frm.txtContactNo.value == "")
	{
		Frm.txtContactNo.focus(); 
		alert("Enter your Contact No.")
		return false;
	}

	if(Frm.txtSubject.value == "")
	{
		Frm.txtSubject.focus(); 
		alert("Enter subject")
		return false;
	}

	if(Frm.txtComments.value == "")
	{
		Frm.txtComments.focus(); 
		alert("Enter your Comments / Queries")
		return false;
	}
}

</script>
</head>
<body>
  <form method="post" action="ContactusImageVerification.asp" id="form1" onSubmit="return check()">
  
    <fieldset>
        
        <label for="CaptchaCode">Retype the characters from the picture:</label>
        
        <% ' Adding BotDetect CAPTCHA to the page 
				  Dim SampleCaptcha : Set SampleCaptcha = (New Captcha)("SampleCaptcha")
          SampleCaptcha.UserInputID = "CaptchaCode"
          Response.Write SampleCaptcha.Html %>
        
        <div class="validationDiv">
            <input name="CaptchaCode" type="text" id="CaptchaCode" />
            <input type="submit" name="ValidateCaptchaButton" value="Submit" id="ValidateCaptchaButton" />
						<br/>
						<% ' CAPTCHA user input validation (only if the form was sumbitted)
							If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
						    Dim isHuman : isHuman = SampleCaptcha.Validate()
						    If Not isHuman Then 
								  ' CAPTCHA validation failed, show error message
								  Response.Write "<span class=""incorrect"">Incorrect code</span>"
						    Else 
								  ' CAPTCHA validation passed, perform protected action
								  Response.Write " <script>	  parent.document.forms[0].submit();  	</script>"
						    End If 
						  End If
						%>
        </div>
    </fieldset>
   
  </form>
</body>
</html>
