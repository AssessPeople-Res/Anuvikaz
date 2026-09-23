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

	if(Frm.txtFrientName_1.value != "")
	{
		if(Frm.txtFrientEmailID_1.value == "")
		{
			alert("Enter your Friend e-Mail ID");
			Frm.txtFrientEmailID_1.focus();
			return false;
		}
	}

	if(Frm.txtFrientName_2.value != "")
	{
		if(Frm.txtFrientEmailID_2.value == "")
		{
			alert("Enter your Friend e-Mail ID");
			Frm.txtFrientEmailID_2.focus();
			return false;
		}
	}

	if(Frm.txtFrientName_3.value != "")
	{
		if(Frm.txtFrientEmailID_3.value == "")
		{
			alert("Enter your Friend e-Mail ID");
			Frm.txtFrientEmailID_3.focus();
			return false;
		}
	}

	if(Frm.txtFrientName_4.value != "")
	{
		if(Frm.txtFrientEmailID_4.value == "")
		{
			alert("Enter your Friend e-Mail ID");
			Frm.txtFrientEmailID_4.focus();
			return false;
		}
	}

	if(Frm.txtFrientName_5.value != "")
	{
		if(Frm.txtFrientEmailID_5.value == "")
		{
			alert("Enter your Friend e-Mail ID");
			Frm.txtFrientEmailID_5.focus();
			return false;
		}
	}

	if(Frm.txtFrientEmailID_1.value != "")
	{
		if(Frm.txtFrientName_1.value == "")
		{
			alert("Please Enter your Friend Name")
			Frm.txtFrientName_1.focus()
			return false;
		}

	   var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;  
	   if(emailPattern.test(Frm.txtFrientEmailID_1.value) == false)
	   {
			Frm.txtFrientEmailID_1.focus(); 
			alert("Email ID is not valid")
			return false;
	   }
	}

	if(Frm.txtFrientEmailID_2.value != "")
	{
		if(Frm.txtFrientName_2.value == "")
		{
			alert("Please Enter your Friend Name")
			Frm.txtFrientName_2.focus()
			return false;
		}

	   var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;  
	   if(emailPattern.test(Frm.txtFrientEmailID_2.value) == false)
	   {
			Frm.txtFrientEmailID_2.focus(); 
			alert("Email ID is not valid")
			return false;
	   }
	}

	if(Frm.txtFrientEmailID_3.value != "")
	{
		if(Frm.txtFrientName_3.value == "")
		{
			alert("Please Enter your Friend Name")
			Frm.txtFrientName_3.focus()
			return false;
		}

	   var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;  
	   if(emailPattern.test(Frm.txtFrientEmailID_3.value) == false)
	   {
			Frm.txtFrientEmailID_3.focus(); 
			alert("Email ID is not valid")
			return false;
	   }
	}

	if(Frm.txtFrientEmailID_4.value != "")
	{
		if(Frm.txtFrientName_4.value == "")
		{
			alert("Please Enter your Friend Name")
			Frm.txtFrientName_4.focus()
			return false;
		}

	   var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;  
	   if(emailPattern.test(Frm.txtFrientEmailID_4.value) == false)
	   {
			Frm.txtFrientEmailID_4.focus(); 
			alert("Email ID is not valid")
			return false;
	   }
	}

	if(Frm.txtFrientEmailID_5.value != "")
	{
		if(Frm.txtFrientName_5.value == "")
		{
			alert("Please Enter your Friend Name")
			Frm.txtFrientName_5.focus()
			return false;
		}

	   var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;  
	   if(emailPattern.test(Frm.txtFrientEmailID_5.value) == false)
	   {
			Frm.txtFrientEmailID_5.focus(); 
			alert("Email ID is not valid")
			return false;
	   }
	}

	if(Frm.txtFrientEmailID_1.value == "" && Frm.txtFrientEmailID_2.value == "" && Frm.txtFrientEmailID_3.value == "" && Frm.txtFrientEmailID_4.value == "" && Frm.txtFrientEmailID_5.value == "")
	{
		alert("Please enter the Email IDs")
		Frm.txtFrientEmailID_1.focus(); 
		return false;
	}

	if(Frm.txtFullName.value == "")
	{
		alert("Please enter Your Name")
		Frm.txtFullName.focus(); 
		return false;
	}

	if(Frm.txtEmailSubject.value == "")
	{
		alert("Please enter Email Subject")
		Frm.txtEmailSubject.focus(); 
		return false;
	}

	if(Frm.txtEmailContent.value == "")
	{
		alert("Please enter Email Content")
		Frm.txtEmailContent.focus(); 
		return false;
	}
}

</script>
</head>
<body>
  <form method="post" action="ImageVerification.asp" id="form1" onSubmit="return check()">
  
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
