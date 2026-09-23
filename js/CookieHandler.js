function setCookie(cname, cvalue, exdays) {
    var d = new Date();
    d.setHours(23);d.setMinutes(59);d.setSeconds(59);
    var expires = "expires="+d.toUTCString();
    document.cookie = cname + "=" + cvalue + "; " + expires;
}

function getCookie(cname) {
    var name = cname + "=";
    var ca = document.cookie.split(';');
    for(var i=0; i<ca.length; i++) {
        var c = ca[i];
        while (c.charAt(0)==' ') c = c.substring(1);
        if (c.indexOf(name) == 0)
		{
			return c.substring(name.length, c.length);
		}
    }
    return "";
}

                 