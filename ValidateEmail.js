	function ValidateEmail( target, cmt ) {
		if( target.value.match(/\S/) != null && target.value.match(/[\w\-\~\.\_]+\@[\w\-\~]+(\.[\w\-\~]+)+/g) != target.value || target.value == ' ' 
			alert( cmt + " is not valid format.");
			target.focus();
			return true;
		}
		return false;
	}