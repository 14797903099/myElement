var ajax = function(info){
	if(window.XMLHttpRequest){
		var xhr = new XMLHttpRequest();	
	}else{
		var xhr = new ActiveXObject("Microsoft.XMLHTTP");
	}
	var url = info.url;
	var type = info.type||"GET";
	var data = info.data||"";
	
	if(type=="GET"){
		xhr.open("GET",url+"?"+data+"&t="+new Date().getTime(),true);
		xhr.send();
	}else{
		xhr.open("POST",url,true);
		xhr.setRequestHeader("myHeader","myvalue");
		xhr.send(data);
	}
	
	xhr.onload = function(){
		if(info.success){
			info.success(xhr.responseText);
		}
	}
}

export var ajax = ajax;