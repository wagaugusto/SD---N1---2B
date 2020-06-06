$("#btn-save").click( function() {
    var text1 = $("#textarea1").val();
    var text2 = $("#textarea2").val();
    var text = "";
    var numOM = text1.replace(/\D/g,'');
    var listaOM = [];
    for(var i=0; i<numOM.length; i = i +10){
        text = text + "If Not IsObject(application) Then" + "\r\n";
        text = text + "   Set SapGuiAuto  = GetObject(\"SAPGUI\")" + "\r\n";
        text = text + "   Set application = SapGuiAuto.GetScriptingEngine" + "\r\n";
        text = text + "End If" + "\r\n";
        text = text + "If Not IsObject(connection) Then" + "\r\n";
        text = text + "   Set connection = application.Children(0)" + "\r\n";
        text = text + "End If" + "\r\n";
        text = text + "If Not IsObject(session) Then" + "\r\n";
        text = text + "   Set session    = connection.Children(0)" + "\r\n";
        text = text + "End If" + "\r\n";
        text = text + "If IsObject(WScript) Then" + "\r\n";
        text = text + "   WScript.ConnectObject session,     \"on\"" + "\r\n";
        text = text + "   WScript.ConnectObject application, \"on\"" + "\r\n";
        text = text + "End If" + "\r\n";
        text = text + "session.findById(\"wnd[0]\").maximize" + "\r\n";
        text = text + "session.findById(\"wnd[0]/tbar[0]/okcd\").text = \"iw32\"" + "\r\n";
        text = text + "session.findById(\"wnd[0]\").sendVKey 0" + "\r\n";
        text = text + "session.findById(\"wnd[0]/usr/ctxtCAUFVD-AUFNR\").text = \""+numOM.substring(i,i+10)+"\"" + "\r\n";
        text = text + "session.findById(\"wnd[0]\").sendVKey 0" + "\r\n";
        text = text + "session.findById(\"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell\").text = \"*"+text2+"*\" + vbCr + \"\"" + "\r\n";
        text = text + "session.findById(\"wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/subSUB_TEXT:SAPLCOIH:1103/cntlLTEXT/shell\").setSelectionIndexes 0,0" + "\r\n";
        text = text + "session.findById(\"wnd[0]/tbar[1]/btn[36]\").press" + "\r\n";
        text = text + "session.findById(\"wnd[1]/tbar[0]/btn[0]\").press" + "\r\n";
        text = text + "session.findById(\"wnd[0]/tbar[0]/btn[3]\").press" + "\r\n";                        
        //text = text + numOM.substring(i,i+10) +"\r\n";
    }    
    var filename = $("#input-fileName").val();    
    filename = "script.vbs";
    var blob = new Blob([text], {type: "" });
    saveAs(blob, filename, "text/javascript");     

});