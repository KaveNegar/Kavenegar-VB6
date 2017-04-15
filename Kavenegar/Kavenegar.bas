Attribute VB_Name = "kavenegar"
Public apikey As String
Public Function toUnix(dt) As Long
    toUnix = DateDiff("s", "1/1/1970", dt)
End Function

Function request(base As String, method As String, parameters As String) As Object
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "https://api.kavenegar.com/v1/" & apikey & "/" & base & "/" & method & ".json"
    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlhttp.setRequestHeader "Content-Length", Len(parameters)
    xmlhttp.send parameters
    Dim result As String
    result = xmlhttp.responseText
    'Set xmlhttp = Nothing
    'request = JSON.parse(result)
    Dim p As Object
    Set p = json.parse(result)
    Set request = p
End Function
Function sms_send(sender As String, message As String, receptor() As String) As Object
    Dim parameters As String
    parameters = "sender=" & sender & "&" _
                 & "message=" & message & "&" _
                 & "receptor=" & Join(receptor, ",")
                                 
    Set sms_send = request("sms", "send", parameters)
End Function

Function sms_sendarray(sender() As String, message() As String, receptor() As String) As Object
    Dim parameters As String
    parameters = "sender=[" & Join(sender, ",") & "]&" _
                 & "message=[" & Join(message, ",") & "]&" _
                 & "receptor=[" & Join(receptor, ",") & "]&"
                                  
    sms_sendarray = request("sms", "sendarray", parameters)
End Function
Function verify_lookup(receptor As String, token As String, token2 As String, token3 As String, template As String) As Object
    Dim parameters As String
    parameters = "token=" & token & "&" _
                   & "token2=" & token2 & "&" _
                  & "token3=" & token3 & "&" _
                 & "template=" & template & "&" _
                 & "receptor=" & receptor
      
    verify_lookup = request("verify", "lookup", parameters)
End Function
Function sms_status(messageid As String) As Object
    Dim parameters As String
    parameters = "messageid=" & messageid & "&" _
                 
     sms_status = request("sms", "status", parameters)

End Function
Function sms_statuslocalmessageid(localid As String) As Object
    Dim parameters As String
    parameters = "localid=" & localid & "&" _
                 
    sms_statuslocalmessageid = request("sms", "statuslocalmessageid", parameters)
End Function
Function sms_select(messageid() As String) As Object
    Dim parameters As String
    parameters = "messageid=" & Join(messageid, ",") & "&"
                                
    sms_select = request("sms", "select", parameters)
End Function
Function sms_selectoutbox(startdate As String, enddate As String) As Object
    Dim parameters As String
    parameters = "startdate=" & startdate & "&" _
                    & "enddate=" & enddate & "&" _
              
    sms_selectoutbox = request("sms", "selectoutbox", parameters)
End Function
Function sms_latestoutbox(localid As String) As Object
    Dim parameters As String
    parameters = "pagesize=" & PageSize & "&" _
                            
    sms_latestoutbox = request("sms", "latestoutbox", parameters)
End Function
Function sms_countoutbox(status As String, startdate As String, enddate As String) As Object
    Dim parameters As String
    parameters = "status=" & status & "&" _
                     & "startdate=" & startdate & "&" _
                     & "enddate=" & enddate & "&" _
              
    sms_countoutbox = request("sms", "countoutbox", parameters)
End Function
Function sms_cancel(messageid() As String) As Object
    Dim parameters As String
    parameters = "messageid=" & Join(messageid, ",") & "&"
                            
    sms_cancel = request("sms", "cancel", parameters)
End Function
Function sms_receive(linenumber As String, isread As String) As Object
    Dim parameters As String
    parameters = "linenumber=" & linenumber & "&" _
                    & "isread=" & isread & "&" _
                
   sms_receive = request("sms", "receive", parameters)
End Function
Function sms_countinbox(linenumber As String, isread As String, startdate As String, enddate As String) As Object
    Dim parameters As String
    parameters = "linenumber=" & linenumber & "&" _
                     & "isread=" & isread & "&" _
                     & "startdate=" & startdate & "&" _
                     & "enddate=" & enddate & "&" _
                
   sms_countinbox = request("sms", "countinbox", parameters)
End Function
Function sms_countpostalcode(postalcode As String) As Object
    Dim parameters As String
    parameters = "postalcode=" & postalcode & "&" _
                            
    sms_countpostalcode = request("sms", "countpostalcode", parameters)
End Function

Function sms_sendbypostalcode(postalcode As String, sender As String, message As String, mcistartindex As String, mcicount As String, mtnstartindex As String, mtncount As String) As Object
    Dim parameters As String
    parameters = "postalcode=" & postalcode & "&" _
                     & "sender=" & sender & "&" _
                     & "message=" & message & "&" _
                     & "mcistartindex=" & mcistartindex & "&" _
                      & "mcicount=" & mcicount & "&" _
                      & "mtnstartindex=" & mtnstartindex & "&" _
                      & "mtncount=" & mtncount & "&" _
                 
     sms_sendbypostalcode = request("sms", "sendbypostalcode", parameters)
End Function
Function account_info(parameters As String) As Object
    s = request("account", "info", parameters)
End Function
Function account_config(apilogs As String) As Object
    Dim parameters As String
    parameters = "apilogs=" & apilogs & "&" _
        
    account_config = request("account", "config", parameters)
End Function
