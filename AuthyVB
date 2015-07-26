'## Two-factor authentication with Authy
dim authyURL, authyID, authyKEY, authyParms, authyToken, o
authyURL = "http://api.authy.com/protected/json/verify/"
authyID = "_______" '## user ID to validate
authyKEY = "________" '## your API key
authyToken = "_______" '## token supplied by user

authyParms = "api_key=" & authyKEY	

authyURL = authyURL & authyToken & "/" & authyID & "?" & authyParms
'verify token format: https://api.authy.com/protected/{FORMAT}/verify/{TOKEN}/{AUTHY_ID}?api_key={KEY}	
	
'## verify the token
Set o = CreateObject("Msxml2.ServerXMLHTTP")
on error resume next
o.open "GET", authyURL, False 
o.send

'## check o.responseText & o.Status
if InStr(o.responseText,"is valid") AND o.Status = "200" then
  '## token was valid
  '## good code here
else
  '## token was invalid
  '## bad code here
end if
