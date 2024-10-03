@echo off 
cd\ 
%systemdrive% 
:Setup_SMTP_Server // Section for setting the name of the exchange 
server to be used and type of authentication to be used. 1 means to use 
Basic authentication, 2 means to use NTLM authentication, 0 is for Anonymous authentication 
certutil -setreg exit\smtp\SMTPServer "<ExchangeServerNameOrIP>" 
certutil -setreg exit\smtp\SMTPAuthenticate 1 
:Setup_CA_For_Exit_Module // Section for turning events on or off. In 
this case, on. 
certutil -setsmtpinfo -p "<Account>" Administrator 
certutil -setreg exit\smtp\eventfilter +EXITEVENT_CRLISSUED 
certutil -setreg exit\smtp\eventfilter +EXITEVENT_CERTDENIED 
certutil -setreg exit\smtp\eventfilter +EXITEVENT_CERTISSUED 
certutil -setreg exit\smtp\eventfilter +EXITEVENT_CERTPENDING 
certutil -setreg exit\smtp\eventfilter +EXITEVENT_CERTREVOKED 
certutil -setreg exit\smtp\eventfilter +EXITEVENT_SHUTDOWN 
certutil -setreg exit\smtp\eventfilter +EXITEVENT_STARTUP 
:CrlIssued // Section for setting CRLIssued parameters. 
certutil -setreg exit\smtp\CRLissued\To "<EmailAddress>" 
certutil -setreg exit\smtp\CRLissued\From "<EmailAddress>" 
certutil -setreg exit\smtp\CRLissued\CC "<EmailAddress>" 
certutil -setreg exit\smtp\CRLissued\bodyformat "A new CRL has been 
issued" 
certutil -setreg exit\smtp\CRLissued\titleformat "A new CRL was issued by %%1" 
certutil -setreg exit\smtp\CRLissued\BodyArg "" 
certutil -setreg exit\smtp\CRLissued\TitleArg +"SanitizedCAName" 
:Denied // Section for setting Denied parameters 
certutil -setreg exit\smtp\Denied\From "<EmailAddress>" 
certutil -setreg exit\smtp\Denied\CC "<EmailAddress>" 
certutil -setreg exit\smtp\Denied\titleformat "Your certificate request 
was denied by %%1" 
certutil -setreg exit\smtp\Denied\BodyArg "" 
Certutil -setreg exit\smtp\Denied\BodyFormat "" 
call Stop_Start_CA 
certutil -setreg exit\smtp\Denied\BodyArg +"Request.RequestID" 
certutil -setreg exit\smtp\Denied\BodyArg +"Request.RequesterName" 
certutil -setreg exit\smtp\Denied\BodyArg +"Request.SubmittedWhen" 
certutil -setreg exit\smtp\Denied\BodyArg +"Request.DistinguishedName" 
certutil -setreg exit\smtp\Denied\BodyArg +"Request.DispositionMessage" 
certutil -setreg exit\smtp\Denied\BodyArg +"Request.StatusCode" 
Certutil -setreg exit\smtp\Denied\BodyFormat +"Your Request ID is: %%1" 
Certutil -setreg exit\smtp\Denied\BodyFormat +"The Requester Name is: %%2" 
Certutil -setreg exit\smtp\Denied\BodyFormat +"The Request Submission Date was: %%3" 
Certutil -setreg exit\smtp\Denied\BodyFormat +"Subject Name: %%4" 
Certutil -setreg exit\smtp\Denied\BodyFormat +"Request Disposition Message: %%5" 
Certutil -setreg exit\smtp\Denied\BodyFormat +"Request StatusCode: %%6" 
certutil -setreg exit\smtp\Denied\TitleArg +"SanitizedCAName" 
:Certificate_Issued // Section for setting Issued parameters. 
certutil -setreg exit\smtp\Issued\From "<EmailAddress>" 
certutil -setreg exit\smtp\Issued\CC "<EmailAddress>" 
certutil -setreg exit\smtp\Issued\titleformat "Your certificate has been issued by %%1" 
certutil -setreg exit\smtp\Issued\BodyArg +"RawCertificate" 
Certutil -setreg exit\smtp\Issued\BodyFormat "" 
net stop certsvc 
call Stop_Start_CA 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Request ID: %%1" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"UPN: %%2" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Requester Name: %%3" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Serial Number: %%4" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Valid not before: %%5" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Valid not after: %%6" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Distinguished Name: %%7" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Certificate Template: %%8" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Certificate Hash: %%9" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Request Disposition Message: %%10" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Copy and paste the 
following in Notepad, save and install" 
Certutil -setreg exit\smtp\Issued\BodyFormat +"Binary Certificate: %%11" 
:Certificate_Pending // Section for setting Pending parameters. 
certutil -setreg exit\smtp\Pending\From "<EmailAddress>" 
certutil -setreg exit\smtp\Pending\CC "<EmailAddress>" 
certutil -setreg exit\smtp\Pending\titleformat "Your certificate is pending on %%1" 
Certutil -setreg exit\smtp\Pending\BodyFormat "" 
call Stop_Start_CA 
Certutil -setreg exit\smtp\Pending\BodyFormat +"Request ID: %%1" 
Certutil -setreg exit\smtp\Pending\BodyFormat +"UPN: %%2" 
Certutil -setreg exit\smtp\Pending\BodyFormat +"Requester Name: %%3" 
Certutil -setreg exit\smtp\Pending\BodyFormat +"Time submitted: %%4" 
Certutil -setreg exit\smtp\Pending\BodyFormat +"Distinguished Name: %%5" 
Certutil -setreg exit\smtp\Pending\BodyFormat +"Certificate Template used: %%6" 
Certutil -setreg exit\smtp\Pending\BodyFormat +"Request Disposition Message: %%7" 
:Certificate_Revoked // Section for setting Revoked parameters. 
certutil -setreg exit\smtp\Revoked\From "<EmailAddress>" 
certutil -setreg exit\smtp\Revoked\CC "<EmailAddress>" 
certutil -setreg exit\smtp\Revoked\titleformat "Your certificate was revoked by %%1" 
Certutil -setreg exit\smtp\Revoked\BodyFormat "" 
call Stop_Start_CA 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Request ID: %%1" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Revoked when: %%2" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Effective: %%3" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Reason for being revoked: %%4" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"UPN: %%5" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Requester Name: %%6" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Serial Number: %%7" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Was not valid until: %%8" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Was not valid after: %%9" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Distinguished Name: %%10" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Certificate Template: %%11" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Certificate Hash: %%12" 
Certutil -setreg exit\smtp\Revoked\BodyFormat +"Request Status: %%13" 
:Certificate_Authority_Shutdown // Section for setting Shutdown parameters. 
certutil -setreg exit\smtp\Shutdown\To "<EmailAddress>" 
certutil -setreg exit\smtp\Shutdown\From "<EmailAddress>" 
certutil -setreg exit\smtp\Shutdown\CC "<EmailAddress>" 
:Certificate_Authority_Startup // Section for setting Startup parameters. 
certutil -setreg exit\smtp\Startup\To "<EmailAddress>" 
certutil -setreg exit\smtp\Startup\From "<EmailAddress>" 
certutil -setreg exit\smtp\Startup\CC "<EmailAddress>" 
:Stop_Start_CA // This is just a sub-routine for stopping and starting the CA. 
net stop certsvc & net start certsvc 
:Exit 
echo Certificate Services SMTP Exit module has now been configured. 
echo . 
pause 
exit