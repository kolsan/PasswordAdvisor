$MailSubject_ENU_ES = 'Password-Contraseña // News Capsule'
$MailBody_ENU_ES1 =   '<img src="cid:IMAGE.jpg" /><p>&nbsp;</p>
<table style="height: 106px; margin-left: auto; margin-right: auto;" border="4" width="95%" cellspacing="3" cellpadding="3">
<tbody>

<tr>
<td style="width: 76px;">&nbsp;
<p>Hi '
$MailBody_ENU_ES2 = '!</p>
<p>We inform you that your password <strong>will expire in the 3 next days</strong>. So, please proceed to change it. Keep in mind that, if you don&acute;t change it in time, you can experience some services malfunctions including an account lock.</p>
<p>If you are in the office, you can change the password any time simply pressing <strong>ctrl + alt + supr</strong> buttons. If you are outside, you can change it too, via the <strong><a href="https://mail.domain.com" target="_blank" rel="noopener">webmail</a></strong> web and there, in options (up and right), choose change password.</p>
<p>Once the password has been changed, remember to do it as well in your mobile devices, wifi, mail etc.</p>
<p>Security information concern to all of us. By this reason, the IT department want to involve you with this compromised and share with you a news capsule.</p>
</td>
</tr>

<tr>
<td style="width: 76px;">
<p>Hola '

$MailBody_ENU_ES3 = '!</p>
<p>Te recordamos que tu contrase&ntilde;a <strong>va a caducar en 3 dias</strong>, asi que , por favor, procede a cambiarla. Te recordamos que si no la cambias a tiempo puedes experimentar bloqueos de cuenta o malfuncionamiento en los servicios.</p>
<p>Si estas en la oficina, puedes cambiar la contrase&ntilde;a en cualquier momento pulsando la combinaci&oacute;n de teclas <strong>ctrl + alt + supr</strong>. Si te encuentras fuera, tambien puedes. Entra en el <a href="https://mail.domain.com" target="_blank" rel="noopener"><strong>correo</strong></a> web y en opciones (arriba a la derecha), selecciona <em>cambiar contrase&ntilde;a</em>.</p>
<p>Una vez cambiada, recuerda hacerlo tambien en tus dispositivos moviles, correo electronico, wifi etc.</p>
<p>La seguridad de la informacion nos concierne a todos. Es por ello que desde IKS queremos hacerte participe en este compromiso y compartir contigo esta pildora informativa.</p>
</td>
</tr>
</tbody>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<blockquote>

<p>This is an automatic mail, so do not respond to it. If further information is required, contact with helpdesk using the application o calling to 1110.</p>
  <p>Este correo se ha generado automaticamente, asi que por favor no respondas. Si tienes cualquier duda contacta con helpdesk.</p>
</blockquote>'

#------------------------------------------------------------------------------------

$MailSender = "Sender  <sender@domain.com>"
$SMTPServer = 'server.domain.com'

$hoy = get-date;
$Capsule = Get-Random -Maximum 9 -Minimum 1

$users = Get-ADUser -Server "ldapserver.domain.com:3268" -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0 } `
 -Properties "Name","GivenName", "EmailAddress", "UserPrincipalName", "msDS-UserPasswordExpiryTimeComputed","PasswordNeverExpires" ,PasswordLastSet| Select-Object -Property  @{Name = "PassExpirationDate"; Expression = {($_."PasswordLastSet").AddDays(30) }}, PasswordLastSet,"Name","GivenName", "UserPrincipalName", "EmailAddress","PasswordNeverExpires"

$Capsule
$Quote = [char]39

 foreach ($user in $users) {
 $d1 = $user.PassExpirationDate
 $ts = New-TimeSpan -Start $hoy -End $d1
  


 if ( $ts.days -eq 3 -and $user.EmailAddress.Length -ne 0) 
 {
      $user.Name, $user.PassExpirationDate,$ts.days,$user.PasswordNeverExpires,$user."UserPrincipalName",$user.EmailAddress,$user.GivenName
      if ($user."UserPrincipalName" -like '*@subdomain1.domain.com')
        {
        $Subject = $MailSubject_ENU_ES
        $EmailBody = $MailBody_ENU_ES1 + $user.GivenName + $MailBody_ENU_ES2 + $user.GivenName + $MailBody_ENU_ES3
        $File0 = "PATH\IMAGE.jpg"
        $File1 = “PATHPATPATH\ES" +"\File_ES_00"+ $Capsule +".pdf”
        $File2 = “PATPATH\ENU" +"\File_ENU_00"+ $Capsule +".pdf”
        $File3 = “PATH\CH" +"\File_CH_00"+ $Capsule +".pdf”
        Send-MailMessage -To $user.EmailAddress -From $MailSender -SmtpServer $SMTPServer -Subject $Subject -BodyAsHtml $EmailBody -Encoding "UTF32" -Attachments $File0,$File1,$File2,$File3 '

         }

     if ($user."UserPrincipalName" -like '*@subdomain2.domain.com')
        {
        $Subject = $MailSubject_ENU_ES
        $EmailBody = $MailBody_ENU_ES1 + $user.GivenName + $MailBody_ENU_ES2 + $user.GivenName + $MailBody_ENU_ES3
        $File0 = "PATH\IMAGE.jpg"
        $File1 = “PATH\ES" +"\File_ES_00"+ $Capsule +".pdf”
        $File2 = “PATH\ENU" +"\File_ENU_00"+ $Capsule +".pdf”
        $File3 = “PATH\CZ" +"\File_CZ_00"+ $Capsule +".pdf”
        Send-MailMessage -To $user.EmailAddress -From $MailSender -SmtpServer $SMTPServer -Subject $Subject -BodyAsHtml $EmailBody -Encoding "UTF32" -Attachments $File0,$File1,$File2,$File3 '
        }

     if ($user."UserPrincipalName" -like '*@subdomain3.domain.com')
        {
        $Subject = $MailSubject_ENU_ES
        $EmailBody = $MailBody_ENU_ES1 + $user.GivenName + $MailBody_ENU_ES2 + $user.GivenName + $MailBody_ENU_ES3
        $File0 = "PATH\IMAGE.jpg"
        $File1 = “PATH\ES" +"\File_ES_00"+ $Capsule +".pdf”
        $File2 = "PATH\ENU" +"\File_ENU_00"+ $Capsule +".pdf”
        Send-MailMessage -To $user.EmailAddress -From $MailSender -SmtpServer $SMTPServer -Subject $Subject -BodyAsHtml $EmailBody -Encoding "UTF32" -Attachments $File0,$File1,$File2 '
        
        
        }

 


   

    }

    
 
 }
