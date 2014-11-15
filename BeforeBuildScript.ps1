$psw = (Get-ItemProperty -path HKCU:\SOFTWARE\Tulpep).PfxPassword
if(!$psw)
{
throw "HKCU\Software\Tulpep\PfxPassword does not exist"
}
$psw = ConvertTo-SecureString $psw -AsPlainText -Force
Import-PfxCertificate -FilePath .\Tulpep-Certificate.pfx -CertStoreLocation cert:\CurrentUser\My -Password $psw
