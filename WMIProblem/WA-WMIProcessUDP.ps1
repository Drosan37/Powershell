# Workaround
$a = [WMI]'root\policy:__Win32Provider.name="PolicSOM"'
$a.HostingModel = "NetworkServiceHost:PolicSOM"
$a.put()


# Revert
#$a = [WMI]'root\policy:__Win32Provider.name="PolicSOM"'
#$a.HostingModel = "NetworkServiceHost"
#$a.put() 