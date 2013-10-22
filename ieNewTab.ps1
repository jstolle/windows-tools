$win = New-Object -comObject Shell.Application

If ($args.length -lt 1)
{
    Write-Error "You must pass at least one URL to the script."
}
Else
{
    ForEach ($arg in $args)
    {
        If ($arg -like 'http*')
        {
            $inurl = $arg
        }
        Else
        {
            $inurl = "http://$arg"
        }

        $ie = @($win.windows() | ? { $_.Name -eq 'Windows Internet Explorer' -and $_.LocationName -NotMatch 'Do Not Close' })

        If ($ie.length -eq 0)
        {
            $newie = New-Object -comObject InternetExplorer.Application
            $newie.Navigate2($inurl)
            $newie.visible = $true
        }
        Else
        {
            $ie[0].Navigate2($inurl, 0x10000)
        }
    }
}
