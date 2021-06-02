# Printers Logon Script
#
# Uses a Group Policy Object's User Preference items to add printers.
# Create the GPO but do not link it, it should be used by this script but NOT processed by group policy at logon.
#
# Katy Nicholson 2020-08-04, updated 2021-02-02
# https://katystech.blog

# Set your domain and the GUID of the policy object here.
$domain = "fqdn.of.your.domain"
$policyGUID = "{C96580A0-00DC-4E63-899A-3D82CF5EE141}"




$GPOxmlFile = "\\$domain\sysvol\$domain\Policies\$policyGUID\User\Preferences\Printers\Printers.xml"

#Set up the form to show progress to the user
Add-Type -AssemblyName System.Windows.Forms
#Force the form to show
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
try{
    Add-Type -Name win -Member $t -Namespace native
    [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)
} catch {
    $Null
}
$global:PrinterForm = New-Object System.Windows.Forms.Form
$global:Font = New-Object System.Drawing.Font("Calibri", 12)
$global:LabelTop = 25

$PrinterForm.ClientSize = '500,300'
$PrinterForm.text = "Printers"
$PrinterForm.BackColor = "#ffffff"

$Label1 = New-Object System.Windows.Forms.Label
$Label1.Text = "Please wait while your printer connections are configured..."
$Label1.Font = $Font
$Label1.Width = 500
$Label1.Height = 30
$Label1.Top = 10
$Label1.Left = 10

$PrinterForm.Controls.Add($Label1)
[void]$PrinterForm.Refresh()
[void]$PrinterForm.Show()
[void]$PrinterForm.Focus()

#Supporting functions

#Removes domain when in the format DOMAIN\username
Function stripDomain($item) {
    if ($item.IndexOf('\') -gt 0) { 
        return $item.Substring(($item.IndexOf('\')+1))
    } else {
        return $item
    }
}

#Searches AD to check if the specified username is in the specified group
Function isInGroup([String]$groupname, [String]$username) {
    #Strip domain from group and user names, if present
    $groupname = stripDomain($groupname)
    $username = stripDomain($username)
    
    if ($group = ([ADSISearcher]"sAMAccountName=$groupname").FindOne()) {
        $groupDN = $group.Properties['distinguishedName']
        if ($isMember = ([ADSISearcher]"(&(|(objectClass=computer)(objectClass=person))(sAMAccountName=$username)(memberOf:1.2.840.113556.1.4.1941:=$groupDN))").FindOne()) {
            return $true
        }
    }
    return $false
}

#Evaluate the item targeting filters
Function evaluateFilters($Filters) {

    $FilterResults = @()
    foreach ($Filter in $Filters.ChildNodes) { 
        #Process filters on this item
        [Bool]$return = $false
        switch ($Filter.LocalName) {
            "FilterCollection" {
                #Recurse through collections
                $return = evaluateFilters $Filter
            }
            "FilterGroup" {
                #Filter to whether the user (or computer) is a member of the specified group
                if ($Filter.userContext -eq "1") {
                    $username = $env:USERNAME
                } else {
                    $username = "$env:COMPUTERNAME$"
                }
                if (isInGroup $Filter.name $username) {
                    $return = $true
                }
                break
            }
            "FilterComputer" {
                #Filter if the computer is the specified name
                if ($env:COMPUTERNAME -ieq (stripDomain $Filter.name)) {
                    $return = $true
                }
                break
            }
            "FilterUser" {
                #Filter if the user is the specified name
                if ($env:USERNAME -ieq (stripDomain $Filter.name)) {
                    $return = $true
                }
                break
            }
            "FilterOrgUnit" {
                #Filter if the user or computer is a member of the specified OU (either directly or in a child OU)
                if ($Filter.userContext -eq "1") {
                    $itemDN = ([ADSISearcher]"sAMAccountName=$env:USERNAME").FindOne().Properties.distinguishedname[0]
                } else {
                    $itemDN = ([ADSISearcher]"sAMAccountName=$env:COMPUTERNAME$").FindOne().Properties.distinguishedname[0]
                }
                if ($itemDN.IndexOf($Filter.name, [System.StringComparison]::OrdinalIgnoreCase) -gt -1) {
                    #Item is in the OU
                    $OUPath = $itemDN.SubString($itemDN.IndexOf("OU="))
                    if ($Filter.directmember -eq "1") {
                        $return = ($OUPath -ieq $Filter.name)
                    } else {
                        $return = $true
                    }
                }
                break
            }
        }
        # If filter is set to NOT then negate return value
        if ($Filter.not -eq "1") {
            $FilterResults += ,(@($Filter.bool.ToString(),(!$return)))
        } else {
            $FilterResults += ,(@($Filter.bool.ToString(),($return)))
        }
    }
    
    #Evaluate the filters = If only one, return that result, otherwise AND or OR the first two results, depending on what is configured, then AND/OR this with the next value, repeat until done.
    if ($FilterResults.Count -gt 1) {
        $firstItem = $true
        foreach ($item in $FilterResults) {
            if ($firstItem) {
                $previousItem = $item[1]
                $firstItem = $false
            } else {
                if ($item[0] -eq "AND") {
                    if ($previousItem -and $item[1]) {
                        $previousItem = $true
                    } else {
                        $previousItem = $false
                    }
                } else {
                    if ($previousItem -or $item[1]) {
                        $previousItem = $true
                    } else {
                        $previousItem = $false
                    }
                }
            }
        }
        return $previousItem
    } else {
        return $FilterResults[0][1]
    }
    
}

#Update the status labels on the form
Function updateStatus($printerName, $status) {
    
    foreach ($Label in $PrinterForm.Controls) {
        if ($Label.Tag -and $printerName -eq $Label.Tag.Printer.ToString()) {
            $Label.Text = $Label.Tag.Display.ToString() + " - " + $status
            return
        }
    }
    $PrinterForm.Refresh()
}

#Add an item status label to the form
Function createStatus($printerName, $displayName, $status) {
    $LabelTop += 25
    $NewLabel = New-Object System.Windows.Forms.Label
    $NewLabel.Text = $displayName + " - " + $status
    $NewLabel.Tag = [PSCustomObject]@{
        "Printer"=$printerName
        "Display"=$displayName
    }
    $NewLabel.Font = $Font
    $NewLabel.Width = 500
    $NewLabel.Height = 20
    $NewLabel.Top = $LabelTop
    $NewLabel.Left = 20
    $PrinterForm.Controls.Add($NewLabel)
    $PrinterForm.Refresh()
    Set-Variable -Name "LabelTop" -Value $LabelTop -Scope Global
}


[xml]$Printers = Get-Content -Path $GPOxmlFile
#Logic for adding printers, to run in background job
$AddPrinterScriptBlock = {
   
    Function AddPrinter([String]$printer, $default) {

        # When Windows 10 decides it doesn't want to add on the first go at random intervals - retry forever until it adds
        try {
            Add-Printer -ConnectionName $printer -ErrorAction Stop
        } catch {
            AddPrinter $printer $default
        }
        #Set default printer if specified
        if ($default -eq "1") { 
            (New-Object -ComObject WScript.Network).SetDefaultPrinter($printer)
        }
    }
    AddPrinter $args[0] $args[1]
}

#Logic for removing printers, to run in background job
$RemovePrinterScriptBlock = {

    Function RemovePrinter([String]$printer) {
        if ($printer -eq "All") {
            #Remove all printer connections
            Get-Printer | Where {$_.Type -eq "Connection"} | Remove-Printer
        } else {
            #Remove specified printer connections
            (New-Object -ComObject WScript.Network).RemovePrinterConnection($printer)
        }
    }
    RemovePrinter $args[0]
}

#Loop through each printer connection in the GPP list, if it has item level targeting (filters) then evaluate these. If matched, perform the relevant action.
foreach ($Printer in $Printers.Printers.SharedPrinter) {
    [Bool]$matched = $true
    
    if ($Printer.Filters.HasChildNodes) {
        $matched = evaluateFilters $Printer.Filters
    }
    if ($matched -eq $true) {
        if ($Printer.Properties.action -eq "D") {
            if ($Printer.Properties.deleteAll -eq "1") {
                #Delete All network printers - we Wait-Job on this one as we don't want it running at the same time as we are creating printer connections.
                #Delete All needs to be priority 1 in the GPP list.
                createStatus "DeleteAll" "Clean up old printers" "Running..."
                Start-Job $RemovePrinterScriptBlock -Name "DeleteAll" -ArgumentList "All" | Out-Null
                Wait-Job -Name "DeleteAll"
                updateStatus "DeleteAll" "Complete"
            } else {
                createStatus $Printer.name ("Remove " + $printer.name.ToString()) "Not started"
                Start-Job $RemovePrinterScriptBlock -Name $printer.name.ToString() -ArgumentList $Printer.Properties.path | Out-Null
            }
        } else {
            createStatus $Printer.name ("Add " + $Printer.name.ToString()) "Not started"
            Start-Job $AddPrinterScriptBlock -Name $Printer.name.ToString() -ArgumentList @($Printer.Properties.path, $Printer.Properties.default) | Out-Null
        }
    }
}

#Check on the background jobs, updating the UI.
While ($jobs = Get-Job -State "Running")
{
    ForEach ($job in $jobs) {
        
        switch ($job.State) {
            "Running" {
                updateStatus $job.Name "Running..."
                break
            }
            "Completed" {
                updateStatus $job.Name "Complete"
                break
            }
            "Failed" {
                updateStatus $job.Name "Failed"
                break
            }
        }
    }
    Start-Sleep -Milliseconds 100
    [System.Windows.Forms.Application]::DoEvents()

}
Start-Sleep 1
$PrinterForm.Hide()
#Clean up
Get-Job | Remove-Job
