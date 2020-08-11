#NMC XML Generator
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$app_name = 'NMC XML Generator'
#Will Prompt for Mednax Login
$name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Mednax Login ID", $app_name , "ex.JDoe")
$test_name = Get-ADUser -Identity $name
## Validate Mednax Login 

if ($test_name) {
    # NMC Required Feilds
    # name gathered above
    $first_name = Get-ADUser -Identity $name -Properties GivenName | Select-Object -ExpandProperty GivenName
    $last_name = Get-ADUser -Identity $name -Properties sn | Select-Object -ExpandProperty sn
    $login_id = $name + "_mednax"
    $token = $name
    
    #Prompts for Ticket Number 
    $full_name = Get-ADUser -Identity $name -Properties DisplayName | Select-Object -ExpandProperty DisplayName 
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $reqid = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Request Number", $app_name , "REQ, RITM, SCTASK")

    #Prompt to Validate who will be recieving a license
    [System.Windows.MessageBox]::Show('Would  you like to license ' + $full_name + ' for Dragon Medical One ?', $app_name , 'OKCancel')


    $xml_template = @'
<?xml version="1.0" encoding="utf-8"?>
<Users xmlns="http://nuance.com/NAS/UserImport">
    <User FirstName="$($First_Name)" LastName="$($Last_Name)" LoginId="$($login_id)" Password="">
        <TokenCredential>"$($name)"</TokenCredential>
        <GroupMembership>Default Site\Dragon Medical Users</GroupMembership>
        <GroupMembership>Default Site\PowerMic Mobile Users</GroupMembership>
    </User>
</Users>
'@

    # Save Dialog Box and 
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.forms")
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.FileName = $reqid # Populates with Requst Number in filename
    $dlg.DefaultExt = ".xml" # Default file extension
    $dlg.Filter = "XML Document (.xml)|*.xml" # Filter files by extension

    # Show save file dialog box
    $result = $dlg.ShowDialog()

    # Process save file dialog box results
    if ($result = 'OK') {
        # Save document
        $filename = $dlg.FileName
        #$template = Add-Content -path $filename -Value $xml_template
        $expanded = Invoke-Expression "@`"`r`n$xml_template`r`n`"@"
        # New-Item -Path $filename -ItemType File -Force
        Add-Content -path $filename -Value $expanded
        [System.Windows.MessageBox]::Show('XML for ' + $full_name + ' has been created. You will now be directed to https://nmc-hc-prod-us.nuancehdp.com/ to finish licensing.', $app_name, 'OK', 'Error')
        Start-Sleep -s 2
        Start-Process "https://nmc-hc-prod-us.nuancehdp.com/"
    }
} 
else {
    [System.Windows.MessageBox]::Show($name + ' does not exist, please start again', $app_name, 'OK', 'Error')
}