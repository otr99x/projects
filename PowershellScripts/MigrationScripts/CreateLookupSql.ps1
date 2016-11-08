function AddAdditionalFields($newitem)
{
    try
    {
        $urlValue = ($newitem.URL).ToUpper()
        if([string]::IsNullOrEmpty($urlValue))
        {
            $newitem | Add-Member -MemberType NoteProperty -Name "LivelinkID" -Value ""
        }
        else
        {
            if(($index = $urlValue.LastIndexOf('&OBJID=')) -ge 0)
            {
                $value = $urlValue.Substring($index + '&OBJID='.Length, 9)
                $newitem | Add-Member -MemberType NoteProperty -Name "LivelinkID" -Value $value
            }
            else
            {
                $newitem | Add-Member -MemberType NoteProperty -Name "LivelinkID" -Value ""
            }
        }
    }
    catch
    {
        $newitem | Add-Member -MemberType NoteProperty -Name "LivelinkID" -Value ""
    }

    # if there is no Order Type field then it is an RFI
    try
    {
        $orderType = $newitem.'Order Type'
        $lem = $newitem.'LEM Value'
        if([string]::IsNullOrEmpty($orderType))
        {
            $newitem | Add-Member -MemberType NoteProperty -Name "DocumentType" -Value "ChangeRequest"
        }
        elseif($orderType -eq 'OM')
        {
            $newitem | Add-Member -MemberType NoteProperty -Name "DocumentType" -Value "Contract"
        }
        elseif( ($orderType -eq 'OC') -or ($orderType -eq 'OS') -or ($orderType -eq 'OP'))
        {
            if([string]::IsNullOrEmpty($lem))
            {
                $newitem | Add-Member -MemberType NoteProperty -Name "DocumentType" -Value "PO"
            }
            else
            {
                $newitem | Add-Member -MemberType NoteProperty -Name "DocumentType" -Value "Contract"
            }
        }
        else
        {
            $newitem | Add-Member -MemberType NoteProperty -Name "DocumentType" -Value ""
        }

    }
    catch
    {
        $newitem | Add-Member -MemberType NoteProperty -Name "DocumentType" -Value ""
    }


}


function ProcessList($jdelist)
{
    $newitems = @()
    $propertynames = $jdelist[0] | gm -MemberType NoteProperty
    foreach($JDEItem in $jdelist)
    {
       $newitem = New-Object -TypeName psobject
       
       foreach($propertyname in $propertynames)
       {
            $prop = $propertyname.Name
            $value = $JDEItem.$prop
            $newitem | Add-Member -MemberType NoteProperty -Name $propertyname.Name -Value $value
       }
       # now add additional fields
       AddAdditionalFields -newitem $newitem
       $newitems += $newitem
    }
    return $newitems

}

function CreateLookupSql1Clause($dataitemlist)
{
    $resultstring = @()
    $count = 1
    $lastidex = $dataitemlist.Length
    $resultstring += 'select dataid, name, getparentpath(dataid) from dtree where dataid in ('
    foreach($dataitem in $dataitemlist)
    {
        if($count -lt $lastidex)
        {
            $resultstring += $dataitem.LivelinkID + ','
        }
        else
        {
            $resultstring += $dataitem.LivelinkID
        }
        $count++
    }
    $resultstring += ')'
    return $resultstring
}

function CreateLookupSql2Clause($dataitemlist)
{
    $resultstring = @()
    $count = 1
    $lastidex = $dataitemlist.Length
    $resultstring += 'SELECT     dtree.dataid, dtree.NAME, getparentpath (dataid) "Path",
           ll_subtypes.NAME,
           (SELECT llattrdata.valstr
              FROM llattrdata
             WHERE llattrdata.ID = dtree.dataid AND llattrdata.attrid = 2)
      FROM (dtree LEFT JOIN ll_subtypes ON ll_subtypes.SUBTYPE = dtree.SUBTYPE)
	  WHERE dataid in ('
    foreach($dataitem in $dataitemlist)
    {
        if($count -lt $lastidex)
        {
            $resultstring += $dataitem.LivelinkID + ','
        }
        else
        {
            $resultstring += $dataitem.LivelinkID
        }
        $count++
    }
    $resultstring += ' )
CONNECT BY PRIOR dataid = parentid
START WITH dataid = [Folder DataID]
'
    return $resultstring
}

function Create-LookupSQL($configfilepath)
{
    $config = Import-Clixml -Path $configfilepath

    $ContractPoListFile = $config.ContractPoListFile
    $RFIListFile = $config.RFIListFile

    $ContractPOJDEList = Import-Csv ((Get-Location).Path + "\" + $ContractPoListFile)
    $RFIJDEList = Import-Csv ((Get-Location).Path + "\" + $RFIListFile)

    $ContractPOCompleteList = ProcessList -jdelist $ContractPOJDEList
    $RFICompleteList = ProcessList -jdelist $RFIJDEList

    CreateLookupSql1Clause -dataitemlist ( $ContractPOCompleteList | where { $_.DocumentType -eq 'Contract' } ) | Out-File -FilePath 'SQLOutputFiles\ContractLookupSql1.txt'
    CreateLookupSql1Clause -dataitemlist ( $ContractPOCompleteList | where { $_.DocumentType -eq 'PO' } ) | Out-File -FilePath 'SQLOutputFiles\POLookupSql1.txt'
    CreateLookupSql1Clause -dataitemlist $RFICompleteList | Out-File -FilePath 'SQLOutputFiles\RFILookupSql1.txt'

    
    CreateLookupSql2Clause -dataitemlist ( $ContractPOCompleteList | where { $_.DocumentType -eq 'Contract' } ) | Out-File -FilePath 'SQLOutputFiles\ContractLookupSql2.txt'
    CreateLookupSql2Clause -dataitemlist ( $ContractPOCompleteList | where { $_.DocumentType -eq 'PO' } ) | Out-File -FilePath 'SQLOutputFiles\POLookupSql2.txt'
    CreateLookupSql2Clause -dataitemlist $RFICompleteList | Out-File -FilePath 'SQLOutputFiles\RFILookupSql2.txt'


}

Create-LookupSQL -configfilepath 'CreateBulkLoadSheet.Config.xml'