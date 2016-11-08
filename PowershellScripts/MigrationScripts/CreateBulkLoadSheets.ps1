function LoadLookupHash($hashtable, $lookuplist)
{
    foreach($item in $lookuplist)
    {
        if(($item.Path).Contains('RFIInProgress') -or ($item.Path).Contains('Orders in Progress'))
        {
            $hashtable.Add($item.DATAID, $item.Path)  
        }
    }
}

function AddAdditionalFields($newitem, $hashlookup)
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

    #Get the path based on the lookup
    try
    {
        $pathlookup = $hashlookup.Item($newitem.LivelinkID)
        $newitem | Add-Member -MemberType NoteProperty -Name "Path" -Value $pathlookup
    }
    catch
    {
        $newitem | Add-Member -MemberType NoteProperty -Name "Path" -Value ""
    }


    #if contract, determine if confidential or non confidential
    try
    {
        if($newitem.DocumentType -eq "Contract")
        {
            $pathUppercase = ($newitem.Path).ToUpper()
            if([string]::IsNullOrEmpty($pathUppercase))
            {
                $newitem | Add-Member -MemberType NoteProperty -Name "Confidentiality" -Value ""
            }
            elseif(($pathUppercase).Contains("NON CONFIDENTIAL"))
            {
                $newitem | Add-Member -MemberType NoteProperty -Name "Confidentiality" -Value "Non Confidential"
            }
            else
            {
                $newitem | Add-Member -MemberType NoteProperty -Name "Confidentiality" -Value "Confidential"
            }

        }
    }
    catch
    {
        $newitem | Add-Member -MemberType NoteProperty -Name "Confidentiality" -Value ""
    }


}

function ProcessList($jdelist, $lookuphash)
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
       AddAdditionalFields -newitem $newitem -hashlookup $lookuphash
       $newitems += $newitem
    }
    return $newitems

}

function GenerateRFILoadSheet($datalist, $LoadSheetList, $CollectionID, $CollectionList)
{
    $workingdatalist = $datalist | Where-Object {(![string]::IsNullOrEmpty($_.LivelinkID)) -and (![string]::IsNullOrEmpty($_.Path))}

    foreach($item in $workingdatalist)
    {
        $LoadSheetItem = New-Object -TypeName PSObject
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name '$OBJECTID' -Value $item.LivelinkID
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name '$MetadataOnly' -Value ""
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Change Requests Document Set:IsAttachedCat" -Value "1"
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Change Requests Document Set:Contract No" -Value $item.'Order Number'
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Change Requests Document Set:Supplier Name" -Value $item.'Supplier Number'
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Change Requests Document Set:PCR No" -Value $item.'RFI Number'
        $suppressOutput = $LoadSheetList.Add($LoadSheetItem)

        $insertString = "insert into collections values({0},{1},null);" -f $CollectionID,$item.LivelinkID
        $suppressOutput = $CollectionList.Add($insertString)
    }
}

function GeneratePOLoadSheet($datalist, $LoadSheetList, $CollectionID, $CollectionList)
{
    $workingdatalist = $datalist | Where-Object {(![string]::IsNullOrEmpty($_.LivelinkID)) -and (![string]::IsNullOrEmpty($_.Path))}

    foreach($item in $workingdatalist)
    {
        $LoadSheetItem = New-Object -TypeName PSObject
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name '$OBJECTID' -Value $item.LivelinkID
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name '$MetadataOnly' -Value ""
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Purchase Orders Document Set:IsAttachedCat" -Value "1"
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Purchase Orders Document Set:PO No" -Value $item.'Order Number'
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Purchase Orders Document Set:Supplier Name" -Value $item.'Supplier Number'
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Purchase Orders Document Set:Contract Reference" -Value $item.'Master Contract'
        $suppressOutput = $LoadSheetList.Add($LoadSheetItem)

        $insertString = "insert into collections values({0},{1},null);" -f $CollectionID,$item.LivelinkID
        $suppressOutput = $CollectionList.Add($insertString)
    }
 }

function GenerateContractLoadSheet($datalist, $LoadSheetList, $ConfidentialCollectionID, $ConfidentialCollectionList, $NonConfidentialCollectionID, $NonConfidentialCollectionList)
{
    $workingdatalist = $datalist | Where-Object {(![string]::IsNullOrEmpty($_.LivelinkID)) -and (![string]::IsNullOrEmpty($_.Path))}

    foreach($item in $workingdatalist)
    {
        $LoadSheetItem = New-Object -TypeName PSObject
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name '$OBJECTID' -Value $item.LivelinkID
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name '$MetadataOnly' -Value ""
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Contracts Document Set:IsAttachedCat" -Value "1"
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Contracts Document Set:Contract No" -Value $item.'Order Number'
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Contracts Document Set:Supplier Name" -Value $item.'Supplier Number'
        $LoadSheetItem | Add-Member -MemberType NoteProperty -Name ":Enterprise:ShareLink:Contracts Document Set:Contract Reference" -Value $item.'Master Contract'
        $suppressOutput = $LoadSheetList.Add($LoadSheetItem)

        if($item.Confidentiality -eq 'Confidential')
        {
            $insertString = "insert into collections values({0},{1},null);" -f $ConfidentialCollectionID,$item.LivelinkID
            $suppressOutput = $ConfidentialCollectionList.Add($insertString)
        }
        elseif($item.Confidentiality -eq 'Non Confidential')
        {
            $insertString = "insert into collections values({0},{1},null);" -f $NonConfidentialCollectionID,$item.LivelinkID
            $suppressOutput = $NonConfidentialCollectionList.Add($insertString)
       }
    }
}

function CreateLookupSql($dataitemlist)
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

function Run-Migration( $configfilepath )
{
    $config = Import-Clixml -Path $configfilepath

    $ContractPoListFile = $config.ContractPoListFile
    $ContractPoLookupFile = $config.ContractPoLookupFile
    $RFIListFile = $config.RFIListFile
    $RFILookupFile = $config.RFILookupFile
    $RFICollectionID = $config.RFICollectionID
    $POCollectionID = $config.POCollectionID
    $ContractsConfidentialCollectionID = $config.ContractsConfidentialCollectionID
    $ContractsNonConfidentialCollectionID = $config.ContractsNonConfidentialCollectionID

    # if we can't load the lookups, then must generate sql to create the lookup lists
    $ContractPOLookupHash = @{}
    $RFILookupHash = @{}

    $lookupsLoaded = $true
    try
    {
        $ContractPOLookup = Import-Csv ((Get-Location).Path + "\" + $ContractPoLookupFile)
        $RFILookup = Import-Csv ((Get-Location).Path + "\" + $RFILookupFile)
        LoadLookupHash -hashtable $ContractPOLookupHash -lookuplist $ContractPOLookup
        LoadLookupHash -hashtable $RFILookupHash -lookuplist $RFILookup
    }
    catch
    {
        $lookupsLoaded = $false
    }

    $ContractPOJDEList = Import-Csv ((Get-Location).Path + "\" + $ContractPoListFile)
    $RFIJDEList = Import-Csv ((Get-Location).Path + "\" + $RFIListFile)


    #Process the list
    #$ContractPOCompleteList = ProcessList -jdelist ($ContractPOJDEList | select -first 500) -lookuphash $ContractPOLookupHash
    #$RFICompleteList = ProcessList -jdelist ($RFIJDEList | select -first 500) -lookuphash $RFILookupHash
    $TempContractPOCompleteList = ProcessList -jdelist $ContractPOJDEList -lookuphash $ContractPOLookupHash
    $RFICompleteList = ProcessList -jdelist $RFIJDEList -lookuphash $RFILookupHash

    # if lookups werent loaded, then just generate sql to get the lookups

    if($lookupsLoaded -eq $false)
    {
          (CreateLookupSql -dataitemlist $TempContractPOCompleteList) | Out-File -FilePath 'OutputFiles\ContractPOLookupSql.txt'
          (CreateLookupSql -dataitemlist $RFICompleteList) | Out-File -FilePath 'OutputFiles\RFILookupSql.txt'
    }
    else
    {
        $ContractCompleteList = $TempContractPOCompleteList | where {$_.DocumentType -eq 'Contract'}
        $POCompleteList = $TempContractPOCompleteList | where {$_.DocumentType -eq 'PO'}

        $RFILoadSheetList = New-Object System.Collections.ArrayList
        $POLoadSheetList = New-Object System.Collections.ArrayList
        $ContractLoadSheetList= New-Object System.Collections.ArrayList

        $RFICollectionList = New-Object System.Collections.ArrayList
        $POCollectionList = New-Object System.Collections.ArrayList
        $ContractConfidentialCollectionList = New-Object System.Collections.ArrayList
        $ContractNonConfidentialCollectionList = New-Object System.Collections.ArrayList

        GenerateRFILoadSheet -datalist $RFICompleteList -LoadSheetList $RFILoadSheetList -CollectionID $RFICollectionID -CollectionList $RFICollectionList
        $RFILoadSheetList| Export-Csv -Path 'OutputFiles\metadata-BulkUploadChangeRequest.csv' -NoTypeInformation

        GeneratePOLoadSheet -datalist $POCompleteList -LoadSheetList $POLoadSheetList -CollectionID $POCollectionID -CollectionList $POCollectionList
        $POLoadSheetList | Export-Csv -Path 'OutputFiles\metadata-BulkUploadPO.csv' -NoTypeInformation

        GenerateContractLoadSheet -datalist $ContractCompleteList -LoadSheetList $ContractLoadSheetList -ConfidentialCollectionID $ContractsConfidentialCollectionID -ConfidentialCollectionList $ContractConfidentialCollectionList -NonConfidentialCollectionID $ContractsNonConfidentialCollectionID -NonConfidentialCollectionList $ContractNonConfidentialCollectionList
        $ContractLoadSheetList | Export-Csv -Path 'OutputFiles\metadata-BulkUploadContract.csv' -NoTypeInformation


        #Generate Collection SQL
        $CollectionSQLFilename = 'OutputFiles\CollectionSql.txt'
        'Collection Query for Change Requests----------' | Out-File -FilePath $CollectionSQLFilename
        $RFICollectionList | Out-File -FilePath $CollectionSQLFilename -Append
        'Collection Query for PO----------' | Out-File -FilePath $CollectionSQLFilename -Append
        $POCollectionList | Out-File -FilePath $CollectionSQLFilename -Append
        'Collection Query for Confidential Contracts----------' | Out-File -FilePath $CollectionSQLFilename -Append
        $ContractConfidentialCollectionList | Out-File -FilePath $CollectionSQLFilename -Append
        'Collection Query for NonConfidential Contracts----------' | Out-File -FilePath $CollectionSQLFilename -Append
        $ContractNonConfidentialCollectionList | Out-File -FilePath $CollectionSQLFilename -Append

        $ContractCompleteList | export-csv -Path 'OutputFiles\workingcontract.csv' -NoTypeInformation
        $POCompleteList | export-csv -Path 'OutputFiles\workingpo.csv' -NoTypeInformation
        $RFICompleteList | export-csv -Path 'OutputFiles\workingrfi.csv' -NoTypeInformation
        $TempContractPOCompleteList | export-csv -Path 'OutputFiles\tempContractPoworking.csv' -NoTypeInformation
    }
}

Run-Migration -configfilepath 'CreateBulkLoadSheet.Config.xml'