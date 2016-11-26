

function LoadLookup($hashtable, $lookuplist, $keyname)
{
    foreach($item in $lookuplist)
    {
        try
        {
            if(![string]::IsNullOrEmpty($item.$keyname))
            {
                $hashtable.Add($item.$keyname, $item)  
            }
        }
        catch
        {
            'Hash fail {0}' -f $item.$keyname | Out-Host
        }
    }
}

function AddAdditionalFieldsJDE($newitem, $hashlookup)
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
        $lookupitem = $hashlookup.Item($newitem.LivelinkID)
        $newitem | Add-Member -MemberType NoteProperty -Name "Path" -Value $lookupitem.path
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

function ProcessListJDE($jdelist, $lookuphash)
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
       AddAdditionalFieldsJDE -newitem $newitem -hashlookup $lookuphash
       $newitems += $newitem
    }
    return $newitems

}


function AddAdditionalFields($newitem, $hashlookup, $additionalfieldtype)
{
    if($additionalfieldtype -eq 'ContentServer')
    {
        #add SP-Dataid-lookup field containing the SP Item Type in the SPlookupHash
        try
        {
            $hashitem = $hashlookup.Item($newitem.DATAID)
            $newitem | Add-Member -MemberType NoteProperty -Name "sp-dataid-lookup" -Value $hashitem.'Item Type'
        }
        catch
        {
            $newitem | Add-Member -MemberType NoteProperty -Name "sp-dataid-lookup" -Value ""
        }

    }
    elseif($additionalfieldtype -eq 'Sharepoint')
    {
        # modify the URL field to add the webapp url to the front
        if(![string]::IsNullOrEmpty( $newItem.Path ) )
        {
            $newitem.Path = $sharepointwebappurl + $newitem.Path
        }
         #add CS-Dataid-lookup field containing the CS SubType in the CSlookuphash
        try
        {
            $hashitem = $hashlookup.Item($newitem.'Livelink ID')
            $newitem | Add-Member -MemberType NoteProperty -Name "cs-dataid-lookup" -Value $hashitem.Subtype
        }
        catch
        {
            $newitem | Add-Member -MemberType NoteProperty -Name "cs-dataid-lookup" -Value ""
        }

   }

}

function ProcessList($datalist, $lookuphash, $additionalfieldtype)
{
    $newitems = @()
    $propertynames = $datalist[0] | gm -MemberType NoteProperty
    foreach($dataitem in $datalist)
    {
       $newitem = New-Object -TypeName psobject
       
       foreach($propertyname in $propertynames)
       {
            $prop = $propertyname.Name
            $value = $dataitem.$prop
            $newitem | Add-Member -MemberType NoteProperty -Name $propertyname.Name -Value $value
       }
       # now add additional fields
       AddAdditionalFields -newitem $newitem -hashlookup $lookuphash -additionalfieldtype $additionalfieldtype
       $newitems += $newitem
    }
    return $newitems

}
function Run-Report ($summaryinputfolder, $sharepointwebappurl)
{
    #Load each of the csv files based on naming convention
    $TempJDEContractPOSheet = Import-Csv -Path ($summaryinputfolder + '\JDEContractPO.csv')

    $TempJDEContract = $TempJDEContractPOSheet | where -FilterScript { ($_.'Order Type' -eq 'OM') -or ($_.'LEM Value' -eq 'LEM') }
    $TempJDEPO = $TempJDEContractPOSheet | where -FilterScript { ($_.'Order Type' -in 'OC','OS','OP') -and ($_.'LEM Value' -ne 'LEM') }
    $TempJDERFI = Import-Csv -Path ($summaryinputfolder + '\JDERFI.csv')

    $TempCSContract = Import-Csv -Path ($summaryinputfolder + '\CSContract.csv')
    $TempCSPO = Import-Csv -Path ($summaryinputfolder + '\CSPO.csv')
    $TempCSRFI = Import-Csv -Path ($summaryinputfolder + '\CSRFI.csv')

    $TempSPContract = Import-Csv -Path ($summaryinputfolder + '\SPContract.csv')
    $TempSPPO = Import-Csv -Path ($summaryinputfolder + '\SPPO.csv')
    $TempSPRFI = Import-Csv -Path ($summaryinputfolder + '\SPRFI.csv')

    $csContractlookup = @{}
    $csPOlookup = @{}
    $csRFIlookup = @{}
    $spContractlookup = @{}
    $spPOlookup = @{}
    $spRFIlookup = @{}

    LoadLookup -hashtable $csContractlookup -lookuplist $TempCSContract -keyname 'DATAID'
    LoadLookup -hashtable $csPOlookup -lookuplist $TempCSPO -keyname 'DATAID'
    LoadLookup -hashtable $csRFIlookup -lookuplist $TempCSRFI -keyname 'DATAID'

    LoadLookup -hashtable $spContractlookup -lookuplist $TempSPContract -keyname 'Livelink ID'
    LoadLookup -hashtable $spPOlookup -lookuplist $TempSPPO -keyname 'Livelink ID'
    LoadLookup -hashtable $spRFIlookup -lookuplist $TempSPRFI -keyname 'Livelink ID'

    $JDEContract = ProcessListJDE -jdelist $TempJDEContract -lookuphash $csContractLookup
    $JDEPO = ProcessListJDE -jdelist $TempJDEPO -lookuphash $csPOlookup
    $JDERFI = ProcessListJDE -jdelist $TempJDERFI -lookuphash $csRFIlookup

    $CSContract = ProcessList -datalist $TempCSContract -lookuphash $spContractlookup -additionalfieldtype 'ContentServer'
    $CSPO = ProcessList -datalist $TempCSPO -lookuphash $spPOlookup -additionalfieldtype 'ContentServer'
    $CSRFI = ProcessList -datalist $TempCSRFI -lookuphash $spRFIlookup -additionalfieldtype 'ContentServer'

    $SPContract = ProcessList -datalist $TempSPContract -lookuphash $csContractlookup -additionalfieldtype 'Sharepoint'
    $SPPO = ProcessList -datalist $TempSPPO -lookuphash $csPOlookup -additionalfieldtype 'Sharepoint'
    $SPRFI = ProcessList -datalist $TempSPRFI -lookuphash $csRFIlookup -additionalfieldtype 'Sharepoint'

    $JDEContract | Export-Csv -Path "SummaryOutputFiles\JDEContract.csv" -NoTypeInformation
    $JDEPO | Export-Csv -Path "SummaryOutputFiles\JDEPO.csv" -NoTypeInformation
    $JDERFI | Export-Csv -Path "SummaryOutputFiles\JDERFI.csv" -NoTypeInformation
    $CSContract | Export-Csv -Path "SummaryOutputFiles\CSContract.csv" -NoTypeInformation
    $CSPO | Export-Csv -Path "SummaryOutputFiles\CSPO.csv" -NoTypeInformation
    $CSRFI | Export-Csv -Path "SummaryOutputFiles\CSRFI.csv" -NoTypeInformation
    $SPContract | Export-Csv -Path "SummaryOutputFiles\SPContract.csv" -NoTypeInformation
    $SPPO | Export-Csv -Path "SummaryOutputFiles\SPPO.csv" -NoTypeInformation
    $SPRFI | Export-Csv -Path "SummaryOutputFiles\SPRFI.csv" -NoTypeInformation

    # create the reporting object
    $Report = New-Object -TypeName psobject

    $Report | Add-Member -MemberType NoteProperty -Name 'ContractJDETotal' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractJDEMigratable' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractJDEPercent' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractSPRecordsMigrated' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractSPFilesMigrated' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractRecordsInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractFilesInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractRecordsInCSNotInSP' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractItemsInSPNotInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractItemsInCSNotInSP' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractURLInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractMigratedURLInSP' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractMigratedFilePercent' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractMigratedRecordPercent' -Value 0
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractMigrationEfficiency' -Value 0
    $Report | Add-Member -MemberType NoteProperty -Name 'SEPERATOR1' -Value '-----------------------------------------------------------------------'
    
    $Report | Add-Member -MemberType NoteProperty -Name 'POJDETotal' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POJDEMigratable' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POJDEPercent' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POSPRecordsMigrated' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POSPFilesMigrated' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'PORecordsInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POFilesInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'PORecordsInCSNotInSP' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POItemsInSPNotInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POItemsInCSNotInSP' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POURLInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POMigratedURLInSP' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POMigratedFilePercent' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POMigratedRecordPercent' -Value 0
    $Report | Add-Member -MemberType NoteProperty -Name 'POMigrationEfficiency' -Value 0
    $Report | Add-Member -MemberType NoteProperty -Name 'SEPERATOR2' -Value '-----------------------------------------------------------------------'
          
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIJDETotal' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIJDEMigratable' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIJDEPercent' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFISPRecordsMigrated' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFISPFilesMigrated' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIRecordsInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIFilesInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIRecordsInCSNotInSP' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIItemsInSPNotInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIItemsInCSNotInSP' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIURLInCS' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIMigratedURLInSP' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIMigratedFilePercent' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIMigratedRecordPercent' -Value 0
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIMigrationEfficiency' -Value 0

    
    #Contracts
    $JDEList = $JDEContract
    $SPList = $SPContract
    $CSList = $CSContract
    $DocContentType = 'Contracts'
    $DocSetContentType = 'Contracts Document Set'
    $CSDocumentPath = ':Enterprise:Upstream Operations:Upstream Business Services:Supply Management:JDE Attachments:SCM-CCA-Contracts'
        
    $Report.ContractJDETotal = @($JDEList).count
    $Report.ContractJDEMigratable= @($JDEList | Where-Object -Property Path -Like ($CSDocumentPath + '*')).count
    $Report.ContractJDEPercent = "{0:p2}" -f ($Report.ContractJDEMigratable / $Report.ContractJDETotal)
    $Report.ContractSPRecordsMigrated = @($SPList | Where-Object -Property 'Content Type' -EQ $DocSetContentType | Where-Object -Property 'Livelink ID' -GT 0).Count
    $Report.ContractSPFilesMigrated = @($SPList | Where-Object -Property 'Content Type' -EQ $DocContentType | Where-Object -Property 'LiveLink ID' -GT 0).Count + @($SPList | Where-Object -Property 'Content Type' -EQ 'Document' | Where-Object -Property 'LiveLink ID' -GT 0).Count
    $Report.ContractRecordsInCS = @($CSList | Where-Object -Property Path -EQ $CSDocumentPath).count
    $Report.ContractFilesInCS = @($CSList | Where-Object -Property Subtype -EQ 'Document').Count + @($CSList | Where-Object -Property Subtype -EQ '' | Where-Object -Property DATAID -GT 0).Count
    $Report.ContractRecordsInCSNotInSP = @($CSList | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ 'Folder' | Where-Object -Property Path -EQ $CSDocumentPath).Count
    $Report.ContractItemsInSPNotInCS = @($SPList | Where-Object -Property cs-dataid-lookup -EQ '' | Where-Object -Property 'Item Type' -NE 'Folder' | Where-Object -Property 'Item Type' -NE 'URL' | Where-Object -Property DATAID -GT 0).Count
    $Report.ContractItemsInCSNotInSP = @($CSList | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ 'Document').Count + @($CSList | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ '').Count
    $Report.ContractURLInCS = @($CSList | Where-Object -Property Subtype -EQ 'URL').Count
    $Report.ContractMigratedURLInSP = @($SPList | Where-Object -Property 'Content Type' -EQ 'Link to a Document').Count
    $Report.ContractMigratedFilePercent = "{0:p2}" -f ($Report.ContractSPFilesMigrated / $Report.ContractFilesInCS)
    $Report.ContractMigratedRecordPercent = "{0:p2}" -f ($Report.ContractSPRecordsMigrated / $Report.ContractJDEMigratable)
    $Report.ContractMigrationEfficiency = "{0:p2}" -f (($Report.ContractJDEMigratable / $Report.ContractJDETotal) * ($Report.ContractSPRecordsMigrated / $Report.ContractJDEMigratable))

    #PO
    $JDEList = $JDEPO
    $SPList = $SPPO
    $CSList = $CSPO
    $DocContentType = 'Purchase Orders'
    $DocSetContentType = 'Purchase Orders Document Set'
    $CSDocumentPath = ':Enterprise:Upstream Operations:Upstream Business Services:Supply Management:JDE Attachments:SCM-CCA-PurchaseOrders'

    $Report.POJDETotal = @($JDEList).count
    $Report.POJDEMigratable= @($JDEList | Where-Object -Property Path -Like ($CSDocumentPath + '*')).count
    $Report.POJDEPercent = "{0:p2}" -f ($Report.POJDEMigratable / $Report.POJDETotal)
    $Report.POSPRecordsMigrated = @($SPList | Where-Object -Property 'Content Type' -EQ $DocSetContentType | Where-Object -Property 'Livelink ID' -GT 0).Count
    $Report.POSPFilesMigrated = @($SPList | Where-Object -Property 'Content Type' -EQ $DocContentType | Where-Object -Property 'LiveLink ID' -GT 0).Count + @($SPList | Where-Object -Property 'Content Type' -EQ 'Document' | Where-Object -Property 'LiveLink ID' -GT 0).Count
    $Report.PORecordsInCS = @($CSList | Where-Object -Property Path -EQ $CSDocumentPath).count
    $Report.POFilesInCS = @($CSList | Where-Object -Property Subtype -EQ 'Document').Count + @($CSList | Where-Object -Property Subtype -EQ '' | Where-Object -Property DATAID -GT 0).Count
    $Report.PORecordsInCSNotInSP = @($CSList | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ 'Folder' | Where-Object -Property Path -EQ $CSDocumentPath).Count
    $Report.POItemsInSPNotInCS = @($SPList | Where-Object -Property cs-dataid-lookup -EQ '' | Where-Object -Property 'Item Type' -NE 'Folder' | Where-Object -Property 'Item Type' -NE 'URL' | Where-Object -Property DATAID -GT 0).Count
    $Report.POItemsInCSNotInSP = @($CSList | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ 'Document').Count + @($CSList | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ '').Count
    $Report.POURLInCS = @($CSList | Where-Object -Property Subtype -EQ 'URL').Count
    $Report.POMigratedURLInSP = @($SPList | Where-Object -Property 'Content Type' -EQ 'Link to a Document').Count
    $Report.POMigratedFilePercent = "{0:p2}" -f ($Report.POSPFilesMigrated / $Report.POFilesInCS)
    $Report.POMigratedRecordPercent = "{0:p2}" -f ($Report.POSPRecordsMigrated / $Report.POJDEMigratable)
    $Report.POMigrationEfficiency = "{0:p2}" -f (($Report.POJDEMigratable / $Report.POJDETotal) * ($Report.POSPRecordsMigrated / $Report.POJDEMigratable))
    
    #RFI
    $JDEList = $JDERFI
    $SPList = $SPRFI
    $CSList = $CSRFI
    $DocContentType = 'Change Requests'
    $DocSetContentType = 'Change Requests Document Set'
    $CSDocumentPath = ':Enterprise:Upstream Operations:Upstream Business Services:Supply Management:JDE Attachments:SCM-CCA-RFI'

    $Report.RFIJDETotal = @($JDEList).count
    $Report.RFIJDEMigratable= @($JDEList | Where-Object -Property Path -Like ($CSDocumentPath + '*')).count
    $Report.RFIJDEPercent = "{0:p2}" -f ($Report.RFIJDEMigratable / $Report.RFIJDETotal)
    $Report.RFISPRecordsMigrated = @($SPList | Where-Object -Property 'Content Type' -EQ $DocSetContentType | Where-Object -Property 'Livelink ID' -GT 0).Count
    $Report.RFISPFilesMigrated = @($SPList | Where-Object -Property 'Content Type' -EQ $DocContentType | Where-Object -Property 'LiveLink ID' -GT 0).Count + @($SPList | Where-Object -Property 'Content Type' -EQ 'Document' | Where-Object -Property 'LiveLink ID' -GT 0).Count
    $Report.RFIRecordsInCS = @($CSList | Where-Object -Property Path -EQ $CSDocumentPath).count
    $Report.RFIFilesInCS = @($CSList | Where-Object -Property Subtype -EQ 'Document').Count + @($CSList | Where-Object -Property Subtype -EQ '' | Where-Object -Property DATAID -GT 0).Count
    $Report.RFIRecordsInCSNotInSP = @($CSList | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ 'Folder' | Where-Object -Property Path -EQ $CSDocumentPath).Count
    $Report.RFIItemsInSPNotInCS = @($SPList | Where-Object -Property cs-dataid-lookup -EQ '' | Where-Object -Property 'Item Type' -NE 'Folder' | Where-Object -Property 'Item Type' -NE 'URL' | Where-Object -Property DATAID -GT 0).Count
    $Report.RFIItemsInCSNotInSP = @($CSList | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ 'Document').Count + @($CSList | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ '').Count
    $Report.RFIURLInCS = @($CSList | Where-Object -Property Subtype -EQ 'URL').Count
    $Report.RFIMigratedURLInSP = @($SPList | Where-Object -Property 'Content Type' -EQ 'Link to a Document').Count
    $Report.RFIMigratedFilePercent = "{0:p2}" -f ($Report.RFISPFilesMigrated / $Report.RFIFilesInCS)
    $Report.RFIMigratedRecordPercent = "{0:p2}" -f ($Report.RFISPRecordsMigrated / $Report.RFIJDEMigratable)
    $Report.RFIMigrationEfficiency = "{0:p2}" -f (($Report.RFIJDEMigratable / $Report.RFIJDETotal) * ($Report.RFISPRecordsMigrated / $Report.RFIJDEMigratable))


    $Report

}

Run-Report -summaryinputfolder 'SummaryInputFiles' -sharepointwebappurl 'https://teamsiteppd.cenovus.com/'