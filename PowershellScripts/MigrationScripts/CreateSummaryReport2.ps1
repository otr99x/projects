    $JDEContract = Import-Csv -Path "SummaryOutputFiles\JDEContract.csv" 
    $JDEPO = Import-Csv -Path "SummaryOutputFiles\JDEPO.csv"
    $JDERFI = Import-Csv -Path "SummaryOutputFiles\JDERFI.csv"
    $CSContract = Import-Csv -Path "SummaryOutputFiles\CSContract.csv"
    $CSPO = Import-Csv -Path "SummaryOutputFiles\CSPO.csv"
    $CSRFI = Import-Csv -Path "SummaryOutputFiles\CSRFI.csv"
    $SPContract = Import-Csv -Path "SummaryOutputFiles\SPContract.csv"
    $SPPO = Import-Csv -Path "SummaryOutputFiles\SPPO.csv"
    $SPRFI = Import-Csv -Path "SummaryOutputFiles\SPRFI.csv"
    
    
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
