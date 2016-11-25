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
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractMigratedContractFilePercent' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'ContractMigratedContractRecordPercent' -Value 0
    
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
    $Report | Add-Member -MemberType NoteProperty -Name 'POMigratedContractFilePercent' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'POMigratedContractRecordPercent' -Value 0
          
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
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIMigratedContractFilePercent' -Value 0      
    $Report | Add-Member -MemberType NoteProperty -Name 'RFIMigratedContractRecordPercent' -Value 0


    $Report.ContractJDETotal = @($JDEContract).count
    $Report.ContractJDEMigratable= @($JDEContract | Where-Object -Property Path -Like ':Enterprise:Upstream Operations:Upstream Business Services:Supply Management:JDE Attachments:SCM-CCA-Contracts*').count
    $Report.ContractJDEPercent = "{0:p2}" -f ($Report.ContractJDEMigratable / $Report.ContractJDETotal)
    $Report.ContractSPRecordsMigrated = @($SPContract | Where-Object -Property 'Content Type' -EQ 'Contracts Document Set' | Where-Object -Property 'Livelink ID' -GT 0).Count
    $Report.ContractSPFilesMigrated = @($SPContract | Where-Object -Property 'Content Type' -EQ 'Contracts' | Where-Object -Property 'LiveLink ID' -GT 0).Count + @($SPContract | Where-Object -Property 'Content Type' -EQ 'Document' | Where-Object -Property 'LiveLink ID' -GT 0).Count
    $Report.ContractRecordsInCS = @($CSContract | Where-Object -Property Path -EQ ':Enterprise:Upstream Operations:Upstream Business Services:Supply Management:JDE Attachments:SCM-CCA-Contracts').count
    $Report.ContractFilesInCS = @($CSContract | Where-Object -Property Subtype -EQ 'Document').Count + @($CSContract | Where-Object -Property Subtype -EQ '' | Where-Object -Property DATAID -GT 0).Count
    $Report.ContractRecordsInCSNotInSP = @($CSContract | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ 'Folder' | Where-Object -Property Path -EQ ':Enterprise:Upstream Operations:Upstream Business Services:Supply Management:JDE Attachments:SCM-CCA-Contracts').Count
    $Report.ContractItemsInSPNotInCS = @($SPContract | Where-Object -Property cs-dataid-lookup -EQ '' | Where-Object -Property 'Item Type' -NE 'Folder' | Where-Object -Property 'Item Type' -NE 'URL' | Where-Object -Property DATAID -GT 0).Count
    $Report.ContractItemsInCSNotInSP = @($CSContract | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ 'Document').Count + @($CSContract | Where-Object -Property sp-dataid-lookup -EQ '' | Where-Object -Property Subtype -EQ '').Count
    $Report.ContractURLInCS = @($CSContract | Where-Object -Property Subtype -EQ 'URL').Count
    $Report.ContractMigratedURLInSP = @($SPContract | Where-Object -Property 'Content Type' -EQ 'Link to a Document').Count
    $Report.ContractMigratedContractFilePercent = "{0:p2}" -f ($Report.ContractSPFilesMigrated / $Report.ContractFilesInCS)
    $Report.ContractMigratedContractRecordPercent = "{0:p2}" -f ($Report.ContractSPRecordsMigrated / $Report.ContractJDEMigratable)

    #TODO build report object for PO and RFI

    $Report
