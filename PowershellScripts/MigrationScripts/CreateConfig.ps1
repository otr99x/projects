$configitem = New-Object -TypeName psobject
$configitem | Add-Member -MemberType NoteProperty -Name 'ContractPoListFile' -Value 'InputFiles\ContractPoList.csv'
$configitem | Add-Member -MemberType NoteProperty -Name 'ContractPoLookupFile' -Value 'InputFiles\ContractPoLookup.csv'
$configitem | Add-Member -MemberType NoteProperty -Name 'RFIListFile' -Value 'InputFiles\RFIList.csv'
$configitem | Add-Member -MemberType NoteProperty -Name 'RFILookupFile' -Value 'InputFiles\RFILookup.csv'
$configitem | Add-Member -MemberType NoteProperty -Name 'RFICollectionID' -Value 'RFI_COLLECTION_ID'
$configitem | Add-Member -MemberType NoteProperty -Name 'POCollectionID' -Value 'PO_COLLECTION_ID'
$configitem | Add-Member -MemberType NoteProperty -Name 'ContractsConfidentialCollectionID' -Value 'CONTRACT_CONFIDENTIAL_COLLECTION_ID'
$configitem | Add-Member -MemberType NoteProperty -Name 'ContractsNonConfidentialCollectionID' -Value 'CONTRACT_NON_CONFIDENTIAL_COLLECTION_ID'

$configitem | Export-Clixml -Path 'CreateBulkLoadSheet.Config.xml'