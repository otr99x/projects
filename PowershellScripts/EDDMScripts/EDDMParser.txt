#modify these two values for your environment
param
(
    [string]$workingdir="d:\work",
    [string]$zipfolder="d:\work\zipfiles"
)

#put excluded file types here

$exclusionfilter = @(".xls", ".xlsx")

if (-not (test-path "$env:ProgramFiles\7-Zip\7z.exe")) {throw "$env:ProgramFiles\7-Zip\7z.exe needed"} 
set-alias sz "$env:ProgramFiles\7-Zip\7z.exe"

$homedir = Convert-Path .
$files = Get-ChildItem $workingdir -Filter "*.dwg"
foreach($file in $files)
{
    $basefilename = $file.BaseName
    $id = $basefilename.LastIndexOf("-");
    if($id -ge 1)
    {
        $zipfilename = $file.BaseName + ".zip"
        $prefix = $baseFilename.Substring(0,$id)
        # get all the items with the basefilename
        $folderitems = Get-ChildItem $workingdir -Filter ($baseFilename + "*") | Where {$exclusionfilter -notcontains $_.Extension}
        
        if($folderitems.Count -eq 1)
        {
            Copy-Item $folderitems[0].FullName ($zipfolder + "\" + $folderitems[0].BaseName + $folderitems[0].Extension)
            Set-Location $homedir
        }
        else
        {
            # create a temporary folder to put the files into
            $tempfolder = $workingdir + "\" + $baseFilename
            $zipfilename = $tempfolder + "\" + $baseFilename + ".zip"
            New-Item -ItemType directory -Path $tempfolder

            foreach($folderitem in $folderitems)
            {
                if($folderitem.Extension -eq ".dwg")
                {
                   $destinationpath = $tempfolder + "\" + $folderitem.Name 
                }
                else
                {
                    $destinationpath = $tempfolder + "\" + $prefix + $folderitem.Extension
                }
                Copy-Item $folderitem.FullName $destinationpath
            }
            Set-Location $tempfolder
            #zip $zipfilename "*"
            sz a -t7z $zipfilename "*"
            Set-Location $homedir
            Copy-Item $zipfilename $zipfolder
            # now remove the tempfolder
            Remove-Item -Recurse -Force $tempfolder

        }
    }
}

Set-Location $zipfolder
Get-ChildItem | ForEach-Object {$_.Name | Add-Content manifest.txt}
set-location $homedir