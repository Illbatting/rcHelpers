Set-StrictMode -Version 2.0

# Simple Timestamping
function timeStamp {
    Write-Output $([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss.fff'))
}

# Save xlsx-file as a csv, using an Excel.Application COM-object.
function xlsxToCsv {
    param ([string[]]$OriginalFile, [switch]$DeleteOriginal)
    process
    {
        foreach($file in $OriginalFile)
        {
            if($file.GetType().FullName -ne "System.IO.FileInfo")
            {
                $file = Get-ChildItem $file
            }

            $csvFile = "$($file.DirectoryName)\$($file.BaseName).csv"

            Write-Output "$(timeStamp) `t Re-saving file $($file.FullName) as .csv"

            $excel = New-Object -ComObject Excel.Application
            $wb = $excel.WorkBooks.Open($file.FullName)
            $wb.SaveAs($csvFile,6)
            $wb.Close($false)
            $excel.Quit()

            if($DeleteOriginal -eq $true)
            {
                if(Test-Path $csvFile)
                {
                    Remove-Item -Path $file.FullName -Confirm
                    Write-Output "$(timeStamp) `t Deleted file $($file.FullName)"
                }
                else
                {
                    Write-Output "$(timeStamp) `t Could not find the file $($csvFile)"
                    Write-Output "$(timeStamp) `t Sure you want to delete the original?"
                    Remove-Item -Path $file.FullName -Confirm
                }
            }


            # Attempt to release the COM-object.
            $releaseStatus = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
            if($releaseStatus -eq 0)
            {
                Write-Output "$(timeStamp) `t The Excel ComObject was successfully released. Exitcode: $releaseStatus"
            }
            else
            {
                Write-Output "$(timeStamp) `t The Excel ComObject could not be released. Exitcode: $releaseStatus"
            }

        }
    }
}