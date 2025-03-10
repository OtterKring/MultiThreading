function ThreadedProcessing {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Path = ( Get-Location ).Path + '\FakeNameGenerator.com_71e0a582.csv'
    )

    if ( $PSVersionTable.PSVersion.Major -gt 5 ) {

        if ( Test-Path $Path -ErrorAction Stop ) {

            cls

            $ThreadVerbosity = $VerbosePreference
            Write-Verbose ( "PSVersion {0} detected" -f $PSVersionTable.PSVersion )

            $Threads = [math]::Floor( ( Get-CimInstance -Class Win32_Processor 4>$null ).NumberOfCores * 1.5 )
            $InputQueue = [System.Collections.Concurrent.ConcurrentQueue[psobject]]::new()
            $OutputQueue = [System.Collections.Concurrent.ConcurrentQueue[psobject]]::new()
            $InputCompleted = [System.Collections.Concurrent.ConcurrentQueue[System.Boolean]]@($false)

            Write-Verbose "calculated Threads: $Threads"

            1..$Threads | Foreach-Object -ThrottleLimit $Threads -Parallel {

                $VerbosePreference = $using:ThreadVerbosity

                $threadNum = $_
                $inputFile = $using:Path
                $inputData = $using:InputQueue
                $outputData = $using:OutputQueue
                $inputDone = $using:InputCompleted

                if ( $threadNum -eq 1 ) {

                    Write-Verbose "Thread $threadNum" -Verbose
                    Write-Verbose "InputFile: $inputFile" -Verbose

                    $streamReader = [System.IO.StreamReader]::new( $inputFile )

                    $headers = $streamReader.ReadLine().Split(',')

                    if ( $? -and $headers ) {

                        $item = $streamReader.ReadLine()
                        $lineCount = 1

                        # increase value for lineCount to read more lines from file
                        while ( $item -and $lineCount -le 1000 ) {

                            $item = ConvertFrom-Csv -InputObject $item -Header $headers
                            $inputData.Enqueue( $item )
                            $item = $streamReader.ReadLine()
                            $lineCount++

                        }

                        $streamReader.Dispose()
                        while ( -not $inputDone.TryDequeue( [ref]$null ) ) {}
                        $inputDone.Enqueue( $true )
                        Write-Verbose 'Input done'

                    }

                } elseif ( $threadNum -gt 1 ) {
                    
                    Write-Verbose "Thread $threadNum" -Verbose

                    $idone = $null
                    while ( -not $inputDone.TryPeek( [ref]$idone ) ) {}
                    while ( -not $idone -and $inputData.Count -eq 0 ) {
                        Start-Sleep -Milliseconds 50
                        while ( -not $inputDone.TryPeek( [ref]$idone ) ) {}
                    }

                    $item = $null
                    while ( $inputData.Count -gt 0 -and -not $inputData.TryDequeue( [ref] $item ) ) {}
                    while ( $item -or -not $idone -or ( $idone -and $inputData.Count -gt 0 ) ) {

                        if ( $item ) {
                            $item = [pscustomobject]@{
                                Thread = $threadNum
                                Number = $item.Number
                                Gender = $item.Gender
                                FirstName = $item.Givenname
                                LastName = $item.Surname
                                BoD = $item.Birthday
                            }
                            $null = $outputData.Enqueue( $item )
                            $item = $null
                        } else {
                            Start-Sleep -Milliseconds 50
                        }

                        Write-Verbose ( "Thread {0} : inputDone = {1} : dataCount = {2} : itemCount = {3}" -f $threadNum, [int]$idone, $inputData.Count, $item.Count )

                        while ( $inputData.Count -gt 0 -and -not $inputData.TryDequeue( [ref] $item ) ) {}
                        while ( -not $inputDone.TryPeek( [ref]$idone ) ) {}

                    }

                }

            }

            # $InputData
            $OutputQueue #| Group-Object thread -NoElement | Measure-Object Count -Sum | select -ExpandProperty sum
        
        } else {
            Write-Error -Message "Invalid path [$Path]" -ErrorAction Stop
        }

    } else {
        Write-Error -Message "Powershell version 6 or higher required to run this function" -ErrorAction Stop
    }

}