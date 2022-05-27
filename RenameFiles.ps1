param (
        [Parameter()]
        [String]$sheet = 'Example.xlsx',
        [String]$directory = $pwd,
        [String[]]$extensions = ('*.jpg', '*.jpeg', '*.png', '*.svg', '*.txt')
    )

try {
    Import-Module ImportExcel;
    Write-Host "`nModule installed successfully :)`n" -ForegroundColor green;
} catch {
    Write-Host "Missing module please install using the following command..." -ForegroundColor red;
    Write-Host "Install-Module -Name ImportExcel" -ForegroundColor blue
    throw;
    return;
}

Write-Host "
___________.__.__        __________                                           
\_   _____/|__|  |   ____\______   \ ____   ____ _____    _____   ___________ 
 |    __)  |  |  | _/ __ \|       _// __ \ /    \\__  \  /     \_/ __ \_  __ \
 |     \   |  |  |_\  ___/|    |   \  ___/|   |  \/ __ \|  Y Y  \  ___/|  | \/
 \___  /   |__|____/\___  >____|_  /\___  >___|  (____  /__|_|  /\___  >__|   
     \/                 \/       \/     \/     \/     \/      \/     \/    
" -ForegroundColor magenta

function Rename-Files {
    cd $directory;
    $success = $false;

    function Confirm-Prompt {
        $confirmation = Read-Host "Confirm? [y/n]"
            while($confirmation -ne "y")
            {
                if ($confirmation -eq 'n') {cd $directory; exit;}
                $confirmation = Read-Host "Confirm? [y/n]"
            }
    }

    function Check-Path {
        try {    
            Write-Host "`n`nRenamed already generated from a previous run`n`n" -ForegroundColor yellow;
            cd $directory;
            Write-Host "Ready to cleanup previous run..."
            Confirm-Prompt;
            rmdir "$directory\Renamed";
            Write-Host "`n`nCleanup done. Please run script again`n`n" -ForegroundColor green;
        } catch {
            cd $directory;
            throw;
            return;
        }
    }

    function Copy-Children {
        try {
            mkdir 'Renamed';
            cd 'Renamed';
            Get-ChildItem "$directory\*" -Include $extensions | ForEach-Object -Process {
                Copy-Item $_ -Recurse -Verbose;
            }
            Write-Host "`nOriginal Contents`n" -ForegroundColor green;
            ls $directory;
            Write-Host "`nRenamed Contents" -ForegroundColor yellow;
            ls "$directory\Renamed";
        } catch {
            cd $directory;
            throw;
            return
        }
    }

    function Rename-Children {
        try {
            Write-Host "`nReady to rename files...`n";
            Confirm-Prompt;                           
            $data | ForEach-Object {
                Rename-Item -Path $_.Old -NewName $_.New;
            }
        } catch {
            cd $directory;
            throw;
            return
        }
    }

    $f = @{};
    $data = Import-Excel $sheet;
    $data |
        ForEach-Object{
            $f[$_.Old] = $_.New
        };

    if( Test-Path "$directory\Renamed" ) {
        Check-Path;
    } else {
        Copy-Children;
        Rename-Children;
        $success = $true;
    }

    if ($success) {
        
        cd $directory;

        Write-Host "`nOperation complete below are your renamed files at [$directory\Renamed]`n" -ForegroundColor green;
        ls "$directory\Renamed";

        Write-Host "`nDont worry the old files are still there just in case";
        ls $directory;
    }

    cd $directory;
}

try { Rename-Files $sheet $directory } catch {cd $directory; throw; return}