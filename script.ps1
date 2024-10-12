Write-Host @"
____   ____    .__   __         .__         
\   \ /   /___ |  |_/  |______  |__| ____   
 \   Y   /  _ \|  |\   __\__  \ |  |/ ___\  
  \     (  <_> )  |_|  |  / __ \|  \  \___  
   \___/ \____/|____/__| (____  /__|\___  > 
                              \/        \/  
"@ -ForegroundColor Cyan

function Check-GCloudInstallation {
    try {
        gcloud --version | Out-Null
        Write-Host "Google Cloud SDK is available."
        return $true
    } catch {
        Write-Host "Google Cloud SDK is not installed or not configured properly." -ForegroundColor Yellow
        $installResponse = Get-YesNoInput "Would you like to install the Google Cloud SDK now? (Y/N)"
        if ($installResponse -eq 'Y') {
            Install-GCloudSDK
            return $false
        } else {
            Write-Host "Please install the Google Cloud SDK manually and then rerun the script." -ForegroundColor Red
            exit
        }
    }
}

function Install-GCloudSDK {
    Write-Host "Installing Google Cloud SDK..."
    $installerUrl = "https://dl.google.com/dl/cloudsdk/channels/rapid/GoogleCloudSDKInstaller.exe"
    $tempPath = [System.IO.Path]::GetTempFileName() + ".exe"
    Write-Host "Downloading Google Cloud SDK installer..."
    Invoke-WebRequest -Uri $installerUrl -OutFile $tempPath
    Write-Host "Download complete. Beginning installation..."
    $installer = Start-Process -FilePath $tempPath -ArgumentList "/S" -PassThru
    $progressBar = 0
    $step = 1
    while (-not $installer.HasExited) {
        Write-Host -NoNewline "`rInstalling... $progressBar%"
        Start-Sleep -Milliseconds 500
        $progressBar += $step
        if ($progressBar -ge 99) {
            $progressBar = 99
        }
    }
    Write-Host "`rInstalling... 100%"
    if ($installer.ExitCode -eq 0) {
        Write-Host "Google Cloud SDK installed successfully."
    } else {
        Write-Host "Installation failed with exit code $($installer.ExitCode)." -ForegroundColor Red
    }
    Remove-Item $tempPath
    Write-Host "Please restart the terminal and run the script again to apply the changes."
    Read-Host "Press any key to close..."
    exit
}

function Login-GCloud {
    try {
        Write-Host "Authenticating with Google Cloud..."
        gcloud auth login
        Write-Host "Login successful."
    } catch {
        Write-Host "Authentication failed. Please check your Google Cloud credentials." -ForegroundColor Red
        exit
    }
}

function ListExistingProjects {
    Write-Host "Retrieving existing Google Cloud projects..."
    try {
        $projectsJsonArray = gcloud projects list --format="json"
        $projectsJson = $projectsJsonArray -join "`n"
        if (-not $projectsJson -or $projectsJson -eq '[]') {
            Write-Host "No projects found."
            return @()
        }
        $projectsList = $projectsJson | ConvertFrom-Json
        if ($projectsList -isnot [System.Array]) {
            $projectsList = @($projectsList)
        }
        if ($projectsList.Count -eq 0) {
            Write-Host "No projects found."
            return @()
        }
        Write-Host "`nActive Google Cloud Projects:"
        $index = 1
        foreach ($project in $projectsList) {
            Write-Host "$index. Project ID: $($project.projectId), Name: $($project.name)"
            $index++
        }
        return $projectsList
    } catch {
        Write-Host "Failed to retrieve projects: $_" -ForegroundColor Red
        exit
    }
}

function Get-YesNoInput {
    param (
        [string]$promptMessage
    )
    $input = Read-Host $promptMessage
    while ($input -ne 'Y' -and $input -ne 'N') {
        Write-Host "Invalid input. Please enter 'Y' for Yes or 'N' for No."
        $input = Read-Host $promptMessage
    }
    return $input
}

function Get-ValidProjectNumber {
    param (
        [int]$maxProjectCount
    )
    if ($maxProjectCount -eq 0) {
        Write-Host "No projects available for selection."
        exit
    }
    $selectedProjectNumber = Read-Host "Enter the number of the project you want to use"
    while (-not ($selectedProjectNumber -as [int]) -or ($selectedProjectNumber -lt 1 -or $selectedProjectNumber -gt $maxProjectCount)) {
        Write-Host "Invalid selection. Please enter a valid project number (1 to $maxProjectCount):"
        $selectedProjectNumber = Read-Host
    }
    return $selectedProjectNumber
}

function SelectOrCreateProject {
    $projects = ListExistingProjects
    if ($projects -isnot [System.Array]) {
        $projects = @($projects)
    }
    $projectsCount = $projects.Count
    if ($projectsCount -gt 0) {
        if ($projectsCount -eq 1) {
            $singleProject = $projects[0]
            $singleProjectId = $singleProject.projectId
            $useSingleProjectResponse = Get-YesNoInput "There is only one project [$singleProjectId]. Would you like to use it? (Y/N)"
            if ($useSingleProjectResponse -eq 'Y') {
                Write-Host "Setting project [$singleProjectId] as the active project..."
                gcloud config set project $singleProjectId | Out-Null
                Write-Host "Project set to: $singleProjectId"
                return $singleProjectId
            } else {
                Write-Host "User chose not to use the single project."
                $createNewProjectResponse = Get-YesNoInput "Would you like to create a new project? (Y/N)"
                if ($createNewProjectResponse -eq 'Y') {
                    return Create-GCloudProject
                } else {
                    Write-Host "No project selected. Exiting..." -ForegroundColor Red
                    exit
                }
            }
        } else {
            $useExistingResponse = Get-YesNoInput "Would you like to use an existing project? (Y/N)"
            if ($useExistingResponse -eq 'Y') {
                $selectedProjectNumber = Get-ValidProjectNumber -maxProjectCount $projectsCount
                $selectedProject = $projects[$selectedProjectNumber - 1]
                $selectedProjectId = $selectedProject.projectId
                Write-Host "Setting project [$selectedProjectId] as the active project..."
                gcloud config set project $selectedProjectId | Out-Null
                Write-Host "Project set to: $selectedProjectId"
                return $selectedProjectId
            } else {
                $createNewProjectResponse = Get-YesNoInput "Would you like to create a new project? (Y/N)"
                if ($createNewProjectResponse -eq 'Y') {
                    return Create-GCloudProject
                } else {
                    Write-Host "No project selected. Exiting..." -ForegroundColor Red
                    exit
                }
            }
        }
    } else {
        Write-Host "No existing projects available."
        return Create-GCloudProject
    }
}

function Create-GCloudProject {
    try {
        $randomNumber = -join ((0..9) | Get-Random -Count 11)
        $projectID = "voltaic-$randomNumber"
        Write-Host "Creating new Google Cloud project with ID: $projectID"
        gcloud projects create $projectID --name=$projectID | Out-Null
        Write-Host "Project created successfully."
        gcloud config set project $projectID | Out-Null
        Write-Host "Project set to: $projectID"
        return $projectID
    } catch {
        Write-Host "Failed to create Google Cloud project: $_" -ForegroundColor Red
        exit
    }
}

function Enable-GSheetsAPI {
    try {
        Write-Host "Enabling Google Sheets API..."
        gcloud services enable sheets.googleapis.com | Out-Null
        Write-Host "Google Sheets API enabled."
    } catch {
        Write-Host "Failed to enable Google Sheets API: $_" -ForegroundColor Red
        exit
    }
}

function Create-ServiceAccountOAuthCredentials {
    try {
        $serviceAccountName = Read-Host "Enter the Service Account Name (6-30 characters, letters, digits, and hyphens only)"
        while ($serviceAccountName.Length -lt 6 -or $serviceAccountName.Length -gt 30 -or $serviceAccountName -notmatch "^[a-zA-Z][a-zA-Z\d\-]*[a-zA-Z\d]$") {
            $serviceAccountName = Read-Host "Invalid Service Account Name. Please enter again:"
        }
        $projectID = gcloud config get-value project
        $serviceAccountEmail = "$serviceAccountName@$projectID.iam.gserviceaccount.com"
        $serviceAccountExists = gcloud iam service-accounts list --filter="email:$serviceAccountEmail" --format="value(email)"
        if ($serviceAccountExists) {
            Write-Host "Service account $serviceAccountName already exists. Fetching credentials..."
        } else {
            Write-Host "Creating a new service account for OAuth credentials."
            gcloud iam service-accounts create $serviceAccountName --display-name="$serviceAccountName" | Out-Null
            Write-Host "Service account created."
            gcloud projects add-iam-policy-binding $projectID --member="serviceAccount:$serviceAccountEmail" --role="roles/editor" | Out-Null
            Write-Host "Roles assigned."
        }
        gcloud iam service-accounts keys create credentials.json --iam-account=$serviceAccountEmail | Out-Null
        Write-Host "Credentials saved as credentials.json."
        return $serviceAccountEmail
    } catch {
        Write-Host "Failed to create or fetch OAuth credentials: $_" -ForegroundColor Red
        exit
    }
}

function Confirm-BeforeExit {
    Write-Host "`nProcess complete. Press any key to close."
    [System.Console]::ReadKey() > $null
}

if (-not (Check-GCloudInstallation)) {
    exit
}

try {
    Login-GCloud
    $projectID = SelectOrCreateProject
    Enable-GSheetsAPI
    $serviceAccountEmail = Create-ServiceAccountOAuthCredentials

    Write-Host "`nPlease share your Google Spreadsheet with the following service account email address to grant it access:"
    Write-Host "$serviceAccountEmail" -ForegroundColor Yellow
    Write-Host "You can share the spreadsheet by opening it in Google Sheets, clicking the 'Share' button, and adding the service account email as an editor."

} catch {
    Write-Host "An error occurred during execution: $_" -ForegroundColor Red
} finally {
    Confirm-BeforeExit
}
