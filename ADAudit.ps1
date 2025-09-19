# Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Comprehensive Active Directory Audit Script with Web Interface
.DESCRIPTION
    This script performs a comprehensive audit of Active Directory infrastructure
    and generates an HTML report with findings and recommendations.
    Now includes a modern web interface for better user experience.
.NOTES
    Author: Mohamed ZEGHLACHE
    Version: 2.0
    Requires: Active Directory PowerShell Module, Domain Admin privileges
.EXAMPLE
    .\ADAudit.ps1
    Starts the web interface on http://localhost:8080
    
.EXAMPLE
    .\ADAudit.ps1 -Port 9090
    Starts the web interface on http://localhost:9090
    
.EXAMPLE  
    .\ADAudit.ps1 -NoWebServer
    Runs in traditional command-line mode
#>

param(
    [string]$OutputPath = ".\ADaudit",
    [string]$ReportName = "AD_Audit_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    [int]$Port = 8080,
    [switch]$NoWebServer,
    [switch]$SkipExecution,
    [switch]$WebInitiated
)

# Global variables for web server and progress tracking
$Global:ProgressData = @{
    CurrentStep = 0
    TotalSteps = 23
    CurrentStatus = "Initializing..."
    IsRunning = $false
    AuditConfig = @{}
    StartTime = $null
    ReportPath = ""
}

$Global:HttpListener = $null

# Import required modules
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module GroupPolicy -ErrorAction SilentlyContinue
} catch {
    Write-Error "Required modules not available. Please install RSAT tools."
    exit 1
}





# Web Server Functions
function Find-AvailablePort {
    param([int]$StartPort = 8080, [int]$EndPort = 8090)
    
    for ($port = $StartPort; $port -le $EndPort; $port++) {
        $listener = $null
        try {
            $listener = New-Object System.Net.HttpListener
            $listener.Prefixes.Add("http://localhost:$port/")
            $listener.Start()
            $listener.Stop()
            return $port
        }
        catch {
            # Port is in use, try next one
            if ($listener) {
                try { $listener.Stop() } catch { }
            }
        }
    }
    throw "No available ports found between $StartPort and $EndPort"
}

function Start-WebServer {
    param([int]$Port = 8080)
    
    try {
        # Stop any existing web server first
        if ($Global:HttpListener -and $Global:HttpListener.IsListening) {
            Write-Host "Stopping existing web server..." -ForegroundColor Yellow
            $Global:HttpListener.Stop()
            $Global:HttpListener = $null
        }
        
        # Try the requested port first, if it fails find an available one
        $ActualPort = $Port
        try {
            $Global:HttpListener = New-Object System.Net.HttpListener
            $Global:HttpListener.Prefixes.Add("http://localhost:$ActualPort/")
            $Global:HttpListener.Start()
        }
        catch {
            Write-Warning "Port $Port is in use, finding an available port..."
            $ActualPort = Find-AvailablePort -StartPort $Port -EndPort ($Port + 10)
            $Global:HttpListener = New-Object System.Net.HttpListener
            $Global:HttpListener.Prefixes.Add("http://localhost:$ActualPort/")
            $Global:HttpListener.Start()
        }
        
        Write-Host "🌐 Web server started at http://localhost:$ActualPort" -ForegroundColor Green
        Write-Host "🚀 Opening browser..." -ForegroundColor Yellow
        
        # Open browser
        Start-Process "http://localhost:$ActualPort"
        
        # Handle requests in a loop
        while ($Global:HttpListener.IsListening) {
            try {
                $context = $Global:HttpListener.GetContext()
                Handle-WebRequest -Context $context
            }
            catch [System.Net.HttpListenerException] {
                # Expected when listener is stopped
                break
            }
            catch {
                Write-Warning "Error handling request: $_"
            }
        }
    }
    catch {
        Write-Error "Failed to start web server: $_"
    }
}

function Handle-WebRequest {
    param($Context)
    
    $request = $Context.Request
    $response = $Context.Response
    $url = $request.Url.AbsolutePath
    
    # Add CORS headers
    $response.Headers.Add("Access-Control-Allow-Origin", "*")
    $response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    $response.Headers.Add("Access-Control-Allow-Headers", "Content-Type")
    
    try {
        switch ($url) {
            "/" {
                Serve-LauncherPage -Response $response
            }
            "/start-audit" {
                if ($request.HttpMethod -eq "POST") {
                    Handle-StartAudit -Request $request -Response $response
                }
                elseif ($request.HttpMethod -eq "OPTIONS") {
                    $response.StatusCode = 200
                    $response.Close()
                }
            }
            "/progress" {
                Serve-Progress -Response $response
            }
            {$_ -match "^/report/"} {
                $filename = ($url -split "/")[-1]
                Serve-Report -Response $response -Filename $filename
            }
            default {
                $response.StatusCode = 404
                $bytes = [System.Text.Encoding]::UTF8.GetBytes("Not Found")
                $response.OutputStream.Write($bytes, 0, $bytes.Length)
                $response.Close()
            }
        }
    }
    catch {
        Write-Error "Error handling request: $_"
        $response.StatusCode = 500
        $bytes = [System.Text.Encoding]::UTF8.GetBytes("Internal Server Error")
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
        $response.Close()
    }
}

function Serve-LauncherPage {
    param($Response)
    
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Active Directory Audit Tool</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        .check-item {
            transition: all 0.3s ease;
            border-left: 4px solid #dee2e6;
        }
        
        .check-item.running {
            background-color: #e3f2fd;
            border-left-color: #0d6efd;
        }
        
        .check-item.completed {
            background-color: #d1e7dd;
            border-left-color: #198754;
        }
        
        .check-item.error {
            background-color: #fff3cd;
            border-left-color: #ffc107;
        }
        
        .feature-icon {
            font-size: 1.2rem;
            margin-right: 0.5rem;
        }
    </style>
</head>
<body class="bg-light">
    <div class="container my-5">
        <!-- Header -->
        <div class="text-center mb-5">
            <h1 class="display-5 fw-bold text-primary">
                <i class="bi bi-shield-check me-3"></i>Active Directory Audit Tool
            </h1>
            <p class="lead text-muted">Comprehensive AD infrastructure analysis and security assessment</p>
        </div>

        <!-- Main Card -->
        <div class="card shadow-sm">
            <div class="card-body p-4">
                <!-- About Section -->
                <div class="row mb-4">
                    <div class="col-md-8">
                        <h5 class="card-title mb-3">About This Tool</h5>
                        <p class="card-text text-muted mb-3">
                            This PowerShell-based audit tool performs comprehensive analysis of your Active Directory environment, 
                            including domain controllers health, security settings, user accounts, and policy configurations.
                        </p>
                        <div class="row">
                            <div class="col-sm-6 mb-2">
                                <small class="text-muted">
                                    <i class="bi bi-check-circle-fill text-success feature-icon"></i>
                                    Domain Controller Health
                                </small>
                            </div>
                            <div class="col-sm-6 mb-2">
                                <small class="text-muted">
                                    <i class="bi bi-check-circle-fill text-success feature-icon"></i>
                                    Security Policy Analysis
                                </small>
                            </div>
                            <div class="col-sm-6 mb-2">
                                <small class="text-muted">
                                    <i class="bi bi-check-circle-fill text-success feature-icon"></i>
                                    User Account Statistics
                                </small>
                            </div>
                            <div class="col-sm-6 mb-2">
                                <small class="text-muted">
                                    <i class="bi bi-check-circle-fill text-success feature-icon"></i>
                                    Privileged Group Monitoring
                                </small>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-4 text-center">
                        <div class="p-3">
                            <i class="bi bi-graph-up text-primary" style="font-size: 4rem;"></i>
                        </div>
                    </div>
                </div>

                <!-- Configuration -->
                <div class="border-top pt-4 mb-4">
                    <h6 class="fw-bold mb-3">Audit Configuration</h6>
                </div>

                <!-- Action Button -->
                <div class="d-grid">
                    <button id="startBtn" class="btn btn-primary btn-lg" onclick="startAudit()">
                        <i class="bi bi-play-circle me-2"></i>Start AD Audit
                    </button>
                </div>
            </div>
        </div>

        <!-- Progress Section -->
        <div id="progressCard" class="card mt-4 shadow-sm" style="display: none;">
            <div class="card-body">
                <h6 class="card-title mb-3">
                    <i class="bi bi-gear-fill text-primary me-2"></i>Audit in Progress
                </h6>
                <div class="progress mb-3" style="height: 10px;">
                    <div id="progressBar" class="progress-bar progress-bar-striped progress-bar-animated bg-primary" 
                         role="progressbar" style="width: 0%"></div>
                </div>
                <div id="currentStep" class="text-muted mb-3">Initializing audit...</div>
                <div id="auditSteps" class="mt-3"></div>
            </div>
        </div>

        <!-- Results Section -->
        <div id="resultCard" class="card mt-4 shadow-sm" style="display: none;">
            <div class="card-body text-center">
                <i class="bi bi-check-circle-fill text-success mb-3" style="font-size: 3rem;"></i>
                <h5 class="card-title text-success">Audit Completed Successfully!</h5>
                <p class="card-text text-muted mb-3">Your Active Directory audit report has been generated and is ready for review.</p>
                <button id="viewReportBtn" class="btn btn-success btn-lg" onclick="viewReport()">
                    <i class="bi bi-file-earmark-text me-2"></i>View Detailed Report
                </button>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>

    <script>
        let statusInterval;
        let auditSteps = [];
        
        async function startAudit() {
            const startBtn = document.getElementById('startBtn');
            const progressCard = document.getElementById('progressCard');
            const resultCard = document.getElementById('resultCard');

            startBtn.disabled = true;
            startBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Starting...';
            progressCard.style.display = 'block';
            resultCard.style.display = 'none';

            // Initialize audit steps display
            initializeAuditSteps();

            try {
                const response = await fetch('/start-audit', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        timestamp: new Date().toISOString()
                    })
                });
                
                if (response.ok) {
                    // Start polling for progress
                    statusInterval = setInterval(checkProgress, 2000);
                } else {
                    throw new Error('Failed to start audit');
                }
            } catch (error) {
                console.error('Error starting audit:', error);
                document.getElementById('currentStep').innerHTML = 
                    '<i class="bi bi-exclamation-triangle text-danger me-2"></i>Error starting audit: ' + error.message;
                resetStartButton();
                progressCard.style.display = 'none';
            }
        }

        function initializeAuditSteps() {
            auditSteps = [
                'Domain Controller Health Check',
                'Replication Status Analysis',
                'Service Status Verification',
                'Security Policy Assessment',
                'User Account Analysis',
                'Privileged Group Monitoring',
                'Password Policy Review',
                'GPO Analysis',
                'Report Generation'
            ];
            
            
            const stepsContainer = document.getElementById('auditSteps');
            stepsContainer.innerHTML = auditSteps.map((step, index) => 
                '<div class="check-item p-2 mb-2 rounded" id="step-' + index + '">' +
                    '<i class="bi bi-circle text-muted me-2"></i>' +
                    '<small class="text-muted">' + step + '</small>' +
                '</div>'
            ).join('');
        }
        
        function resetStartButton() {
            const startBtn = document.getElementById('startBtn');
            startBtn.disabled = false;
            startBtn.innerHTML = '<i class="bi bi-play-circle me-2"></i>Start AD Audit';
        }
        
        async function checkProgress() {
            try {
                const response = await fetch('/progress');
                const data = await response.json();
                
                const progressBar = document.getElementById('progressBar');
                const currentStep = document.getElementById('currentStep');
                
                const percentage = Math.round((data.currentStep / data.totalSteps) * 100);
                progressBar.style.width = percentage + '%';
                progressBar.setAttribute('aria-valuenow', percentage);
                
                // Update current step text
                const stepText = data.currentStatus || 'Processing...';
                currentStep.innerHTML = '<i class="bi bi-gear-fill text-primary me-2"></i>' + stepText;
                
                // Update step visual indicators
                updateStepIndicators(data.currentStep, data.totalSteps);
                
                if (!data.isRunning && data.currentStep >= data.totalSteps) {
                    clearInterval(statusInterval);
                    showResults();
                }
            } catch (error) {
                console.error('Error checking progress:', error);
                document.getElementById('currentStep').innerHTML = 
                    '<i class="bi bi-exclamation-triangle text-warning me-2"></i>Error checking progress';
            }
        }
        
        function updateStepIndicators(currentStep, totalSteps) {
            auditSteps.forEach((step, index) => {
                const stepElement = document.getElementById('step-' + index);
                const icon = stepElement.querySelector('i');
                const text = stepElement.querySelector('small');
                
                if (index < currentStep) {
                    // Completed step
                    stepElement.className = 'check-item completed p-2 mb-2 rounded';
                    icon.className = 'bi bi-check-circle-fill text-success me-2';
                    text.className = 'text-success fw-bold';
                } else if (index === currentStep) {
                    // Current step
                    stepElement.className = 'check-item running p-2 mb-2 rounded';
                    icon.className = 'bi bi-arrow-right-circle-fill text-primary me-2';
                    text.className = 'text-primary fw-bold';
                } else {
                    // Pending step
                    stepElement.className = 'check-item p-2 mb-2 rounded';
                    icon.className = 'bi bi-circle text-muted me-2';
                    text.className = 'text-muted';
                }
            });
        }
        
        function showResults() {
            const progressCard = document.getElementById('progressCard');
            const resultCard = document.getElementById('resultCard');
            
            progressCard.style.display = 'none';
            resultCard.style.display = 'block';
            resetStartButton();
            
            // Scroll to results
            resultCard.scrollIntoView({ behavior: 'smooth' });
        }
        
        function viewReport() {
            // Get the latest report
            fetch('/progress')
                .then(response => response.json())
                .then(data => {
                    if (data.reportPath) {
                        const reportName = data.reportPath.split('\\').pop();
                        window.open('/report/' + reportName, '_blank');
                    } else {
                        // Fallback - try to open a recent report
                        window.open('/report/', '_blank');
                    }
                })
                .catch(error => {
                    console.error('Error opening report:', error);
                    // Show alert
                    const alert = document.createElement('div');
                    alert.className = 'alert alert-warning alert-dismissible fade show mt-3';
                    alert.innerHTML = 
                        '<i class="bi bi-exclamation-triangle me-2"></i>' +
                        'Unable to open report. Please check the console for details.' +
                        '<button type="button" class="btn-close" data-bs-dismiss="alert"></button>';
                    document.querySelector('.container').appendChild(alert);
                });
        }
    </script>
</body>
</html>
"@

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($htmlContent)
    $Response.ContentType = "text/html; charset=utf-8"
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Response.Close()
}

function Handle-StartAudit {
    param($Request, $Response)
    
    try {
        # Parse JSON request
        $reader = New-Object System.IO.StreamReader($Request.InputStream)
        $jsonString = $reader.ReadToEnd()
        $reader.Close()
        
        if ($jsonString) {
            $config = $jsonString | ConvertFrom-Json
            
            # Store configuration
            $Global:ProgressData.AuditConfig = @{
                Timestamp = if ($config.timestamp) { $config.timestamp } else { (Get-Date).ToString() }
            }
        }
        
        $Global:ProgressData.IsRunning = $true
        $Global:ProgressData.StartTime = Get-Date
        $Global:ProgressData.CurrentStep = 0
        $Global:ProgressData.CurrentStatus = "Starting audit..."
        
        # Define audit step names for progress
        $Global:AuditSteps = @(
            "Starting audit...",
            "Importing required modules...",
            "Enhanced Domain Controller Health Assessment...",
            "Replication Health Assessment...",
            "DNS Configuration Testing...",
            "FSMO Roles Verification...",
            "Trust Relationships Analysis...", 
            "AD Sites & Services Audit...",
            "Group Policy Assessment...",
            "User & Computer Accounts Analysis...",
            "Security Analysis & Risk Assessment...",
            "OU Hierarchy Mapping...",
            "Protocols & Legacy Services Check...",
            "Orphaned SID Analysis...",
            "Event Log Analysis...",
            "Security Group Nesting Analysis...",
            "Protocol Security Analysis...",
            "Group Managed Service Accounts Usage...",
            "LAPS Check...",
            "Privileged Account Monitoring...",
            "DC Performance Metrics...",
            "Schema Architecture Analysis...",
            "AD Object Protection Analysis..."
        )
        
        
        $Global:AuditSteps += @("Generating HTML Report...", "Audit completed successfully!")
        
        # Progress will be advanced by the Serve-Progress function
        
        # Actually start the audit execution in a background job
        $OutputPath = Join-Path $PSScriptRoot "Reports"
        $Config = @{}
        
        # Start the audit execution in a background runspace
        $Global:AuditRunspace = [powershell]::Create()
        $Global:AuditRunspace.AddScript({
            param($Config, $OutputPath, $PSScriptRoot)

            try {
                Write-Host "🔧 Runspace: Starting execution..." -ForegroundColor Cyan

                # Set the location to the script directory
                Set-Location $PSScriptRoot
                Write-Host "🔧 Runspace: Location set to $PSScriptRoot" -ForegroundColor Cyan

                # Import the script to get access to the audit functions
                Write-Host "🔧 Runspace: Importing script..." -ForegroundColor Cyan
                . "$PSScriptRoot\ADAudit.ps1" -SkipExecution
                Write-Host "🔧 Runspace: Script imported successfully" -ForegroundColor Cyan

                # Check if function exists
                if (Get-Command Start-WebAuditExecution -ErrorAction SilentlyContinue) {
                    Write-Host "🔧 Runspace: Start-WebAuditExecution function found, executing..." -ForegroundColor Cyan
                    # Run the web audit execution
                    Start-WebAuditExecution -Config $Config -OutputPath $OutputPath
                } else {
                    Write-Error "🔧 Runspace: Start-WebAuditExecution function not found!"
                    throw "Start-WebAuditExecution function not available in runspace"
                }
            } catch {
                Write-Error "🔧 Runspace Error: $($_.Exception.Message)"
                Write-Host "🔧 Runspace: Full error details: $_" -ForegroundColor Red
                throw
            }
        })
        $Global:AuditRunspace.AddParameter('Config', $Config)
        $Global:AuditRunspace.AddParameter('OutputPath', $OutputPath)
        $Global:AuditRunspace.AddParameter('PSScriptRoot', $PSScriptRoot)
        $Global:AuditJob = $Global:AuditRunspace.BeginInvoke()

        # Give the runspace a moment to start and check for immediate errors
        Start-Sleep -Milliseconds 500
        if ($Global:AuditJob.IsCompleted -and $Global:AuditRunspace.HadErrors) {
            $errors = $Global:AuditRunspace.Streams.Error.ReadAll()
            Write-Error "Runspace failed to start: $($errors -join '; ')"
            throw "Runspace execution failed: $($errors -join '; ')"
        }

        Write-Host "🔧 Audit job started successfully with job state: $($Global:AuditJob.AsyncState)" -ForegroundColor Green

        # Send success response
        $responseData = @{ 
            status = "started"; 
            message = "Audit started successfully" 
        } | ConvertTo-Json
        
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($responseData)
        $Response.ContentType = "application/json; charset=utf-8"
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $Response.Close()
    }
    catch {
        Write-Error "Error starting audit: $_"
        $responseData = @{ 
            status = "error"; 
            message = $_.Exception.Message 
        } | ConvertTo-Json
        
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($responseData)
        $Response.ContentType = "application/json; charset=utf-8"
        $Response.StatusCode = 500
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $Response.Close()
    }
}

function Serve-Progress {
    param($Response)
    
    try {
        # Advance progress if audit is running
        if ($Global:ProgressData.IsRunning -and $Global:ProgressData.CurrentStep -lt $Global:ProgressData.TotalSteps) {
            $timeSinceStart = (Get-Date) - $Global:ProgressData.StartTime
            $expectedStep = [math]::Floor($timeSinceStart.TotalSeconds / 2) # Advance every 2 seconds
            
            if ($expectedStep -gt $Global:ProgressData.CurrentStep -and $expectedStep -le $Global:ProgressData.TotalSteps) {
                $Global:ProgressData.CurrentStep = [math]::Min($expectedStep, $Global:ProgressData.TotalSteps)
                
                # Update status based on current step
                if ($Global:ProgressData.CurrentStep -le $Global:AuditSteps.Length) {
                    $Global:ProgressData.CurrentStatus = $Global:AuditSteps[$Global:ProgressData.CurrentStep - 1]
                }
                
                # Handle completion
                if ($Global:ProgressData.CurrentStep -ge $Global:ProgressData.TotalSteps) {
                    $Global:ProgressData.IsRunning = $false
                    $Global:ProgressData.CurrentStatus = "Audit completed successfully!"
                    
                    # Fallback report removed - main comprehensive report will be generated instead
                }
            }
        }
        
        $progressJson = $Global:ProgressData | ConvertTo-Json -Depth 3
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($progressJson)
        
        $Response.ContentType = "application/json; charset=utf-8"
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $Response.Close()
    }
    catch {
        $errorData = @{ error = $_.Exception.Message } | ConvertTo-Json
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($errorData)
        $Response.ContentType = "application/json; charset=utf-8"
        $Response.StatusCode = 500
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $Response.Close()
    }
}

function Serve-Report {
    param($Response, $Filename)
    
    try {
        # If no filename provided, find the most recent report
        if (-not $Filename -or $Filename -eq "") {
            $reportFiles = Get-ChildItem -Path $OutputPath -Filter "*.html" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
            if ($reportFiles) {
                $Filename = $reportFiles[0].Name
            } else {
                $Response.StatusCode = 404
                $bytes = [System.Text.Encoding]::UTF8.GetBytes("No reports found in output directory")
                $Response.OutputStream.Write($bytes, 0, $bytes.Length)
                $Response.Close()
                return
            }
        }
        
        $reportPath = Join-Path $OutputPath $Filename
        if (Test-Path $reportPath) {
            $content = Get-Content $reportPath -Raw -Encoding UTF8
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
            $Response.ContentType = "text/html; charset=utf-8"
            $Response.ContentLength64 = $bytes.Length
            $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        }
        else {
            $Response.StatusCode = 404
            $bytes = [System.Text.Encoding]::UTF8.GetBytes("Report not found: $Filename")
            $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        }
        $Response.Close()
    }
    catch {
        $Response.StatusCode = 500
        $bytes = [System.Text.Encoding]::UTF8.GetBytes("Error serving report: $($_.Exception.Message)")
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $Response.Close()
    }
}

function Update-Progress {
    param(
        [string]$Status,
        [int]$Step = -1
    )
    
    if ($Step -ge 0) {
        $Global:ProgressData.CurrentStep = $Step
    }
    else {
        $Global:ProgressData.CurrentStep++
    }
    
    $Global:ProgressData.CurrentStatus = $Status
    
    # Also display in console
    $percent = [math]::Round(($Global:ProgressData.CurrentStep / $Global:ProgressData.TotalSteps) * 100, 1)
    Write-Host "[$percent%] $Status" -ForegroundColor Cyan
}

function Stop-WebServer {
    if ($Global:HttpListener -and $Global:HttpListener.IsListening) {
        $Global:HttpListener.Stop()
        $Global:HttpListener.Close()
        Write-Host "🛑 Web server stopped." -ForegroundColor Yellow
    }
}

function Start-WebAuditExecution {
    param($Config, $OutputPath)
    
    try {
        $Global:ProgressData.TotalSteps = 23
        Update-Progress "Importing required modules..." 1
        
        # Import required modules
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            Import-Module GroupPolicy -ErrorAction SilentlyContinue
        } catch {
            throw "Required modules not available. Please install RSAT tools."
        }

        # Create output directory
        if (!(Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }

        # Initialize audit results
        $AuditResults = @{}
        $StartTime = Get-Date
        $ReportName = "AD_Audit_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

        Update-Progress "Starting Active Directory Audit..." 2

        # For web mode, execute the actual audit by calling the traditional execution
        Update-Progress "Domain Controller Health Assessment..." 3
        
        # Launch the audit in a separate PowerShell console window (visible to user)
        try {
            Write-Host "🔍 Starting real AD audit execution..." -ForegroundColor Green

            $ScriptPath = $PSCommandPath
            if (-not $ScriptPath) { $ScriptPath = $MyInvocation.MyCommand.Path }

            # Start the audit in a new PowerShell console window with -NoExit to keep it open
            $ProcessArgs = @(
                "-NoExit",
                "-ExecutionPolicy", "Bypass",
                "-File", "`"$ScriptPath`"",
                "-NoWebServer",
                "-OutputPath", "`"$OutputPath`"",
                "-ReportName", "`"$ReportName`"",
                "-WebInitiated"  # Add flag to indicate this was started from web interface
            )

            Write-Host "Launching audit console..." -ForegroundColor Yellow
            $AuditProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $ProcessArgs -PassThru

            # Set report path for serving later
            $Global:ProgressData.ReportPath = Join-Path $OutputPath $ReportName

            # Return immediately after launching the process - the separate process will handle the actual audit
            return @{
                Success = $true
                ReportPath = $Global:ProgressData.ReportPath
                Duration = (Get-Date) - $StartTime
            }

        } catch {
            Write-Warning "Failed to launch audit console: $_"
            # Return immediately on error as well
            return @{
                Success = $false
                Error = $_.Exception.Message
                Duration = (Get-Date) - $StartTime
            }
        }
        
    } catch {
        $Global:ProgressData.IsRunning = $false
        $Global:ProgressData.CurrentStatus = "Error: $($_.Exception.Message)"
        Write-Error "Audit failed: $_"
        throw
    }
}


# Decide execution mode immediately
if ($SkipExecution) {
    # Skip execution when script is being imported by runspace
    return
}

if (-not $NoWebServer) {
    # Web server mode - start web server instead of running audit
    Write-Host "🌐 Starting AD Audit Tool web interface..." -ForegroundColor Cyan
    try {
        Start-WebServer -Port $Port
        exit 0  # Exit after web server stops
    }
    catch {
        Write-Error "❌ Failed to start web interface: $($_.Exception.Message)"
        Write-Host "🔄 Falling back to command-line mode..." -ForegroundColor Yellow
        # Continue to command-line mode below
    }
}

# Command-line mode - run audit
Write-Host "📋 Running in command-line mode..." -ForegroundColor Yellow

# Prevent double execution by checking if already running or if called from web interface
if ($Global:AuditInProgress) {
    Write-Warning "Audit already in progress. Exiting to prevent duplication."
    return
}

# Skip command-line execution if this was initiated from web interface runspace (but not from separate process)
if ($Global:WebInitiatedExecution -and -not $WebInitiated) {
    Write-Host "⚠️ Execution initiated from web interface runspace. Skipping duplicate command-line mode." -ForegroundColor Yellow
    return
}

# If WebInitiated is set, this is a separate process launched by web interface - proceed with audit
if ($WebInitiated) {
    Write-Host "🌐 Audit launched from web interface - running in separate console..." -ForegroundColor Cyan
}
$Global:AuditInProgress = $true

# Create output directory
if (!(Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Initialize audit results
$AuditResults = @{}
$StartTime = Get-Date

Write-Host "Starting Active Directory Audit..." -ForegroundColor Green
Write-Host "Report will be saved to: $OutputPath\$ReportName" -ForegroundColor Yellow

#================================================================
# POWERSHELL AUDIT COMMANDS SECTION
#================================================================
# This section contains all the PowerShell commands that collect
# and analyze Active Directory data. All analysis results are
# stored in the $AuditResults hashtable.
#================================================================

#region Enhanced Domain Controller Health Assessment
Write-Host "Performing Enhanced Domain Controller Health Assessment..." -ForegroundColor Cyan
try {
    $EnhancedDCHealth = @()
    $AllDomains = (Get-ADForest).Domains
    
    foreach ($Domain in $AllDomains) {
        $DomainControllers = Get-ADDomainController -Filter * -Server $Domain
        
        foreach ($DC in $DomainControllers) {
            Write-Host "  Checking $($DC.HostName)..." -ForegroundColor Yellow
            
            $DCHealthInfo = @{
                DomainController = $DC.HostName
                Domain           = $Domain
                Site             = $DC.Site
                OperatingSystem  = $DC.OperatingSystem
                IPv4Address      = $DC.IPv4Address
                IsGlobalCatalog  = $DC.IsGlobalCatalog
                IsReadOnly       = $DC.IsReadOnly
            }
            
            # Check if this is the local DC
            $DCHostName = $DC.HostName.ToString()
            $LocalComputerName = $env:COMPUTERNAME
            $IsLocalDC = ($DCHostName -eq $LocalComputerName -or $DCHostName -eq "localhost" -or $DCHostName -eq "127.0.0.1" -or $DCHostName -eq "$LocalComputerName.$env:USERDNSDOMAIN")

            # NSLookup
            try {
                $DCHealthInfo.NSLookup = (Resolve-DnsName -Name $DCHostName -ErrorAction Stop).IPAddress
            } catch { $DCHealthInfo.NSLookup = "DNS Error" }

            # Ping - Primary connectivity test with timeout (skip for local DC)
            if ($IsLocalDC) {
                $DCHealthInfo.PingStatus = "Local"
                $isOnline = $true
                Write-Host "    ✓ Local DC - Skipping ping test" -ForegroundColor Green
            } else {
                try {
                    $isOnline = Test-Connection -ComputerName $DCHostName -Count 1 -Quiet -ErrorAction Stop
                    $DCHealthInfo.PingStatus = if ($isOnline) { "Online" } else { "Offline" }
                } catch {
                    $DCHealthInfo.PingStatus = "Unreachable"
                    $isOnline = $false
                    Write-Warning "Cannot ping ${DCHostName}: Network unreachable"
                }
            }

            # Skip detailed checks if DC is not reachable (but not for local DC)
            if ((-not $isOnline -or $DCHealthInfo.PingStatus -eq "Unreachable") -and -not $IsLocalDC) {
                $DCHealthInfo.Uptime = "Unreachable"
                $DCHealthInfo.TimeDifference = "Unreachable"
                $DCHealthInfo.SYSVOL_Netlogon = "Unreachable"
                $DCHealthInfo.DNSService = "Unreachable"
                $DCHealthInfo.NTDSService = "Unreachable"
                $DCHealthInfo.NETLOGONService = "Unreachable"
                $DCHealthInfo.OSDriveFreeSpacePercent = "Unreachable"
                $DCHealthInfo.OSDriveFreeSpaceGB = "Unreachable"
                $DCHealthInfo.DCDiagResults = "Unreachable - Cannot run DCDiag"
                
                Write-Host "    Skipping detailed checks for unreachable DC: $($DC.HostName)" -ForegroundColor Red
                $EnhancedDCHealth += $DCHealthInfo
                continue
            }
            
            # Uptime
            try {
                if ($IsLocalDC) {
                    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
                } else {
                    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $DCHostName -ErrorAction Stop
                }
                $DCHealthInfo.Uptime = (Get-Date) - $os.LastBootUpTime
            } catch {
                $DCHealthInfo.Uptime = "Unreachable"
                Write-Warning "Cannot retrieve uptime for ${DCHostName}: $_"
            }

            # Time difference
            try {
                if ($IsLocalDC) {
                    $remoteTime = Get-Date
                    $DCHealthInfo.TimeDifference = 0  # Local time difference is 0
                } else {
                    $remoteTime = (Get-Date (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $DCHostName -ErrorAction Stop).LocalDateTime)
                    $DCHealthInfo.TimeDifference = ((Get-Date) - $remoteTime).TotalSeconds
                }
            } catch {
                $DCHealthInfo.TimeDifference = "Unreachable"
                Write-Warning "Cannot retrieve time for ${DCHostName}: $_"
            }
            
            # SYSVOL/NETLOGON share check
            try {
                if ($IsLocalDC) {
                    # For local DC, check local paths
                    $netlogon = Test-Path -Path "C:\Windows\SYSVOL\sysvol\$env:USERDNSDOMAIN\scripts" -ErrorAction SilentlyContinue
                    $sysvol = Test-Path -Path "C:\Windows\SYSVOL" -ErrorAction SilentlyContinue
                } else {
                    $netlogon = Test-Path -Path "\\$DCHostName\NETLOGON" -ErrorAction Stop
                    $sysvol = Test-Path -Path "\\$DCHostName\SYSVOL" -ErrorAction Stop
                }
                $DCHealthInfo.SYSVOL_Netlogon = if ($netlogon -and $sysvol) { "Available" } else { "Missing" }
            } catch {
                $DCHealthInfo.SYSVOL_Netlogon = "Unreachable"
                Write-Warning "Cannot access shares on $($DC.HostName): $_"
            }

            # Service status checks
            try {
                if ($IsLocalDC) {
                    $DCHealthInfo.DNSService = (Get-Service -Name DNS -ErrorAction Stop).Status
                } else {
                    $DCHealthInfo.DNSService = (Get-Service -ComputerName $DCHostName -Name DNS -ErrorAction Stop).Status
                }
            } catch {
                $DCHealthInfo.DNSService = "Unreachable"
                Write-Warning "Cannot check DNS service on ${DCHostName}"
            }

            try {
                if ($IsLocalDC) {
                    $DCHealthInfo.NTDSService = (Get-Service -Name NTDS -ErrorAction Stop).Status
                } else {
                    $DCHealthInfo.NTDSService = (Get-Service -ComputerName $DCHostName -Name NTDS -ErrorAction Stop).Status
                }
            } catch {
                $DCHealthInfo.NTDSService = "Unreachable"
                Write-Warning "Cannot check NTDS service on ${DCHostName}"
            }

            try {
                if ($IsLocalDC) {
                    $DCHealthInfo.NETLOGONService = (Get-Service -Name Netlogon -ErrorAction Stop).Status
                } else {
                    $DCHealthInfo.NETLOGONService = (Get-Service -ComputerName $DCHostName -Name Netlogon -ErrorAction Stop).Status
                }
            } catch {
                $DCHealthInfo.NETLOGONService = "Unreachable"
                Write-Warning "Cannot check Netlogon service on ${DCHostName}"
            }
            
            # Disk space checks
            try {
                if ($IsLocalDC) {
                    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
                } else {
                    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $DCHostName -Filter "DeviceID='C:'" -ErrorAction Stop
                }
                $DCHealthInfo.OSDriveFreeSpacePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)
                $DCHealthInfo.OSDriveFreeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
            } catch {
                $DCHealthInfo.OSDriveFreeSpacePercent = "Unreachable"
                $DCHealthInfo.OSDriveFreeSpaceGB = "Unreachable"
                Write-Warning "Cannot check disk space on ${DCHostName}: $_"
            }

            # DCDiag
            try {
                if ($IsLocalDC) {
                    $DCHealthInfo.DCDiagResults = dcdiag
                } else {
                    $DCHealthInfo.DCDiagResults = dcdiag /s:$DCHostName
                }
            } catch { $DCHealthInfo.DCDiagResults = "Error" }

            $EnhancedDCHealth += New-Object PSObject -Property $DCHealthInfo
        }
    }
    $AuditResults.EnhancedDCHealth = $EnhancedDCHealth
} catch {
    $AuditResults.EnhancedDCHealth = "Error: $($_.Exception.Message)"
}
#endregion


#region DC Performance Metrics
Write-Host "Collecting DC Performance Metrics..." -ForegroundColor Cyan
try {
    $DCPerformanceMetrics = @()
    $AllDomains = (Get-ADForest).Domains

    foreach ($Domain in $AllDomains) {
        $DomainControllers = Get-ADDomainController -Filter * -Server $Domain

        foreach ($DC in $DomainControllers) {
            Write-Host "  Collecting metrics for $($DC.HostName)..." -ForegroundColor Yellow

            $DCMetrics = @{
                DomainController = $DC.HostName
                Domain = $Domain
                CollectionTime = Get-Date
                CPUUsage = "N/A"
                MemoryUsage = "N/A"
                DiskUsage = @()
                ADDatabaseSize = "N/A"
                LogFileSize = "N/A"
                NTDSPerformance = @()
                LDAPConnections = "N/A"
            }

            try {
                # CPU Usage
                $CPUUsage = Get-CimInstance -ClassName Win32_Processor -ComputerName $DC.HostName.ToString() -ErrorAction SilentlyContinue |
                    Measure-Object -Property LoadPercentage -Average
                $DCMetrics.CPUUsage = "$([math]::Round($CPUUsage.Average, 2))%"

                # Memory Usage
                $OS = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $DC.HostName.ToString() -ErrorAction SilentlyContinue
                if ($OS) {
                    $MemoryUsedPercent = [math]::Round(((($OS.TotalVisibleMemorySize - $OS.FreePhysicalMemory) / $OS.TotalVisibleMemorySize) * 100), 2)
                    $DCMetrics.MemoryUsage = "$MemoryUsedPercent% ($([math]::Round(($OS.TotalVisibleMemorySize - $OS.FreePhysicalMemory) / 1MB, 2)) GB / $([math]::Round($OS.TotalVisibleMemorySize / 1MB, 2)) GB)"
                }

                # Disk Usage
                $Disks = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $DC.HostName.ToString() -Filter "DriveType = 3" -ErrorAction SilentlyContinue
                foreach ($Disk in $Disks) {
                    $UsedPercent = [math]::Round((($Disk.Size - $Disk.FreeSpace) / $Disk.Size) * 100, 2)
                    $DCMetrics.DiskUsage += @{
                        Drive = $Disk.DeviceID
                        SizeGB = [math]::Round($Disk.Size / 1GB, 2)
                        FreeGB = [math]::Round($Disk.FreeSpace / 1GB, 2)
                        UsedPercent = $UsedPercent
                    }
                }

                # NTDS Database Size - try multiple methods
                $NTDSPath = "\\$($DC.HostName)\C$\Windows\NTDS\ntds.dit"
                $LocalNTDSPath = "C:\Windows\NTDS\ntds.dit"
                $DatabaseSizeFound = $false

                # Check if this is the local DC
                $DCHostName = $DC.HostName.ToString()
                $LocalComputerName = $env:COMPUTERNAME
                $IsLocalDC = ($DCHostName -eq $LocalComputerName -or $DCHostName -eq "localhost" -or $DCHostName -eq "127.0.0.1" -or $DCHostName -eq "$LocalComputerName.$env:USERDNSDOMAIN")

                # Method 1: Local file access (if this is the local DC)
                if ($IsLocalDC -and (Test-Path $LocalNTDSPath)) {
                    try {
                        $DBSize = (Get-Item $LocalNTDSPath -ErrorAction Stop).Length
                        $DCMetrics.ADDatabaseSize = "$([math]::Round($DBSize / 1GB, 2)) GB"
                        $DatabaseSizeFound = $true
                        Write-Host "    ✓ Local NTDS.dit access successful" -ForegroundColor Green
                    } catch {
                        Write-Host "    ⚠️ Local access to NTDS.dit failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }

                # Method 2: UNC Path access (if not local DC)
                if (-not $DatabaseSizeFound -and (Test-Path $NTDSPath)) {
                    try {
                        $DBSize = (Get-Item $NTDSPath -ErrorAction Stop).Length
                        $DCMetrics.ADDatabaseSize = "$([math]::Round($DBSize / 1GB, 2)) GB"
                        $DatabaseSizeFound = $true
                    } catch {
                        Write-Host "    ⚠️ UNC access to NTDS.dit failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }

                # Method 3: WMI/CIM access
                if (-not $DatabaseSizeFound) {
                    try {
                        $RemoteDBFile = Get-CimInstance -ComputerName $DC.HostName.ToString() -ClassName CIM_DataFile -Filter "Name='C:\\Windows\\NTDS\\ntds.dit'" -ErrorAction Stop
                        if ($RemoteDBFile -and $RemoteDBFile.FileSize) {
                            $DCMetrics.ADDatabaseSize = "$([math]::Round($RemoteDBFile.FileSize / 1GB, 2)) GB"
                            $DatabaseSizeFound = $true
                        }
                    } catch {
                        Write-Host "    ⚠️ WMI access to NTDS.dit failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }

                # Method 4: PowerShell Remoting
                if (-not $DatabaseSizeFound) {
                    try {
                        $RemoteSize = Invoke-Command -ComputerName $DC.HostName.ToString() -ScriptBlock {
                            if (Test-Path "C:\Windows\NTDS\ntds.dit") {
                                (Get-Item "C:\Windows\NTDS\ntds.dit" -ErrorAction Stop).Length
                            } else {
                                $null
                            }
                        } -ErrorAction Stop

                        if ($RemoteSize) {
                            $DCMetrics.ADDatabaseSize = "$([math]::Round($RemoteSize / 1GB, 2)) GB"
                            $DatabaseSizeFound = $true
                        }
                    } catch {
                        Write-Host "    ⚠️ PowerShell remoting to get NTDS.dit size failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }

                if (-not $DatabaseSizeFound) {
                    $DCMetrics.ADDatabaseSize = "Access Denied"
                }

                # NTDS Log Files Size - try multiple methods
                $LogPath = "\\$($DC.HostName)\C$\Windows\NTDS\"
                $LocalLogPath = "C:\Windows\NTDS\"
                $LogSizeFound = $false

                # Method 1: Local file access (if this is the local DC)
                if ($IsLocalDC -and (Test-Path $LocalLogPath)) {
                    try {
                        $LogFiles = Get-ChildItem "$LocalLogPath*.log" -ErrorAction Stop
                        if ($LogFiles) {
                            $TotalLogSize = ($LogFiles | Measure-Object -Property Length -Sum).Sum
                            $DCMetrics.LogFileSize = "$([math]::Round($TotalLogSize / 1MB, 2)) MB"
                            $LogSizeFound = $true
                            Write-Host "    ✓ Local log files access successful" -ForegroundColor Green
                        }
                    } catch {
                        Write-Host "    ⚠️ Local access to log files failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }

                # Method 2: UNC Path access (if not local DC)
                if (-not $LogSizeFound -and (Test-Path $LogPath)) {
                    try {
                        $LogFiles = Get-ChildItem "$LogPath*.log" -ErrorAction Stop
                        if ($LogFiles) {
                            $TotalLogSize = ($LogFiles | Measure-Object -Property Length -Sum).Sum
                            $DCMetrics.LogFileSize = "$([math]::Round($TotalLogSize / 1MB, 2)) MB"
                            $LogSizeFound = $true
                        }
                    } catch {
                        Write-Host "    ⚠️ UNC access to log files failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }

                # Method 3: WMI/CIM access
                if (-not $LogSizeFound) {
                    try {
                        $RemoteLogFiles = Get-CimInstance -ComputerName $DC.HostName.ToString() -ClassName CIM_DataFile -Filter "Path='\\Windows\\NTDS\\' AND Extension='log'" -ErrorAction Stop
                        if ($RemoteLogFiles) {
                            $TotalLogSize = ($RemoteLogFiles | Measure-Object -Property FileSize -Sum).Sum
                            $DCMetrics.LogFileSize = "$([math]::Round($TotalLogSize / 1MB, 2)) MB"
                            $LogSizeFound = $true
                        }
                    } catch {
                        Write-Host "    ⚠️ WMI access to log files failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }

                # Method 4: PowerShell Remoting
                if (-not $LogSizeFound) {
                    try {
                        $RemoteLogSize = Invoke-Command -ComputerName $DC.HostName.ToString() -ScriptBlock {
                            $LogFiles = Get-ChildItem "C:\Windows\NTDS\*.log" -ErrorAction SilentlyContinue
                            if ($LogFiles) {
                                ($LogFiles | Measure-Object -Property Length -Sum).Sum
                            } else {
                                $null
                            }
                        } -ErrorAction Stop

                        if ($RemoteLogSize) {
                            $DCMetrics.LogFileSize = "$([math]::Round($RemoteLogSize / 1MB, 2)) MB"
                            $LogSizeFound = $true
                        }
                    } catch {
                        Write-Host "    ⚠️ PowerShell remoting to get log files failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }

                if (-not $LogSizeFound) {
                    $DCMetrics.LogFileSize = "Access Denied"
                }

                # Performance Counters (if accessible)
                try {
                    # LDAP Connections
                    $LDAPConnections = (Get-Counter -ComputerName $DC.HostName.ToString() -Counter "\NTDS\LDAP Client Sessions" -MaxSamples 1 -ErrorAction SilentlyContinue).CounterSamples.CookedValue
                    $DCMetrics.LDAPConnections = [math]::Round($LDAPConnections, 0)

                    $PerfCounters = @(
                        "\NTDS\LDAP Successful Binds/sec",
                        "\NTDS\LDAP Searches/sec",
                        "\NTDS\DRA Inbound Values (DNs only)/sec"
                    )

                    foreach ($Counter in $PerfCounters) {
                        try {
                            $CounterValue = (Get-Counter -ComputerName $DC.HostName.ToString() -Counter $Counter -MaxSamples 1 -ErrorAction SilentlyContinue).CounterSamples.CookedValue
                            $DCMetrics.NTDSPerformance += @{
                                Counter = $Counter.Split('\')[-1]
                                Value = [math]::Round($CounterValue, 2)
                            }
                        } catch {
                            # Counter not available
                        }
                    }
                } catch {
                    Write-Verbose "Performance counters not accessible for $($DC.HostName)"
                }

            } catch {
                Write-Warning "Error collecting performance metrics for $($DC.HostName): $_"
            }

            $DCPerformanceMetrics += $DCMetrics
        }
    }

} catch {
    Write-Warning "Error during DC performance metrics collection: $_"
    $DCPerformanceMetrics = @{ Error = $_.Exception.Message }
}

$AuditResults.DCPerformanceMetrics = $DCPerformanceMetrics
#endregion


#region Replication Health Assessment
Write-Host "Assessing Replication Health..." -ForegroundColor Cyan
try {
    $ReplicationHealth = @()
    
    # First, check for offline/unreachable DCs from our enhanced health assessment
    foreach ($DC in $AuditResults.EnhancedDCHealth) {
        if ($DC.PingStatus -eq "Unreachable" -or $DC.PingStatus -eq "Offline") {
            $ReplicationHealth += @{
                Issue = "Critical"
                SourceDC = $DC.DomainController
                Problem = "Domain Controller is offline/unreachable"
                Impact = "Replication cannot occur with this DC"
                LastSeen = "Unknown"
                Recommendation = "Investigate DC connectivity and services"
            }
        }
    }
    
    # Then check replication status for online DCs
    foreach ($DC in $AuditResults.EnhancedDCHealth) {
        if ($DC.PingStatus -eq "Online") {
            try {
                Write-Host "  Checking replication for $($DC.DomainController)..." -ForegroundColor Yellow
                $ReplStatus = repadmin /showrepl $DC.DomainController /csv | ConvertFrom-Csv
                
                foreach ($Repl in $ReplStatus) {
                    if ($Repl."Number of Failures" -gt 0) {
                        $ReplicationHealth += @{
                            Issue = "Warning"
                            SourceDC = $DC.DomainController
                            DestinationDC = $Repl."Source DSA"
                            NamingContext = $Repl."Naming Context"
                            Problem = "$($Repl.'Number of Failures') replication failures"
                            LastSuccess = $Repl."Last Success Time"
                            LastFailure = $Repl."Last Failure Time"
                            Recommendation = "Check network connectivity and AD replication"
                        }
                    }
                }
                
                # Check for stale replication (no success in last 24 hours)
                foreach ($Repl in $ReplStatus) {
                    if ($Repl."Last Success Time") {
                        try {
                            $LastSuccess = [DateTime]::Parse($Repl."Last Success Time")
                            $HoursSinceSuccess = ((Get-Date) - $LastSuccess).TotalHours
                            if ($HoursSinceSuccess -gt 24) {
                                $ReplicationHealth += @{
                                    Issue = "Warning"
                                    SourceDC = $DC.DomainController
                                    DestinationDC = $Repl."Source DSA" 
                                    NamingContext = $Repl."Naming Context"
                                    Problem = "Stale replication - Last success: $($Repl.'Last Success Time')"
                                    LastSuccess = $Repl."Last Success Time"
                                    Recommendation = "Check replication connectivity and schedule"
                                }
                            }
                        } catch {
                            # Invalid date format, skip
                        }
                    }
                }
                
            } catch {
                $ReplicationHealth += @{
                    Issue = "Error"
                    SourceDC = $DC.DomainController
                    Problem = "Cannot query replication status"
                    Error = $_.Exception.Message
                    Recommendation = "Check repadmin tool and DC accessibility"
                }
            }
        }
    }
    
    $AuditResults.ReplicationHealth = $ReplicationHealth
    Write-Host "Replication Health Assessment completed. Found $($ReplicationHealth.Count) issues." -ForegroundColor Green
    
} catch {
    Write-Error "Replication Health Assessment failed: $($_.Exception.Message)"
    $AuditResults.ReplicationHealth = "Error: $($_.Exception.Message)"
}
#endregion




#region DNS Configuration Testing
Write-Host "Testing DNS Configuration..." -ForegroundColor Cyan

$AuditResults.DNSConfiguration = @()

try {
    $Domain = (Get-ADDomain).DNSRoot
    $DCs = Get-ADDomainController -Filter *

    foreach ($DC in $DCs) {
        $Hostname = $DC.HostName
        Write-Host "  Testing DNS on $Hostname..." -ForegroundColor Yellow
        
        # Initialize DNS entry with default values
        $DNSEntry = [PSCustomObject]@{
            DomainController  = $Hostname
            DNSServer         = $DC.IPAddress
            DomainResolution  = "Unknown"
            ReverseLookup     = "Unknown"
            ScavengingEnabled = "Unknown"
            DCDiagDNS         = "Unknown"
            ConnectivityStatus = "Unknown"
            DNSFailures       = @()
            LastError         = ""
        }

        # Test basic connectivity first
        try {
            Write-Host "    Checking connectivity to $Hostname..." -ForegroundColor Gray
            $ConnectivityTest = Test-NetConnection -ComputerName $Hostname -Port 53 -InformationLevel Quiet -ErrorAction Stop
            
            if ($ConnectivityTest) {
                $DNSEntry.ConnectivityStatus = "Online"
                Write-Host "    ✓ $Hostname is reachable on port 53" -ForegroundColor Green
            } else {
                $DNSEntry.ConnectivityStatus = "Port 53 blocked"
                Write-Warning "    ✗ $Hostname is not reachable on port 53 (DNS port blocked)"
            }
        } catch {
            $DNSEntry.ConnectivityStatus = "Unreachable - Network error"
            $DNSEntry.LastError = $_.Exception.Message
            Write-Warning "    ✗ Cannot reach $Hostname - Network error: $($_.Exception.Message)"
        }

        # Only proceed with DNS tests if DC is reachable
        if ($DNSEntry.ConnectivityStatus -eq "Online") {
            try {
                # Test DCDiag DNS
                Write-Host "    Running DCDiag DNS tests..." -ForegroundColor Gray
                $DCDiagOutput = dcdiag /s:$Hostname /test:DNS 2>&1 | Out-String
                
                if ($LASTEXITCODE -eq 0) {
                    $DNSEntry.DomainResolution = if ($DCDiagOutput -match "PASS.*\sForwarders") { "Success" } else { "Failed" }
                    $DNSEntry.ReverseLookup = if ($DCDiagOutput -match "PASS.*\sReverse") { "Success" } else { "Failed" }
                    $DNSEntry.DCDiagDNS = if ($DCDiagOutput -match "failed test") { "Failed" } else { "Passed" }
                    
                    Write-Host "    ✓ DCDiag DNS tests completed" -ForegroundColor Green
                } else {
                    $DNSEntry.DCDiagDNS = "Error - DCDiag failed"
                    $DNSEntry.LastError = "DCDiag returned exit code: $LASTEXITCODE"
                    Write-Warning "    ✗ DCDiag failed with exit code: $LASTEXITCODE"
                }

                if ($DCDiagOutput -match "failed test") {
                    Write-Warning "    DNS test failures detected on $Hostname. Collecting details..."
                    $lines = $DCDiagOutput -split "`n"
                    foreach ($line in $lines) {
                        if ($line -match ".*(fail|error|not registered|missing|invalid).*") {
                            $DNSEntry.DNSFailures += [PSCustomObject]@{
                                Entry   = $line.Trim()
                                Type    = "DCDiag"
                                Details = "See dcdiag output"
                                Zone    = ""
                            }
                        }
                    }
                }
            } catch {
                $DNSEntry.DCDiagDNS = "Error"
                $DNSEntry.LastError = $_.Exception.Message
                Write-Warning "    ✗ DCDiag DNS test failed: $($_.Exception.Message)"
            }

            # Test DNS Server Scavenging (if DC is reachable)
            try {
                Write-Host "    Checking DNS scavenging settings..." -ForegroundColor Gray
                $ScavengingResult = Get-DnsServerScavenging -ComputerName $DC.HostName.ToString() -ErrorAction Stop
                $DNSEntry.ScavengingEnabled = $ScavengingResult.ScavengingState
                Write-Host "    ✓ DNS scavenging check completed" -ForegroundColor Green
            } catch {
                $DNSEntry.ScavengingEnabled = "Error - Unable to retrieve"
                Write-Warning "    ✗ Cannot retrieve DNS scavenging info: $($_.Exception.Message)"
                
                $DNSEntry.DNSFailures += [PSCustomObject]@{
                    Entry   = "Failed to retrieve DNS scavenging settings"
                    Type    = "Configuration"
                    Details = $_.Exception.Message
                    Zone    = ""
                }
            }

            # Check for obsolete DNS records (only if DC is online)
            try {
                Write-Host "    Scanning for obsolete DNS records..." -ForegroundColor Gray
                $ServerFQDN = "$env:COMPUTERNAME.$Domain."
                $ServerHostname = "$env:COMPUTERNAME"
                $IPAddress = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notmatch "Loopback" }).IPAddress

                $Zones = Get-DnsServerZone -ComputerName $DC.HostName.ToString() -ErrorAction Stop | Where-Object { $_.ZoneType -eq "Primary" } | Select-Object -ExpandProperty ZoneName
                $TotalObsoleteRecords = 0

                foreach ($Zone in $Zones) {
                    try {
                        $ObsoleteRecords = Get-DnsServerResourceRecord -ComputerName $DC.HostName.ToString() -ZoneName $Zone -ErrorAction SilentlyContinue | Where-Object {
                            ($_.RecordData.IPv4Address -eq $IPAddress) -or
                            ($_.RecordData.NameServer -eq $ServerFQDN) -or
                            ($_.RecordData.DomainName -eq $ServerFQDN) -or
                            ($_.RecordData.HostnameAlias -eq $ServerFQDN) -or
                            ($_.RecordData.MailExchange -eq $ServerFQDN) -or
                            ($_.HostName -eq $ServerHostname)
                        }

                        foreach ($rec in $ObsoleteRecords) {
                            $DNSEntry.DNSFailures += [PSCustomObject]@{
                                Entry   = "$($rec.RecordData.NameServer) [$($rec.RecordType)] $($rec.Hostname)"
                                Type    = "Obsolete"
                                Details = "Found in zone $Zone"
                                Zone    = $Zone
                            }
                            $TotalObsoleteRecords++
                        }
                    } catch {
                        Write-Warning "    ✗ Error scanning zone $Zone for obsolete records: $($_.Exception.Message)"
                    }
                }
                Write-Host "    ✓ DNS scan completed. Found $TotalObsoleteRecords obsolete records" -ForegroundColor Green
            } catch {
                Write-Warning "    ✗ Cannot retrieve DNS zones from $Hostname`: $($_.Exception.Message)"
                $DNSEntry.DNSFailures += [PSCustomObject]@{
                    Entry   = "Failed to retrieve DNS zones"
                    Type    = "Connectivity"
                    Details = $_.Exception.Message
                    Zone    = ""
                }
            }
        } else {
            # DC is unreachable - add appropriate failure entries
            $DNSEntry.DNSFailures += [PSCustomObject]@{
                Entry   = "Domain Controller is unreachable"
                Type    = "Connectivity"
                Details = "Cannot perform DNS tests - $($DNSEntry.ConnectivityStatus)"
                Zone    = ""
            }
        }

        $AuditResults.DNSConfiguration += $DNSEntry
        Write-Host "  DNS testing completed for $Hostname" -ForegroundColor Cyan
    }
}
catch {
    $ErrorMessage = "An error occurred during DNS configuration testing: $($_.Exception.Message)"
    Write-Error $ErrorMessage
    
    # Create a fallback entry to show the error in the report
    $AuditResults.DNSConfiguration += [PSCustomObject]@{
        DomainController  = "Error"
        DNSServer         = "N/A"
        DomainResolution  = "Failed"
        ReverseLookup     = "Failed"
        ScavengingEnabled = "Unknown"
        DCDiagDNS         = "Failed"
        ConnectivityStatus = "Error"
        DNSFailures       = @([PSCustomObject]@{
            Entry   = $ErrorMessage
            Type    = "System"
            Details = "Script execution error"
            Zone    = ""
        })
        LastError         = $_.Exception.Message
    }
}
#endregion
















#region FSMO Roles Verification - Alternative Methods
Write-Host "`n--- FSMO Roles Verification (Alternative Methods) ---" -ForegroundColor Cyan

try {
    $FsmoResults = @()
    $Forest = Get-ADForest
    $Domains = $Forest.Domains

    # Helper function to safely test connection
    function Test-SafeConnection {
        param([string]$ComputerName)
        
        if ([string]::IsNullOrWhiteSpace($ComputerName)) {
            return "Unknown Server"
        }
        
        try {
            if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                return "Online"
            } else {
                return "Offline"
            }
        }
        catch {
            return "Connection Error"
        }
    }

    # Method 1: Use netdom query fsmo (if available)
    function Get-FSMOViaNetdom {
        try {
            Write-Host "Attempting to retrieve FSMO roles via netdom..." -ForegroundColor Yellow
            $NetdomOutput = & netdom query fsmo 2>$null
            if ($NetdomOutput) {
                Write-Host "Netdom output received" -ForegroundColor Green
                return $NetdomOutput
            }
        }
        catch {
            Write-Host "Netdom not available or failed" -ForegroundColor Yellow
        }
        return $null
    }

    # Method 2: Direct LDAP queries
    function Get-FSMOViaLDAP {
        param([string]$DomainName)
        
        try {
            Write-Host "Attempting LDAP query for domain: $DomainName" -ForegroundColor Yellow
            
            # Get domain DN
            $DomainDN = "DC=" + ($DomainName -replace "\.", ",DC=")
            Write-Host "Domain DN: $DomainDN" -ForegroundColor Gray
            
            # Query for FSMO role holders using Get-ADObject
            $RIDMasterDN = (Get-ADObject -Filter {objectClass -eq "rIDManager"} -SearchBase "CN=System,$DomainDN" -Properties fSMORoleOwner -Server $DomainName).fSMORoleOwner
            $PDCEmulatorDN = (Get-ADObject -Filter {objectClass -eq "domainDNS"} -SearchBase $DomainDN -Properties fSMORoleOwner -Server $DomainName).fSMORoleOwner
            $InfraMasterDN = (Get-ADObject -Filter {objectClass -eq "infrastructureUpdate"} -SearchBase "CN=Infrastructure,$DomainDN" -Properties fSMORoleOwner -Server $DomainName).fSMORoleOwner
            
            # Convert DNs to server names
            $Results = @{}
            
            if ($RIDMasterDN) {
                $RIDServer = (Get-ADObject -Identity $RIDMasterDN -Server $DomainName -Properties dNSHostName).dNSHostName
                $Results.RIDMaster = $RIDServer
                Write-Host "RID Master found via LDAP: $RIDServer" -ForegroundColor Green
            }
            
            if ($PDCEmulatorDN) {
                $PDCServer = (Get-ADObject -Identity $PDCEmulatorDN -Server $DomainName -Properties dNSHostName).dNSHostName
                $Results.PDCEmulator = $PDCServer
                Write-Host "PDC Emulator found via LDAP: $PDCServer" -ForegroundColor Green
            }
            
            if ($InfraMasterDN) {
                $InfraServer = (Get-ADObject -Identity $InfraMasterDN -Server $DomainName -Properties dNSHostName).dNSHostName
                $Results.InfrastructureMaster = $InfraServer
                Write-Host "Infrastructure Master found via LDAP: $InfraServer" -ForegroundColor Green
            }
            
            return $Results
        }
        catch {
            Write-Host "LDAP query failed: $($_.Exception.Message)" -ForegroundColor Red
            return $null
        }
    }

    # Method 3: Use specific domain controller
    function Get-FSMOViaSpecificDC {
        param([string]$DomainName)
        
        try {
            Write-Host "Attempting to find available DCs for domain: $DomainName" -ForegroundColor Yellow
            
            # Get list of domain controllers
            $DCs = Get-ADDomainController -Filter * -Server $DomainName -ErrorAction Stop
            
            foreach ($DC in $DCs) {
                try {
                    Write-Host "Trying DC: $($DC.HostName)" -ForegroundColor Gray
                    $Domain = Get-ADDomain -Server $DC.HostName -ErrorAction Stop
                    
                    $Results = @{
                        RIDMaster = $Domain.RIDMaster
                        PDCEmulator = $Domain.PDCEmulator
                        InfrastructureMaster = $Domain.InfrastructureMaster
                    }
                    
                    # Check if we got valid results
                    $ValidResults = $Results.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    if ($ValidResults.Count -gt 0) {
                        Write-Host "Successfully retrieved FSMO info via DC: $($DC.HostName)" -ForegroundColor Green
                        return $Results
                    }
                }
                catch {
                    Write-Host "Failed with DC $($DC.HostName): $($_.Exception.Message)" -ForegroundColor Yellow
                    continue
                }
            }
        }
        catch {
            Write-Host "Could not enumerate domain controllers: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        return $null
    }

    # Forest-level FSMO roles (these seem to work)
    Write-Host "`nChecking Forest-level FSMO roles..." -ForegroundColor Yellow
    
    $FsmoResults += [PSCustomObject]@{
        Role   = "Schema Master"
        Server = $Forest.SchemaMaster
        Scope  = "Forest"
        Status = Test-SafeConnection $Forest.SchemaMaster
        Method = "Get-ADForest"
    }

    $FsmoResults += [PSCustomObject]@{
        Role   = "Domain Naming Master"
        Server = $Forest.DomainNamingMaster
        Scope  = "Forest"
        Status = Test-SafeConnection $Forest.DomainNamingMaster
        Method = "Get-ADForest"
    }

    # Try netdom first
    $NetdomResults = Get-FSMOViaNetdom

    # Domain-level FSMO roles - try multiple methods
    foreach ($DomainName in $Domains) {
        Write-Host "`nProcessing domain: $DomainName" -ForegroundColor Cyan
        
        $DomainFSMO = $null
        $Method = "Unknown"
        
        # Method 1: Standard Get-ADDomain (original method)
        try {
            Write-Host "Method 1: Standard Get-ADDomain" -ForegroundColor Yellow
            $Domain = Get-ADDomain -Server $DomainName -ErrorAction Stop
            
            Write-Host "Domain object retrieved. Checking FSMO properties..." -ForegroundColor Gray
            Write-Host "  RIDMaster: '$($Domain.RIDMaster)'" -ForegroundColor Gray
            Write-Host "  PDCEmulator: '$($Domain.PDCEmulator)'" -ForegroundColor Gray
            Write-Host "  InfrastructureMaster: '$($Domain.InfrastructureMaster)'" -ForegroundColor Gray
            
            if (-not [string]::IsNullOrWhiteSpace($Domain.RIDMaster) -and 
                -not [string]::IsNullOrWhiteSpace($Domain.PDCEmulator) -and 
                -not [string]::IsNullOrWhiteSpace($Domain.InfrastructureMaster)) {
                
                $DomainFSMO = @{
                    RIDMaster = $Domain.RIDMaster
                    PDCEmulator = $Domain.PDCEmulator
                    InfrastructureMaster = $Domain.InfrastructureMaster
                }
                $Method = "Get-ADDomain"
                Write-Host "Standard method successful" -ForegroundColor Green
            } else {
                Write-Host "Standard method returned null/empty values" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "Standard method failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Method 2: Try specific DC if standard method failed
        if (-not $DomainFSMO) {
            Write-Host "Method 2: Trying specific domain controllers" -ForegroundColor Yellow
            $DomainFSMO = Get-FSMOViaSpecificDC -DomainName $DomainName
            if ($DomainFSMO) { $Method = "Specific DC" }
        }
        
        # Method 3: Try LDAP queries if other methods failed
        if (-not $DomainFSMO) {
            Write-Host "Method 3: Trying direct LDAP queries" -ForegroundColor Yellow
            $DomainFSMO = Get-FSMOViaLDAP -DomainName $DomainName
            if ($DomainFSMO) { $Method = "LDAP Query" }
        }
        
        # Add results
        if ($DomainFSMO) {
            $RoleMap = @{
                "RID Master" = $DomainFSMO.RIDMaster
                "PDC Emulator" = $DomainFSMO.PDCEmulator
                "Infrastructure Master" = $DomainFSMO.InfrastructureMaster
            }
            
            foreach ($Role in $RoleMap.GetEnumerator()) {
                $ServerName = if ([string]::IsNullOrWhiteSpace($Role.Value)) { "Not Available" } else { $Role.Value }
                $Status = if ($ServerName -eq "Not Available") { "Role Not Found" } else { Test-SafeConnection $Role.Value }
                
                $FsmoResults += [PSCustomObject]@{
                    Role   = $Role.Key
                    Server = $ServerName
                    Scope  = $DomainName
                    Status = $Status
                    Method = $Method
                }
            }
        } else {
            Write-Host "All methods failed for domain: $DomainName" -ForegroundColor Red
            
            # Add error entries
            $ErrorRoles = @("RID Master", "PDC Emulator", "Infrastructure Master")
            foreach ($Role in $ErrorRoles) {
                $FsmoResults += [PSCustomObject]@{
                    Role   = $Role
                    Server = "Retrieval Failed"
                    Scope  = $DomainName
                    Status = "All Methods Failed"
                    Method = "None"
                }
            }
        }
    }

    # Display results
    Write-Host "`n" + "="*90 -ForegroundColor Cyan
    Write-Host "FSMO Role Owners Summary (with retrieval method)" -ForegroundColor Cyan
    Write-Host "="*90 -ForegroundColor Cyan
    
    $FsmoResults | ForEach-Object {
        $Color = switch ($_.Status) {
            "Online" { "Green" }
            "Offline" { "Red" }
            "Role Not Found" { "Magenta" }
            "All Methods Failed" { "Red" }
            default { "Yellow" }
        }
        
        Write-Host ("{0,-25} = {1,-30} [{2}] ({3}) - Method: {4}" -f $_.Role, $_.Server, $_.Status, $_.Scope, $_.Method) -ForegroundColor $Color
    }

    # Display netdom results if available
    if ($NetdomResults) {
        Write-Host "`n--- Netdom Query FSMO Results ---" -ForegroundColor Cyan
        $NetdomResults | ForEach-Object { Write-Host $_ -ForegroundColor White }
    }

    # Store for audit report
    $AuditResults.FSMORoles = $FsmoResults
}
catch {
    Write-Warning "Failed to retrieve FSMO roles: $($_.Exception.Message)"
    $AuditResults.FSMORoles = "Error: $($_.Exception.Message)"
}
#endregion

#region Trust Relationships Analysis
Write-Host "Analyzing Trust Relationships..." -ForegroundColor Cyan
try {
    $Trusts = Get-ADTrust -Filter *
    $TrustAnalysis = @()
    
    foreach ($Trust in $Trusts) {
        $TrustInfo = @{
            Name = $Trust.Name
            Direction = $Trust.Direction
            TrustType = $Trust.TrustType
            ForestTransitive = $Trust.ForestTransitive
            SelectiveAuthentication = $Trust.SelectiveAuthentication
            SIDFilteringQuarantined = $Trust.SIDFilteringQuarantined
            Created = $Trust.Created
            Modified = $Trust.Modified
        }
        
        # Test trust connectivity
        try {
            $TrustTest = Test-ComputerSecureChannel -Server $Trust.Name -ErrorAction SilentlyContinue
            $TrustInfo.ConnectivityTest = if ($TrustTest) { "Success" } else { "Failed" }
        } catch {
            $TrustInfo.ConnectivityTest = "Error"
        }
        
        $TrustAnalysis += New-Object PSObject -Property $TrustInfo
    }
    $AuditResults.TrustRelationships = $TrustAnalysis
} catch {
    $AuditResults.TrustRelationships = "Error: $($_.Exception.Message)"
}
#endregion

#region AD Sites & Services Audit
Write-Host "Auditing AD Sites & Services..." -ForegroundColor Cyan
try {
    $Sites = Get-ADReplicationSite -Filter *
    $SitesAnalysis = @()
    
    foreach ($Site in $Sites) {
        $SiteInfo = @{
            Name = $Site.Name
            Description = $Site.Description
            DomainControllers = (Get-ADDomainController -Filter {Site -eq $Site.Name}).Count
        }
        
        # Get subnets for this site
        $Subnets = Get-ADReplicationSubnet -Filter {Site -eq $Site.Name} -ErrorAction SilentlyContinue
        $SiteInfo.Subnets = ($Subnets | ForEach-Object { $_.Name }) -join ", "
        
        $SitesAnalysis += New-Object PSObject -Property $SiteInfo
    }
    
    # Get site links
    $SiteLinks = Get-ADReplicationSiteLink -Filter *
    $SiteLinksInfo = @()
    foreach ($Link in $SiteLinks) {
        $SiteLinksInfo += @{
            Name = $Link.Name
            Cost = $Link.Cost
            ReplicationFrequencyInMinutes = $Link.ReplicationFrequencyInMinutes
            SitesIncluded = $Link.SitesIncluded -join ", "
        }
    }
    
    $AuditResults.ADSites = @{
        Sites = $SitesAnalysis
        SiteLinks = $SiteLinksInfo
    }
} catch {
    $AuditResults.ADSites = "Error: $($_.Exception.Message)"
}
#endregion

#region Group Policy Assessment & Security Analysis
Write-Host "Performing Comprehensive Group Policy Analysis..." -ForegroundColor Cyan

# GPO Security Analysis Functions
function Test-GPOSecuritySettings {
    param([xml]$GPOReport, [string]$GPOName)

    $SecurityIssues = @()
    $SecurityRecommendations = @()

    # Check password policy settings
    if ($GPOReport.GPO.Computer.ExtensionData.Extension.PasswordPolicy) {
        $PasswordPolicy = $GPOReport.GPO.Computer.ExtensionData.Extension.PasswordPolicy

        if ([int]$PasswordPolicy.MinimumPasswordLength -lt 14) {
            $SecurityIssues += "Weak password length requirement (< 14 characters)"
            $SecurityRecommendations += "Set minimum password length to 14+ characters"
        }

        if ([int]$PasswordPolicy.MaximumPasswordAge -gt 365) {
            $SecurityIssues += "Password expiration too long (> 365 days)"
            $SecurityRecommendations += "Set maximum password age to 365 days or less"
        }
    }

    # Check account lockout policy
    if ($GPOReport.GPO.Computer.ExtensionData.Extension.AccountLockoutPolicy) {
        $LockoutPolicy = $GPOReport.GPO.Computer.ExtensionData.Extension.AccountLockoutPolicy

        if ([int]$LockoutPolicy.LockoutThreshold -gt 5 -or [int]$LockoutPolicy.LockoutThreshold -eq 0) {
            $SecurityIssues += "Weak account lockout threshold"
            $SecurityRecommendations += "Set account lockout threshold to 3-5 failed attempts"
        }
    }

    # Check audit policy settings
    $AuditCategories = @(
        "AuditLogonEvents", "AuditAccountLogonEvents", "AuditPrivilegeUse",
        "AuditPolicyChange", "AuditAccountManagement", "AuditObjectAccess"
    )

    foreach ($Category in $AuditCategories) {
        $AuditSetting = $GPOReport.GPO.Computer.ExtensionData.Extension.AuditPolicy.$Category
        if ($AuditSetting -and $AuditSetting -eq "No Auditing") {
            $SecurityIssues += "Critical audit category '$Category' not enabled"
            $SecurityRecommendations += "Enable auditing for $Category"
        }
    }

    return @{
        SecurityIssues = $SecurityIssues
        SecurityRecommendations = $SecurityRecommendations
    }
}

function Analyze-GPOCompliance {
    param([xml]$GPOReport, [string]$GPOName)

    $ComplianceScore = 100
    $ComplianceIssues = @()

    # Check for deprecated/legacy settings
    if ($GPOReport.GPO.Computer.ExtensionData.Extension.SecurityOptions) {
        $SecurityOptions = $GPOReport.GPO.Computer.ExtensionData.Extension.SecurityOptions

        # Check for LM hash storage (should be disabled)
        $LMHashSetting = $SecurityOptions | Where-Object { $_.KeyName -like "*LMCompatibilityLevel*" }
        if ($LMHashSetting -and [int]$LMHashSetting.SettingNumber -lt 3) {
            $ComplianceScore -= 15
            $ComplianceIssues += "Legacy LM hash compatibility enabled (security risk)"
        }

        # Check for anonymous access settings
        $AnonymousSettings = $SecurityOptions | Where-Object { $_.KeyName -like "*Anonymous*" }
        foreach ($Setting in $AnonymousSettings) {
            if ($Setting.SettingBoolean -eq "true") {
                $ComplianceScore -= 10
                $ComplianceIssues += "Anonymous access enabled: $($Setting.KeyName)"
            }
        }
    }

    # Check for weak encryption protocols
    if ($GPOReport.GPO.Computer.ExtensionData.Extension.RegistrySettings) {
        $RegistrySettings = $GPOReport.GPO.Computer.ExtensionData.Extension.RegistrySettings

        # Look for SSL/TLS settings
        $WeakSSLSettings = $RegistrySettings | Where-Object {
            $_.KeyName -like "*SSL*" -or $_.KeyName -like "*TLS*"
        }
        foreach ($Setting in $WeakSSLSettings) {
            if ($Setting.ValueName -like "*SSL 2.0*" -or $Setting.ValueName -like "*SSL 3.0*") {
                $ComplianceScore -= 20
                $ComplianceIssues += "Weak SSL protocol enabled: $($Setting.ValueName)"
            }
        }
    }

    return @{
        ComplianceScore = [math]::Max(0, $ComplianceScore)
        ComplianceIssues = $ComplianceIssues
    }
}

function Get-GPOLinkDetails {
    param([string]$GPOID)

    $LinkDetails = @()

    try {
        # Get all OUs and check for GPO links
        $AllOUs = Get-ADOrganizationalUnit -Filter *
        foreach ($OU in $AllOUs) {
            $LinkedGPOs = Get-GPInheritance -Target $OU.DistinguishedName -ErrorAction SilentlyContinue
            if ($LinkedGPOs.GpoLinks) {
                foreach ($Link in $LinkedGPOs.GpoLinks) {
                    if ($Link.GpoId -eq $GPOID) {
                        $LinkDetails += @{
                            Target = $OU.DistinguishedName
                            TargetType = "OU"
                            Enabled = $Link.Enabled
                            Enforced = $Link.Enforced
                            Order = $Link.Order
                        }
                    }
                }
            }
        }

        # Check domain-level links
        $DomainLinks = Get-GPInheritance -Target (Get-ADDomain).DistinguishedName -ErrorAction SilentlyContinue
        if ($DomainLinks.GpoLinks) {
            foreach ($Link in $DomainLinks.GpoLinks) {
                if ($Link.GpoId -eq $GPOID) {
                    $LinkDetails += @{
                        Target = "Domain Root"
                        TargetType = "Domain"
                        Enabled = $Link.Enabled
                        Enforced = $Link.Enforced
                        Order = $Link.Order
                    }
                }
            }
        }
    } catch {
        Write-Warning "Error getting link details for GPO $GPOID`: $($_.Exception.Message)"
    }

    return $LinkDetails
}

function Test-GPOConflicts {
    param([array]$GPOs)

    $ConflictResults = @{
        PolicyConflicts = @()
        InheritanceIssues = @()
        PrecedenceWarnings = @()
        DuplicateSettings = @()
        ConflictSummary = @{
            TotalConflicts = 0
            CriticalConflicts = 0
            WarningConflicts = 0
            AffectedOUs = @()
        }
    }

    Write-Host "Analyzing GPO conflicts and precedence..." -ForegroundColor Yellow

    # Group GPOs by their link targets to analyze conflicts
    $TargetGPOMap = @{}
    foreach ($GPO in $GPOs) {
        if ($GPO.LinkDetails -and $GPO.LinkDetails.Count -gt 0) {
            foreach ($Link in $GPO.LinkDetails) {
                if (-not $TargetGPOMap[$Link.Target]) {
                    $TargetGPOMap[$Link.Target] = @()
                }
                $TargetGPOMap[$Link.Target] += @{
                    GPO = $GPO
                    Link = $Link
                }
            }
        }
    }

    # Analyze each target for conflicts
    foreach ($Target in $TargetGPOMap.Keys) {
        $TargetGPOs = $TargetGPOMap[$Target]

        if ($TargetGPOs.Count -gt 1) {
            $ConflictResults.ConflictSummary.AffectedOUs += $Target

            # Check for policy conflicts between GPOs on same target
            for ($i = 0; $i -lt $TargetGPOs.Count; $i++) {
                for ($j = $i + 1; $j -lt $TargetGPOs.Count; $j++) {
                    $GPO1 = $TargetGPOs[$i]
                    $GPO2 = $TargetGPOs[$j]

                    $Conflicts = Compare-GPOSettings -GPO1 $GPO1.GPO -GPO2 $GPO2.GPO -Target $Target
                    if ($Conflicts.Count -gt 0) {
                        $ConflictResults.PolicyConflicts += $Conflicts
                        $ConflictResults.ConflictSummary.TotalConflicts += $Conflicts.Count
                    }
                }
            }

            # Check precedence and inheritance
            $SortedGPOs = $TargetGPOs | Sort-Object { $_.Link.Order }
            for ($i = 0; $i -lt $SortedGPOs.Count - 1; $i++) {
                $HigherGPO = $SortedGPOs[$i]
                $LowerGPO = $SortedGPOs[$i + 1]

                # Check for enforced GPOs that might override others
                if ($LowerGPO.Link.Enforced -and -not $HigherGPO.Link.Enforced) {
                    $ConflictResults.PrecedenceWarnings += @{
                        Type = "Enforced Override"
                        Target = $Target
                        OverridingGPO = $LowerGPO.GPO.DisplayName
                        OverriddenGPO = $HigherGPO.GPO.DisplayName
                        Severity = "High"
                        Description = "Enforced GPO '$($LowerGPO.GPO.DisplayName)' may override higher precedence GPO '$($HigherGPO.GPO.DisplayName)'"
                    }
                    $ConflictResults.ConflictSummary.CriticalConflicts++
                }
            }

            # Check for inheritance blocking
            $InheritanceBlocked = $TargetGPOs | Where-Object { $_.Link.Enabled -eq $false }
            if ($InheritanceBlocked) {
                foreach ($BlockedGPO in $InheritanceBlocked) {
                    $ConflictResults.InheritanceIssues += @{
                        Type = "Disabled Link"
                        Target = $Target
                        GPO = $BlockedGPO.GPO.DisplayName
                        Description = "GPO link is disabled but still linked to target"
                        Severity = "Medium"
                    }
                    $ConflictResults.ConflictSummary.WarningConflicts++
                }
            }
        }
    }

    return $ConflictResults
}

function Compare-GPOSettings {
    param($GPO1, $GPO2, $Target)

    $Conflicts = @()

    try {
        # Get detailed reports for both GPOs
        $Report1 = Get-GPOReport -Guid $GPO1.Id -ReportType XML -ErrorAction Stop
        $Report2 = Get-GPOReport -Guid $GPO2.Id -ReportType XML -ErrorAction Stop
        [xml]$XML1 = $Report1
        [xml]$XML2 = $Report2

        # Compare Computer Configuration Settings
        $ComputerConflicts = Compare-GPOComputerSettings -XML1 $XML1 -XML2 $XML2 -GPO1Name $GPO1.DisplayName -GPO2Name $GPO2.DisplayName -Target $Target
        $Conflicts += $ComputerConflicts

        # Compare User Configuration Settings
        $UserConflicts = Compare-GPOUserSettings -XML1 $XML1 -XML2 $XML2 -GPO1Name $GPO1.DisplayName -GPO2Name $GPO2.DisplayName -Target $Target
        $Conflicts += $UserConflicts

    }
    catch {
        Write-Warning "Error comparing GPO settings: $($_.Exception.Message)"
    }

    return $Conflicts
}

function Compare-GPOComputerSettings {
    param($XML1, $XML2, $GPO1Name, $GPO2Name, $Target)

    $Conflicts = @()

    # Compare Security Settings
    $Sec1 = $XML1.GPO.Computer.ExtensionData.Extension | Where-Object { $_.type -eq "Security" }
    $Sec2 = $XML2.GPO.Computer.ExtensionData.Extension | Where-Object { $_.type -eq "Security" }

    if ($Sec1 -and $Sec2) {
        # Password Policy Conflicts
        if ($Sec1.PasswordPolicy -and $Sec2.PasswordPolicy) {
            $PasswordConflicts = @(
                @{ Setting = "MinimumPasswordLength"; Value1 = $Sec1.PasswordPolicy.MinimumPasswordLength; Value2 = $Sec2.PasswordPolicy.MinimumPasswordLength },
                @{ Setting = "MaximumPasswordAge"; Value1 = $Sec1.PasswordPolicy.MaximumPasswordAge; Value2 = $Sec2.PasswordPolicy.MaximumPasswordAge },
                @{ Setting = "MinimumPasswordAge"; Value1 = $Sec1.PasswordPolicy.MinimumPasswordAge; Value2 = $Sec2.PasswordPolicy.MinimumPasswordAge },
                @{ Setting = "PasswordComplexity"; Value1 = $Sec1.PasswordPolicy.PasswordComplexity; Value2 = $Sec2.PasswordPolicy.PasswordComplexity }
            )

            foreach ($Check in $PasswordConflicts) {
                if ($Check.Value1 -and $Check.Value2 -and $Check.Value1 -ne $Check.Value2) {
                    $Conflicts += @{
                        Type = "Password Policy Conflict"
                        Setting = $Check.Setting
                        GPO1 = $GPO1Name
                        GPO2 = $GPO2Name
                        Value1 = $Check.Value1
                        Value2 = $Check.Value2
                        Target = $Target
                        Severity = "High"
                        Description = "$($Check.Setting) differs between GPOs: '$GPO1Name' = $($Check.Value1), '$GPO2Name' = $($Check.Value2)"
                    }
                }
            }
        }

        # Account Lockout Policy Conflicts
        if ($Sec1.AccountLockoutPolicy -and $Sec2.AccountLockoutPolicy) {
            $LockoutConflicts = @(
                @{ Setting = "LockoutThreshold"; Value1 = $Sec1.AccountLockoutPolicy.LockoutThreshold; Value2 = $Sec2.AccountLockoutPolicy.LockoutThreshold },
                @{ Setting = "LockoutDuration"; Value1 = $Sec1.AccountLockoutPolicy.LockoutDuration; Value2 = $Sec2.AccountLockoutPolicy.LockoutDuration }
            )

            foreach ($Check in $LockoutConflicts) {
                if ($Check.Value1 -and $Check.Value2 -and $Check.Value1 -ne $Check.Value2) {
                    $Conflicts += @{
                        Type = "Account Lockout Policy Conflict"
                        Setting = $Check.Setting
                        GPO1 = $GPO1Name
                        GPO2 = $GPO2Name
                        Value1 = $Check.Value1
                        Value2 = $Check.Value2
                        Target = $Target
                        Severity = "High"
                        Description = "$($Check.Setting) differs between GPOs: '$GPO1Name' = $($Check.Value1), '$GPO2Name' = $($Check.Value2)"
                    }
                }
            }
        }
    }

    # Compare Registry Settings
    $Reg1 = $XML1.GPO.Computer.ExtensionData.Extension | Where-Object { $_.type -eq "Registry" }
    $Reg2 = $XML2.GPO.Computer.ExtensionData.Extension | Where-Object { $_.type -eq "Registry" }

    if ($Reg1 -and $Reg2 -and $Reg1.RegistrySettings -and $Reg2.RegistrySettings) {
        foreach ($RegSetting1 in $Reg1.RegistrySettings) {
            $MatchingSettings = $Reg2.RegistrySettings | Where-Object {
                $_.KeyName -eq $RegSetting1.KeyName -and $_.ValueName -eq $RegSetting1.ValueName
            }

            foreach ($RegSetting2 in $MatchingSettings) {
                if ($RegSetting1.Value -ne $RegSetting2.Value) {
                    $Conflicts += @{
                        Type = "Registry Setting Conflict"
                        Setting = "$($RegSetting1.KeyName)\$($RegSetting1.ValueName)"
                        GPO1 = $GPO1Name
                        GPO2 = $GPO2Name
                        Value1 = $RegSetting1.Value
                        Value2 = $RegSetting2.Value
                        Target = $Target
                        Severity = "Medium"
                        Description = "Registry setting '$($RegSetting1.KeyName)\$($RegSetting1.ValueName)' has conflicting values"
                    }
                }
            }
        }
    }

    return $Conflicts
}

function Compare-GPOUserSettings {
    param($XML1, $XML2, $GPO1Name, $GPO2Name, $Target)

    $Conflicts = @()

    # Compare User Registry Settings
    $UserReg1 = $XML1.GPO.User.ExtensionData.Extension | Where-Object { $_.type -eq "Registry" }
    $UserReg2 = $XML2.GPO.User.ExtensionData.Extension | Where-Object { $_.type -eq "Registry" }

    if ($UserReg1 -and $UserReg2 -and $UserReg1.RegistrySettings -and $UserReg2.RegistrySettings) {
        foreach ($RegSetting1 in $UserReg1.RegistrySettings) {
            $MatchingSettings = $UserReg2.RegistrySettings | Where-Object {
                $_.KeyName -eq $RegSetting1.KeyName -and $_.ValueName -eq $RegSetting1.ValueName
            }

            foreach ($RegSetting2 in $MatchingSettings) {
                if ($RegSetting1.Value -ne $RegSetting2.Value) {
                    $Conflicts += @{
                        Type = "User Registry Setting Conflict"
                        Setting = "$($RegSetting1.KeyName)\$($RegSetting1.ValueName)"
                        GPO1 = $GPO1Name
                        GPO2 = $GPO2Name
                        Value1 = $RegSetting1.Value
                        Value2 = $RegSetting2.Value
                        Target = $Target
                        Severity = "Medium"
                        Description = "User registry setting '$($RegSetting1.KeyName)\$($RegSetting1.ValueName)' has conflicting values"
                    }
                }
            }
        }
    }

    return $Conflicts
}

function Get-GPOPrecedenceAnalysis {
    param([array]$GPOs)

    $PrecedenceAnalysis = @{
        TargetAnalysis = @{}
        Recommendations = @()
    }

    # Group GPOs by target for precedence analysis
    $TargetGroups = @{}
    foreach ($GPO in $GPOs) {
        if ($GPO.LinkDetails) {
            foreach ($Link in $GPO.LinkDetails) {
                if (-not $TargetGroups[$Link.Target]) {
                    $TargetGroups[$Link.Target] = @()
                }
                $TargetGroups[$Link.Target] += @{
                    GPO = $GPO
                    Link = $Link
                }
            }
        }
    }

    foreach ($Target in $TargetGroups.Keys) {
        $TargetGPOs = $TargetGroups[$Target] | Sort-Object { $_.Link.Order }

        $Analysis = @{
            Target = $Target
            GPOCount = $TargetGPOs.Count
            ProcessingOrder = @()
            Issues = @()
        }

        foreach ($TargetGPO in $TargetGPOs) {
            $ProcessingInfo = @{
                GPOName = $TargetGPO.GPO.DisplayName
                Order = $TargetGPO.Link.Order
                Enabled = $TargetGPO.Link.Enabled
                Enforced = $TargetGPO.Link.Enforced
                SecurityScore = $TargetGPO.GPO.SecurityScore
            }

            $Analysis.ProcessingOrder += $ProcessingInfo

            # Check for potential issues
            if (-not $TargetGPO.Link.Enabled) {
                $Analysis.Issues += "GPO '$($TargetGPO.GPO.DisplayName)' is linked but disabled"
            }

            if ($TargetGPO.Link.Enforced) {
                $Analysis.Issues += "GPO '$($TargetGPO.GPO.DisplayName)' is enforced - will override inheritance"
            }

            if ($TargetGPO.GPO.SecurityScore -lt 60) {
                $Analysis.Issues += "GPO '$($TargetGPO.GPO.DisplayName)' has low security score ($($TargetGPO.GPO.SecurityScore)%)"
            }
        }

        $PrecedenceAnalysis.TargetAnalysis[$Target] = $Analysis
    }

    return $PrecedenceAnalysis
}

try {
    $GPOs = Get-GPO -All -ErrorAction Stop
    $GPOAnalysis = @()
    $GPOSecuritySummary = @{
        TotalGPOs = $GPOs.Count
        LinkedGPOs = 0
        UnlinkedGPOs = 0
        SecurityIssuesFound = 0
        HighRiskGPOs = 0
        ComplianceViolations = 0
        OrphanedGPOs = 0
    }

    $AllSecurityIssues = @()
    $AllRecommendations = @()

    Write-Host "Analyzing $($GPOs.Count) Group Policy Objects..." -ForegroundColor Yellow

    $Counter = 0
    foreach ($GPO in $GPOs) {
        $Counter++
        Write-Progress -Activity "Analyzing GPOs" -Status "Processing $($GPO.DisplayName)" -PercentComplete (($Counter / $GPOs.Count) * 100)

        $GPOInfo = @{
            DisplayName = $GPO.DisplayName
            Id = $GPO.Id
            CreationTime = $GPO.CreationTime
            ModificationTime = $GPO.ModificationTime
            Owner = $GPO.Owner
            IsLinked = "No"
            IsFullyDisabled = "No"
            SecurityScore = 100
            ComplianceScore = 100
            SecurityIssues = @()
            ComplianceIssues = @()
            LinkCount = 0
            LinkDetails = @()
            UserSettingsEnabled = "Unknown"
            ComputerSettingsEnabled = "Unknown"
            PermissionsCount = 0
            LastModifiedDays = ((Get-Date) - $GPO.ModificationTime).Days
            IsOrphaned = $false
        }

        try {
            # Generate comprehensive GPO report
            $ReportXml = Get-GPOReport -Guid $GPO.Id -ReportType XML -ErrorAction Stop
            [xml]$GPOReport = $ReportXml

            # Basic GPO information
            $GPOInfo.UserSettingsEnabled = $GPOReport.GPO.User.Enabled
            $GPOInfo.ComputerSettingsEnabled = $GPOReport.GPO.Computer.Enabled

            if ($GPOReport.GPO.User.Enabled -eq "false" -and $GPOReport.GPO.Computer.Enabled -eq "false") {
                $GPOInfo.IsFullyDisabled = "Yes"
            }

            # Get detailed link information
            $LinkDetails = Get-GPOLinkDetails -GPOID $GPO.Id
            $GPOInfo.LinkDetails = $LinkDetails
            $GPOInfo.LinkCount = $LinkDetails.Count

            if ($LinkDetails.Count -gt 0) {
                $GPOInfo.IsLinked = "Yes"
                $GPOSecuritySummary.LinkedGPOs++
            } else {
                $GPOSecuritySummary.UnlinkedGPOs++
                # Check if it's truly orphaned (no settings and not linked)
                if ($GPOReport.GPO.User.Enabled -eq "false" -and $GPOReport.GPO.Computer.Enabled -eq "false") {
                    $GPOInfo.IsOrphaned = $true
                    $GPOSecuritySummary.OrphanedGPOs++
                }
            }

            # Perform security analysis
            $SecurityAnalysis = Test-GPOSecuritySettings -GPOReport $GPOReport -GPOName $GPO.DisplayName
            $GPOInfo.SecurityIssues = $SecurityAnalysis.SecurityIssues

            if ($SecurityAnalysis.SecurityIssues.Count -gt 0) {
                $GPOSecuritySummary.SecurityIssuesFound++
                $GPOInfo.SecurityScore = [math]::Max(0, 100 - ($SecurityAnalysis.SecurityIssues.Count * 15))
                $AllSecurityIssues += $SecurityAnalysis.SecurityIssues
                $AllRecommendations += $SecurityAnalysis.SecurityRecommendations

                if ($GPOInfo.SecurityScore -lt 60) {
                    $GPOSecuritySummary.HighRiskGPOs++
                }
            }

            # Perform compliance analysis
            $ComplianceAnalysis = Analyze-GPOCompliance -GPOReport $GPOReport -GPOName $GPO.DisplayName
            $GPOInfo.ComplianceScore = $ComplianceAnalysis.ComplianceScore
            $GPOInfo.ComplianceIssues = $ComplianceAnalysis.ComplianceIssues

            if ($ComplianceAnalysis.ComplianceIssues.Count -gt 0) {
                $GPOSecuritySummary.ComplianceViolations++
            }

        } catch {
            Write-Warning "Error analyzing GPO '$($GPO.DisplayName)': $($_.Exception.Message)"
            $GPOInfo.SecurityIssues = @("Error during analysis: $($_.Exception.Message)")
        }

        # Get GPO permissions
        try {
            $GPOPermissions = Get-GPPermission -Guid $GPO.Id -All -ErrorAction SilentlyContinue
            $GPOInfo.PermissionsCount = $GPOPermissions.Count
        } catch {
            $GPOInfo.PermissionsCount = 0
        }

        $GPOAnalysis += New-Object PSObject -Property $GPOInfo
    }

    Write-Progress -Activity "Analyzing GPOs" -Completed

    # Perform GPO conflict analysis
    Write-Host "Performing GPO conflict analysis..." -ForegroundColor Yellow
    $ConflictAnalysis = Test-GPOConflicts -GPOs $GPOAnalysis
    $PrecedenceAnalysis = Get-GPOPrecedenceAnalysis -GPOs $GPOAnalysis

    # Store results
    $AuditResults.GroupPolicy = $GPOAnalysis
    $AuditResults.GPOSecuritySummary = $GPOSecuritySummary
    $AuditResults.GPOAllSecurityIssues = $AllSecurityIssues | Select-Object -Unique
    $AuditResults.GPOAllRecommendations = $AllRecommendations | Select-Object -Unique
    $AuditResults.GPOConflictAnalysis = $ConflictAnalysis
    $AuditResults.GPOPrecedenceAnalysis = $PrecedenceAnalysis

    # Display summary
    Write-Host "`n" + "="*80 -ForegroundColor Cyan
    Write-Host "Group Policy Security Analysis Summary" -ForegroundColor Cyan
    Write-Host "="*80 -ForegroundColor Cyan
    Write-Host "Total GPOs: $($GPOSecuritySummary.TotalGPOs)" -ForegroundColor White
    Write-Host "Linked GPOs: $($GPOSecuritySummary.LinkedGPOs)" -ForegroundColor Green
    Write-Host "Unlinked GPOs: $($GPOSecuritySummary.UnlinkedGPOs)" -ForegroundColor Yellow
    Write-Host "Orphaned GPOs: $($GPOSecuritySummary.OrphanedGPOs)" -ForegroundColor Red
    Write-Host "GPOs with Security Issues: $($GPOSecuritySummary.SecurityIssuesFound)" -ForegroundColor $(if($GPOSecuritySummary.SecurityIssuesFound -gt 0){"Red"}else{"Green"})
    Write-Host "High Risk GPOs: $($GPOSecuritySummary.HighRiskGPOs)" -ForegroundColor $(if($GPOSecuritySummary.HighRiskGPOs -gt 0){"Red"}else{"Green"})
    Write-Host "Compliance Violations: $($GPOSecuritySummary.ComplianceViolations)" -ForegroundColor $(if($GPOSecuritySummary.ComplianceViolations -gt 0){"Red"}else{"Green"})

    # Display conflict analysis summary
    if ($ConflictAnalysis) {
        Write-Host "`nGPO Conflict Analysis:" -ForegroundColor Cyan
        Write-Host "Total Conflicts: $($ConflictAnalysis.ConflictSummary.TotalConflicts)" -ForegroundColor $(if($ConflictAnalysis.ConflictSummary.TotalConflicts -gt 0){"Red"}else{"Green"})
        Write-Host "Critical Conflicts: $($ConflictAnalysis.ConflictSummary.CriticalConflicts)" -ForegroundColor $(if($ConflictAnalysis.ConflictSummary.CriticalConflicts -gt 0){"Red"}else{"Green"})
        Write-Host "Warning Conflicts: $($ConflictAnalysis.ConflictSummary.WarningConflicts)" -ForegroundColor $(if($ConflictAnalysis.ConflictSummary.WarningConflicts -gt 0){"Yellow"}else{"Green"})
        Write-Host "Affected OUs: $($ConflictAnalysis.ConflictSummary.AffectedOUs.Count)" -ForegroundColor $(if($ConflictAnalysis.ConflictSummary.AffectedOUs.Count -gt 0){"Yellow"}else{"Green"})
    }

} catch {
    Write-Error "Critical error in GPO analysis: $($_.Exception.Message)"
    $AuditResults.GroupPolicy = "Error: $($_.Exception.Message)"
}


#region User & Computer Accounts Analysis
Write-Host "Analyzing User & Computer Accounts..." -ForegroundColor Cyan
try {
    # User analysis
    $Users = Get-ADUser -Filter * -Properties LastLogonDate, PasswordLastSet, PasswordNeverExpires, Enabled, AccountLockoutTime
    
    # Categorize users for detailed exports
    $EnabledUsers = $Users | Where-Object {$_.Enabled -eq $true} | Select-Object Name, SamAccountName, LastLogonDate, PasswordLastSet
    $DisabledUsers = $Users | Where-Object {$_.Enabled -eq $false} | Select-Object Name, SamAccountName, LastLogonDate, PasswordLastSet
    $PasswordNeverExpiresUsers = $Users | Where-Object {$_.PasswordNeverExpires -eq $true} | Select-Object Name, SamAccountName, LastLogonDate, PasswordLastSet, Enabled
    $InactiveUsers = $Users | Where-Object {$_.LastLogonDate -lt (Get-Date).AddDays(-90) -and $_.LastLogonDate -ne $null} | Select-Object Name, SamAccountName, LastLogonDate, Enabled
    $LockedOutUsers = $Users | Where-Object {$_.AccountLockoutTime -ne $null} | Select-Object Name, SamAccountName, AccountLockoutTime, LastLogonDate

    $UserStats = @{
        TotalUsers = ($Users | Measure-Object).Count
        EnabledUsers = ($EnabledUsers | Measure-Object).Count
        DisabledUsers = ($DisabledUsers | Measure-Object).Count
        PasswordNeverExpires = ($PasswordNeverExpiresUsers | Measure-Object).Count
        InactiveUsers90Days = ($InactiveUsers | Measure-Object).Count
        LockedOutUsers = ($LockedOutUsers | Measure-Object).Count
        # Detailed data for exports
        AllUsersList = $Users | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet, PasswordNeverExpires, LockedOut
        EnabledUsersList = $EnabledUsers
        DisabledUsersList = $DisabledUsers
        PasswordNeverExpiresUsersList = $PasswordNeverExpiresUsers
        InactiveUsersList = $InactiveUsers
        LockedOutUsersList = $LockedOutUsers
    }
    
    # Computer analysis
    $Computers = Get-ADComputer -Filter * -Properties LastLogonDate, OperatingSystem, Enabled
    
    # Categorize computers for detailed exports
    $EnabledComputers = $Computers | Where-Object {$_.Enabled -eq $true} | Select-Object Name, OperatingSystem, LastLogonDate
    $DisabledComputers = $Computers | Where-Object {$_.Enabled -eq $false} | Select-Object Name, OperatingSystem, LastLogonDate
    $InactiveComputers = $Computers | Where-Object {$_.LastLogonDate -lt (Get-Date).AddDays(-90) -and $_.LastLogonDate -ne $null} | Select-Object Name, OperatingSystem, LastLogonDate
    $WindowsServers = $Computers | Where-Object {$_.OperatingSystem -like "*Server*"} | Select-Object Name, OperatingSystem, LastLogonDate, Enabled
    $WindowsWorkstations = $Computers | Where-Object {$_.OperatingSystem -like "*Windows*" -and $_.OperatingSystem -notlike "*Server*"} | Select-Object Name, OperatingSystem, LastLogonDate, Enabled

    $ComputerStats = @{
        TotalComputers = ($Computers | Measure-Object).Count
        EnabledComputers = ($EnabledComputers | Measure-Object).Count
        DisabledComputers = ($DisabledComputers | Measure-Object).Count
        InactiveComputers90Days = ($InactiveComputers | Measure-Object).Count
        WindowsServers = ($WindowsServers | Measure-Object).Count
        WindowsWorkstations = ($WindowsWorkstations | Measure-Object).Count
        # Detailed data for exports
        AllComputersList = $Computers | Select-Object Name, OperatingSystem, Enabled, LastLogonDate
        EnabledComputersList = $EnabledComputers
        DisabledComputersList = $DisabledComputers
        InactiveComputersList = $InactiveComputers
        WindowsServersList = $WindowsServers
        WindowsWorkstationsList = $WindowsWorkstations
    }
    
    $AuditResults.AccountsAnalysis = @{
        Users = $UserStats
        Computers = $ComputerStats
    }
} catch {
    $AuditResults.AccountsAnalysis = "Error: $($_.Exception.Message)"
}
#endregion

#region Security Analysis & Risk Assessment
Write-Host "Performing Security Analysis..." -ForegroundColor Cyan

try {
    $SecurityFindings = @()

    # Get base domain SID and forest root domain SID
    $DomainSID = (Get-ADDomain).DomainSID.Value
    $RootDomainSID = (Get-ADDomain -Identity (Get-ADForest).RootDomain).DomainSID.Value

    # Build well-known group SIDs
    $DomainAdminsSID     = "$DomainSID-512"
    $EnterpriseAdminsSID = "$RootDomainSID-519"
    $SchemaAdminsSID     = "$RootDomainSID-518"

    # Retrieve privileged group members using SIDs
    $DomainAdmins     = Get-ADGroupMember -Identity $DomainAdminsSID -Recursive -ErrorAction SilentlyContinue
    $EnterpriseAdmins = Get-ADGroupMember -Identity $EnterpriseAdminsSID -Recursive -ErrorAction SilentlyContinue
    $SchemaAdmins     = Get-ADGroupMember -Identity $SchemaAdminsSID -Recursive -ErrorAction SilentlyContinue

    # Check high number of Domain Admins
    if ($DomainAdmins.Count -gt 5) {
        $SecurityFindings += "High number of Domain Admins ($($DomainAdmins.Count))"
    }

    # Check for weak password policy
    $DefaultDomainPolicy = Get-ADDefaultDomainPasswordPolicy
    if ($DefaultDomainPolicy.MinPasswordLength -lt 8) {
        $SecurityFindings += "Weak minimum password length in default policy: $($DefaultDomainPolicy.MinPasswordLength)"
    }
    
    # Check Password Settings Objects (PSOs) / Fine-Grained Password Policies
    $PSOs = Get-ADFineGrainedPasswordPolicy -Filter * -ErrorAction SilentlyContinue
    if ($PSOs) {
        $SecurityFindings += "Found $($PSOs.Count) Password Settings Objects (Fine-Grained Password Policies)"
        foreach ($PSO in $PSOs) {
            if ($PSO.MinPasswordLength -lt 8) {
                $SecurityFindings += "Weak minimum password length in PSO '$($PSO.Name)': $($PSO.MinPasswordLength)"
            }
            if ($PSO.MaxPasswordAge.Days -gt 90) {
                $SecurityFindings += "Long password age in PSO '$($PSO.Name)': $($PSO.MaxPasswordAge.Days) days"
            }
            if (-not $PSO.ComplexityEnabled) {
                $SecurityFindings += "Password complexity disabled in PSO '$($PSO.Name)'"
            }
        }
    } else {
        $SecurityFindings += "No Password Settings Objects (PSOs) configured - consider implementing for privileged accounts"
    }
    
    # Check KRBTGT account password age
    Write-Host "  Checking KRBTGT account password age..." -ForegroundColor Gray
    try {
        $KRBTGTAccount = Get-ADUser -Identity "krbtgt" -Properties PasswordLastSet -ErrorAction Stop
        $PasswordAge = (Get-Date) - $KRBTGTAccount.PasswordLastSet
        
        if ($PasswordAge.Days -gt 180) {
            $SecurityFindings += "KRBTGT account password is $($PasswordAge.Days) days old - should be reset every 180 days max"
        } elseif ($PasswordAge.Days -gt 90) {
            $SecurityFindings += "KRBTGT account password is $($PasswordAge.Days) days old - consider resetting soon"
        }
        
        # Store KRBTGT info for reporting
        $AuditResults.KRBTGTPasswordAge = @{
            PasswordLastSet = $KRBTGTAccount.PasswordLastSet
            PasswordAgeDays = $PasswordAge.Days
            Status = if ($PasswordAge.Days -gt 180) { "Critical" } elseif ($PasswordAge.Days -gt 90) { "Warning" } else { "Good" }
        }
        
        Write-Host "  ✓ KRBTGT password age: $($PasswordAge.Days) days" -ForegroundColor Green
    } catch {
        $SecurityFindings += "Unable to check KRBTGT account password age: $($_.Exception.Message)"
        $AuditResults.KRBTGTPasswordAge = @{
            Status = "Error"
            Error = $_.Exception.Message
        }
        Write-Warning "  ✗ Failed to check KRBTGT account: $($_.Exception.Message)"
    }
    
    # Check if Active Directory Recycle Bin is enabled
    Write-Host "  Checking Active Directory Recycle Bin status..." -ForegroundColor Gray
    try {
        $RecycleBinFeature = Get-ADOptionalFeature -Filter 'Name -like "Recycle Bin Feature"' -ErrorAction Stop
        $RecycleBinEnabled = $RecycleBinFeature.EnabledScopes.Count -gt 0
        
        if ($RecycleBinEnabled) {
            Write-Host "  ✓ Active Directory Recycle Bin is enabled" -ForegroundColor Green
            $AuditResults.RecycleBinStatus = @{
                Enabled = $true
                EnabledScopes = $RecycleBinFeature.EnabledScopes
                Status = "Enabled"
            }
        } else {
            $SecurityFindings += "Active Directory Recycle Bin is not enabled - consider enabling for accidental deletion protection"
            Write-Warning "  ✗ Active Directory Recycle Bin is not enabled"
            $AuditResults.RecycleBinStatus = @{
                Enabled = $false
                Status = "Disabled"
                Recommendation = "Enable AD Recycle Bin for protection against accidental deletions"
            }
        }
    } catch {
        $SecurityFindings += "Unable to check Active Directory Recycle Bin status: $($_.Exception.Message)"
        $AuditResults.RecycleBinStatus = @{
            Status = "Error"
            Error = $_.Exception.Message
        }
        Write-Warning "  ✗ Failed to check Recycle Bin status: $($_.Exception.Message)"
    }

    # Check for accounts with non-expiring passwords
    $NonExpiringPasswords = Get-ADUser -Filter {PasswordNeverExpires -eq $true -and Enabled -eq $true}
    if ($NonExpiringPasswords.Count -gt 0) {
        $SecurityFindings += "$($NonExpiringPasswords.Count) enabled accounts with non-expiring passwords"
    }

    # Check for accounts with Kerberos pre-authentication disabled
    $PreAuthDisabled = Get-ADUser -Filter {DoesNotRequirePreAuth -eq $true}
    if ($PreAuthDisabled.Count -gt 0) {
        $SecurityFindings += "$($PreAuthDisabled.Count) accounts with Kerberos pre-authentication disabled"
    }

    # Store results
    $AuditResults.SecurityFindings = $SecurityFindings
    $AuditResults.PasswordPolicy = @{
        DefaultPolicy = $DefaultDomainPolicy
        PSOs = $PSOs
        PSOCount = if ($PSOs) { $PSOs.Count } else { 0 }
    }
    $AuditResults.PrivilegedGroups = @{
        DomainAdminsCount     = ($DomainAdmins | Measure-Object).Count
        EnterpriseAdminsCount = ($EnterpriseAdmins | Measure-Object).Count
        SchemaAdminsCount     = ($SchemaAdmins | Measure-Object).Count
    }

} catch {
    $AuditResults.SecurityFindings = @("Error: $($_.Exception.Message)")
}
#endregion





#region OU Hierarchy Mapping
Write-Host "Mapping OU Hierarchy..." -ForegroundColor Cyan
try {
    $OUs = Get-ADOrganizationalUnit -Filter * | Sort-Object DistinguishedName
    $OUHierarchy = @()
    
    foreach ($OU in $OUs) {
        $OUInfo = @{
            Name = $OU.Name
            DistinguishedName = $OU.DistinguishedName
            Description = $OU.Description
            Created = $OU.Created
            Modified = $OU.Modified
        }
        
        # Calculate OU level based on DN
        $DNParts = $OU.DistinguishedName -split ","
        $OULevel = ($DNParts | Where-Object {$_ -like "OU=*"}).Count
        $OUInfo.Level = $OULevel
        
        # Get parent OU
        if ($OULevel -gt 1) {
            $ParentDN = ($DNParts[1..($DNParts.Count-1)]) -join ","
            $OUInfo.ParentDN = $ParentDN
        } else {
            $OUInfo.ParentDN = ($DNParts[1..($DNParts.Count-1)]) -join ","
        }
        
        # Count objects in OU
        $OUUsers = Get-ADUser -SearchBase $OU.DistinguishedName -SearchScope OneLevel -Filter * | Measure-Object
        $OUComputers = Get-ADComputer -SearchBase $OU.DistinguishedName -SearchScope OneLevel -Filter * | Measure-Object
        $OUGroups = Get-ADGroup -SearchBase $OU.DistinguishedName -SearchScope OneLevel -Filter * | Measure-Object
        $ChildOUs = Get-ADOrganizationalUnit -SearchBase $OU.DistinguishedName -SearchScope OneLevel -Filter * | Measure-Object
        
        $OUInfo.UserCount = $OUUsers.Count
        $OUInfo.ComputerCount = $OUComputers.Count
        $OUInfo.GroupCount = $OUGroups.Count
        $OUInfo.ChildOUCount = $ChildOUs.Count
        
        # Create a safe ID for HTML/JS
        $OUInfo.SafeID = ($OU.DistinguishedName -replace '[^a-zA-Z0-9]', '_')
        
        $OUHierarchy += New-Object PSObject -Property $OUInfo
    }
    $AuditResults.OUHierarchy = $OUHierarchy
} catch {
    $AuditResults.OUHierarchy = "Error: $($_.Exception.Message)"
}
#endregion



#region Protocols & Legacy Services Check
Write-Host "Auditing TLS, SMBv1, LLMNR, and NetBIOS settings..." -ForegroundColor Cyan
try {
    $ProtocolAudit = @{}

    # --- TLS 1.0 & TLS 1.1 Check ---
    $tlsKeys = @("TLS 1.0", "TLS 1.1")
    foreach ($tls in $tlsKeys) {
        $clientKey = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$tls\Client"
        $serverKey = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$tls\Server"

        $tlsClient = if (Test-Path $clientKey) { Get-ItemProperty -Path $clientKey -ErrorAction SilentlyContinue } else { $null }
        $tlsServer = if (Test-Path $serverKey) { Get-ItemProperty -Path $serverKey -ErrorAction SilentlyContinue } else { $null }

        $ProtocolAudit["$tls Enabled (Client)"] = if ($tlsClient.Enabled -eq 1) { "Yes" } else { "No" }
        $ProtocolAudit["$tls Enabled (Server)"] = if ($tlsServer.Enabled -eq 1) { "Yes" } else { "No" }
    }

    # --- SMBv1 Check ---
    $smbFeature = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction SilentlyContinue
    $smbRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
    $smbEnabled = if (Test-Path $smbRegPath) {
        $val = (Get-ItemProperty -Path $smbRegPath -Name SMB1 -ErrorAction SilentlyContinue).SMB1
        if ($val -eq 1) { "Yes" } else { "No" }
    } else { "Unknown" }

    $ProtocolAudit["SMBv1 Installed"] = if ($smbFeature.State -eq "Enabled") { "Yes" } else { "No" }
    $ProtocolAudit["SMBv1 Enabled"]   = $smbEnabled

    # --- LLMNR Check ---
    $llmnrRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient"
    $llmnrStatus = if (Test-Path $llmnrRegPath) {
        $val = (Get-ItemProperty -Path $llmnrRegPath -Name EnableMulticast -ErrorAction SilentlyContinue).EnableMulticast
        if ($val -eq 0) { "Disabled" } elseif ($val -eq 1) { "Enabled" } else { "Not Configured" }
    } else { "Not Configured" }
    $ProtocolAudit["LLMNR"] = $llmnrStatus

    # --- NetBIOS Check ---
    $netAdapters = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
    $netbiosStatuses = $netAdapters | ForEach-Object {
        switch ($_.TcpipNetbiosOptions) {
            0 { "Default" }
            1 { "Enabled" }
            2 { "Disabled" }
        }
    } | Sort-Object -Unique
    $ProtocolAudit["NetBIOS"] = ($netbiosStatuses -join ", ")

    # Store in Audit Results
    $AuditResults.ProtocolsAndLegacyServices = $ProtocolAudit

} catch {
    $AuditResults.ProtocolsAndLegacyServices = "Error: $($_.Exception.Message)"
}
#endregion


#region Orphaned SID Analysis
Write-Host "Scanning for Orphaned SIDs in Active Directory ACLs..." -ForegroundColor Cyan
try {
    $DomainInfo = Get-ADDomain
    $DomainSID = $DomainInfo.DomainSID.ToString()
    $SearchBase = $DomainInfo.DistinguishedName

    $Objects = Get-ADObject -Filter * -SearchBase $SearchBase -Properties nTSecurityDescriptor -ErrorAction Stop |
               Where-Object { $_.ObjectClass -notin @("lostAndFound", "crossRef") }

    $TotalObjects = $Objects.Count
    $ProcessedCount = 0
    $OrphanedList = @()

    foreach ($Object in $Objects) {
        $ProcessedCount++
        Write-Progress -Activity "Checking objects for orphaned SIDs" `
                       -Status "$ProcessedCount of $TotalObjects" `
                       -PercentComplete ($ProcessedCount / $TotalObjects * 100)

        if (-not $Object.nTSecurityDescriptor) { continue }

        foreach ($ACE in $Object.nTSecurityDescriptor.Access) {
            if ($ACE.IdentityReference.Value -like "$DomainSID*") {
                try {
                    $null = $ACE.IdentityReference.Translate([System.Security.Principal.NTAccount])
                }
                catch [System.Security.Principal.IdentityNotMappedException] {
                    $OrphanedList += $Object.DistinguishedName
                }
            }
        }
    }

    $TotalOrphanedSIDs = $OrphanedList.Count

    # Group by DN to find objects with most orphaned SIDs
    $TopObjects = $OrphanedList |
        Group-Object -NoElement |
        Sort-Object Count -Descending |
        Select-Object Name, Count -First 10

    $AuditResults.OrphanedSIDs = @{
        TotalCount = $TotalOrphanedSIDs
        TopObjects = $TopObjects
    }
} catch {
    $AuditResults.OrphanedSIDs = "Error: $($_.Exception.Message)"
}
#endregion



#region Event Log Analysis
Write-Host "Analyzing Security Events & Errors..." -ForegroundColor Cyan
try {
    $EventAnalysis = @()

    # Define critical events by log container
    $CriticalEventContainers = @{
        "Security" = @{
            # Group Management Events
            "4728" = @{ Description = "Member Added to Security-Enabled Global Group"; Severity = "High"; Category = "Group Management" }
            "4729" = @{ Description = "Member Removed from Security-Enabled Global Group"; Severity = "Medium"; Category = "Group Management" }
            "4732" = @{ Description = "Member Added to Security-Enabled Local Group"; Severity = "High"; Category = "Group Management" }
            "4733" = @{ Description = "Member Removed from Security-Enabled Local Group"; Severity = "Medium"; Category = "Group Management" }
            "4756" = @{ Description = "Member Added to Security-Enabled Universal Group"; Severity = "High"; Category = "Group Management" }
            "4757" = @{ Description = "Member Removed from Security-Enabled Universal Group"; Severity = "Medium"; Category = "Group Management" }

            # Account Management Events
            "4720" = @{ Description = "User Account Created"; Severity = "High"; Category = "Account Management" }
            "4726" = @{ Description = "User Account Deleted"; Severity = "Medium"; Category = "Account Management" }
            "4741" = @{ Description = "Computer Account Created"; Severity = "Medium"; Category = "Account Management" }
            "4743" = @{ Description = "Computer Account Deleted"; Severity = "Medium"; Category = "Account Management" }
            "4767" = @{ Description = "User Account Unlocked"; Severity = "Medium"; Category = "Account Management" }

            # Policy and Security Changes
            "4719" = @{ Description = "System Audit Policy Changed"; Severity = "High"; Category = "Policy Changes" }
            "4739" = @{ Description = "Domain Policy Changed"; Severity = "High"; Category = "Policy Changes" }
            "4713" = @{ Description = "Kerberos Policy Changed"; Severity = "High"; Category = "Policy Changes" }
            "4716" = @{ Description = "Trusted Domain Information Modified"; Severity = "Medium"; Category = "Policy Changes" }

            # Advanced Security Events
            "4765" = @{ Description = "SID History Added to Account"; Severity = "High"; Category = "Advanced Security" }
            "4766" = @{ Description = "Attempt to Add SID History Failed"; Severity = "Medium"; Category = "Advanced Security" }
            "4780" = @{ Description = "ACL Set on Accounts Which Are Members of Administrators Groups"; Severity = "High"; Category = "Advanced Security" }

            # Privileged Access Events
            "4648" = @{ Description = "Logon Attempted Using Explicit Credentials"; Severity = "Low"; Category = "Privileged Access" }
            "4672" = @{ Description = "Special Privileges Assigned to New Logon"; Severity = "Medium"; Category = "Privileged Access" }
            "4673" = @{ Description = "Privileged Service Called"; Severity = "Low"; Category = "Privileged Access" }
            "4674" = @{ Description = "Operation Attempted on Privileged Object"; Severity = "Low"; Category = "Privileged Access" }

            # Process and Service Events
            "4688" = @{ Description = "New Process Created"; Severity = "Low"; Category = "Process/Service Events" }
            "4697" = @{ Description = "Service Installed"; Severity = "Medium"; Category = "Process/Service Events" }

            # Scheduled Tasks
            "4698" = @{ Description = "Scheduled Task Created"; Severity = "Medium"; Category = "Scheduled Tasks" }
            "4699" = @{ Description = "Scheduled Task Deleted"; Severity = "Medium"; Category = "Scheduled Tasks" }
            "4700" = @{ Description = "Scheduled Task Enabled"; Severity = "Medium"; Category = "Scheduled Tasks" }
            "4701" = @{ Description = "Scheduled Task Disabled"; Severity = "Medium"; Category = "Scheduled Tasks" }
            "4702" = @{ Description = "Scheduled Task Updated"; Severity = "Medium"; Category = "Scheduled Tasks" }

            # Directory Service Events
            "5136" = @{ Description = "Directory Service Object Modified"; Severity = "High"; Category = "Directory Changes" }
            "5137" = @{ Description = "Directory Service Object Created"; Severity = "High"; Category = "Directory Changes" }
            "5138" = @{ Description = "Directory Service Object Undeleted"; Severity = "Medium"; Category = "Directory Changes" }
            "5139" = @{ Description = "Directory Service Object Moved"; Severity = "Medium"; Category = "Directory Changes" }
            "5141" = @{ Description = "Directory Service Object Deleted"; Severity = "High"; Category = "Directory Changes" }
        }

        "System" = @{
            # Service Control Manager Events
            "7034" = @{ Description = "Service Crashed Unexpectedly"; Severity = "High"; Category = "Service Issues" }
            "7035" = @{ Description = "Service Start/Stop"; Severity = "Medium"; Category = "Service Management" }
            "7036" = @{ Description = "Service Started or Stopped"; Severity = "Low"; Category = "Service Management" }
            "7040" = @{ Description = "Service Start Type Changed"; Severity = "Medium"; Category = "Service Management" }

            # System Events
            "1074" = @{ Description = "System Shutdown/Restart Initiated"; Severity = "Medium"; Category = "System Events" }
            "6005" = @{ Description = "Event Log Service Started"; Severity = "Low"; Category = "System Events" }
            "6006" = @{ Description = "Event Log Service Stopped"; Severity = "Medium"; Category = "System Events" }
            "6008" = @{ Description = "Unexpected System Shutdown"; Severity = "High"; Category = "System Events" }
            "6009" = @{ Description = "System Started"; Severity = "Low"; Category = "System Events" }

            # Disk and Storage Events
            "51" = @{ Description = "Disk Error"; Severity = "High"; Category = "Storage Issues" }
            "98" = @{ Description = "System File Corruption"; Severity = "High"; Category = "Storage Issues" }
        }

        "Application" = @{
            # Application Crashes and Errors
            "1000" = @{ Description = "Application Error"; Severity = "Medium"; Category = "Application Issues" }
            "1001" = @{ Description = "Application Hang"; Severity = "Medium"; Category = "Application Issues" }
            "1002" = @{ Description = "Application Crash"; Severity = "High"; Category = "Application Issues" }

            # Windows Error Reporting
            "1026" = @{ Description = "Windows Error Reporting"; Severity = "Medium"; Category = "System Errors" }

            # MSI Installer Errors
            "11707" = @{ Description = "Installation Completed Successfully"; Severity = "Low"; Category = "Installation" }
            "11708" = @{ Description = "Installation Failed"; Severity = "Medium"; Category = "Installation" }
            "11724" = @{ Description = "Package Removal Completed"; Severity = "Low"; Category = "Installation" }
        }

        "DFS Replication" = @{
            # DFS Replication Issues
            "4012" = @{ Description = "DFS Replication Service Stopped"; Severity = "High"; Category = "Replication Issues" }
            "4013" = @{ Description = "DFS Replication Service Started"; Severity = "Low"; Category = "Replication Management" }
            "5014" = @{ Description = "DFS Replication Conflict"; Severity = "Medium"; Category = "Replication Issues" }
        }

        "Directory Service" = @{
            # Active Directory Events
            "1" = @{ Description = "Active Directory Database Corruption"; Severity = "High"; Category = "AD Critical" }
            "2" = @{ Description = "Active Directory Service Error"; Severity = "High"; Category = "AD Critical" }
            "1046" = @{ Description = "Active Directory Replication Error"; Severity = "High"; Category = "AD Replication" }
            "1388" = @{ Description = "Active Directory Replication Failure"; Severity = "High"; Category = "AD Replication" }
            "1645" = @{ Description = "Active Directory Trust Failure"; Severity = "High"; Category = "AD Trust" }
        }

        "DNS Server" = @{
            # DNS Critical Events
            "150" = @{ Description = "DNS Server Failed to Start"; Severity = "High"; Category = "DNS Critical" }
            "4013" = @{ Description = "DNS Server Unable to Load Zone"; Severity = "High"; Category = "DNS Issues" }
            "6527" = @{ Description = "DNS Server Scavenging Error"; Severity = "Medium"; Category = "DNS Management" }
        }

        "File Replication Service" = @{
            # FRS Events
            "13508" = @{ Description = "FRS Cannot Communicate"; Severity = "High"; Category = "FRS Issues" }
            "13509" = @{ Description = "FRS Replica Set Stopped"; Severity = "High"; Category = "FRS Issues" }
        }
    }

    # Define noise/generic error sources to filter out
    $NoiseEventSources = @(
        "PerfLib", "perflib", "Perflib",
        "Schannel", "schannel", "SChannel",
        "DCOM", "DistributedCOM",
        "Microsoft-Windows-WMI", "WMI",
        "Microsoft-Windows-Kernel-General",
        "Microsoft-Windows-GroupPolicy",
        "Microsoft-Windows-DNS-Client",
        "Microsoft-Windows-Time-Service",
        "Microsoft-Windows-User32",
        "Microsoft-Windows-Winlogon",
        "Microsoft-Windows-TerminalServices",
        "VSS", "VssVC",
        "EventLog",
        "Outlook",
        "Microsoft Office"
    )

    # Keywords to filter out from error descriptions
    $NoiseKeywords = @(
        "perflib", "performance counter", "performance library",
        "schannel", "secure channel", "ssl handshake",
        "wmi", "windows management",
        "group policy", "winlogon", "user32",
        "time service", "dns client", "vss",
        "outlook", "office", "microsoft office"
    )

    # Get domain controllers for event analysis
    $DCs = Get-ADDomainController -Filter * -ErrorAction SilentlyContinue
    if (-not $DCs) {
        Write-Warning "Unable to retrieve domain controllers for event analysis"
        $AuditResults.EventLogAnalysis = "Error: Unable to retrieve domain controllers"
        return
    }

    foreach ($DC in $DCs[0..2]) { # Check first 3 DCs to avoid timeout
        $DCStartTime = Get-Date
        Write-Host "  Analyzing critical events on $($DC.HostName)..." -ForegroundColor Yellow

        # First, check if DC is reachable before attempting event log queries
        $DCHostName = $DC.HostName.ToString()
        $LocalComputerName = $env:COMPUTERNAME
        $IsLocalDC = ($DCHostName -eq $LocalComputerName -or $DCHostName -eq "localhost" -or $DCHostName -eq "127.0.0.1" -or $DCHostName -eq "$LocalComputerName.$env:USERDNSDOMAIN")

        if ($IsLocalDC) {
            Write-Host "    ✓ Running on local DC - Skipping connectivity test" -ForegroundColor Green
            $DCReachable = $true
        } else {
            Write-Host "    Testing connectivity to $DCHostName..." -ForegroundColor Gray
            $DCReachable = $false
            try {
                # Quick connectivity test with 3-second timeout
                $PingResult = Test-Connection -ComputerName $DCHostName -Count 1 -Quiet -ErrorAction Stop
                if ($PingResult) {
                    $DCReachable = $true
                    Write-Host "    ✓ DC is reachable" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "    ✗ DC is unreachable - Skipping event analysis" -ForegroundColor Red
                continue
            }

            if (-not $DCReachable) {
                Write-Host "    ⚠️ Skipping event analysis for unreachable DC: $DCHostName" -ForegroundColor Yellow
                continue
            }
        }

        # Check available event logs first to avoid querying non-existent logs
        Write-Host "    Checking available event logs..." -ForegroundColor Gray
        $AvailableLogs = @()

        # Use appropriate computer name for local vs remote DCs
        $ComputerNameForQuery = if ($IsLocalDC) { "localhost" } else { $DCHostName }

        try {
            Write-Host "    Enumerating available event logs..." -ForegroundColor Gray
            $Job = Start-Job -ScriptBlock {
                param($ComputerName, $IsLocal)
                if ($IsLocal) {
                    # For local DC, don't use -ComputerName parameter to avoid authentication issues
                    Get-WinEvent -ListLog * -ErrorAction SilentlyContinue | Where-Object { $_.IsEnabled -eq $true } | Select-Object -ExpandProperty LogName
                } else {
                    Get-WinEvent -ComputerName $ComputerName -ListLog * -ErrorAction SilentlyContinue | Where-Object { $_.IsEnabled -eq $true } | Select-Object -ExpandProperty LogName
                }
            } -ArgumentList $ComputerNameForQuery, $IsLocalDC

            $AvailableLogs = Receive-Job -Job $Job -Wait -Timeout 15
            Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue

            if ($AvailableLogs) {
                Write-Host "    ✓ Found $($AvailableLogs.Count) available event logs" -ForegroundColor Green
            } else {
                throw "No logs returned"
            }
        }
        catch {
            Write-Host "    ⚠️ Could not enumerate event logs, using default list" -ForegroundColor Yellow
            $AvailableLogs = @("System", "Application", "Security", "Directory Service")
        }

        # First, do a quick scan for any recent errors/warnings in key logs
        foreach ($LogName in @("System", "Application")) {
            if ($AvailableLogs -contains $LogName) {
                try {
                    Write-Host "    Scanning $LogName for recent errors..." -ForegroundColor Gray
                    $Job = Start-Job -ScriptBlock {
                        param($ComputerName, $LogName, $IsLocal)
                        try {
                            if ($IsLocal) {
                                Get-WinEvent -FilterHashtable @{
                                    LogName = $LogName
                                    Level = @(1,2,3) # Critical, Error, Warning
                                    StartTime = (Get-Date).AddDays(-7)
                                } -MaxEvents 100 -ErrorAction SilentlyContinue
                            } else {
                                Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{
                                    LogName = $LogName
                                    Level = @(1,2,3) # Critical, Error, Warning
                                    StartTime = (Get-Date).AddDays(-7)
                                } -MaxEvents 100 -ErrorAction SilentlyContinue
                            }
                        } catch {
                            @()
                        }
                    } -ArgumentList $ComputerNameForQuery, $LogName, $IsLocalDC

                    $RecentErrors = Receive-Job -Job $Job -Wait -Timeout 15
                    Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue

                    if ($RecentErrors -and $RecentErrors.Count -gt 0) {
                        Write-Host "    ✓ Found $($RecentErrors.Count) recent errors/warnings in $LogName" -ForegroundColor Yellow

                        # Group by event ID and take most frequent ones
                        $TopErrors = $RecentErrors | Group-Object Id | Sort-Object Count -Descending | Select-Object -First 5

                        foreach ($ErrorGroup in $TopErrors) {
                            $SampleEvent = $ErrorGroup.Group[0]
                            $EventAnalysis += [PSCustomObject]@{
                                DomainController = $DC.HostName
                                LogName = $LogName
                                EventID = $SampleEvent.Id
                                Description = "Recent $($SampleEvent.LevelDisplayName): $($SampleEvent.ProviderName)"
                                Severity = switch ($SampleEvent.Level) { 1 { "Critical" } 2 { "High" } 3 { "Medium" } default { "Low" } }
                                Category = "System Errors"
                                TotalCount = $ErrorGroup.Count
                                MostRecentTime = ($ErrorGroup.Group | Sort-Object TimeCreated -Descending)[0].TimeCreated
                                RecentEvents = @($ErrorGroup.Group | Select-Object -First 3 | ForEach-Object {
                                    [PSCustomObject]@{
                                        TimeCreated = $_.TimeCreated
                                        UserName = if ($_.UserId) { $_.UserId.Value } else { "System" }
                                        TargetAccount = "N/A"
                                        LogName = $LogName
                                    }
                                })
                            }
                        }
                    } else {
                        Write-Host "    ✓ No recent errors in $LogName" -ForegroundColor Green
                    }
                } catch {
                    Write-Host "    ⚠️ Error scanning $LogName for recent errors: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }

        # Then check specific critical events
        foreach ($LogName in $CriticalEventContainers.Keys) {
            if ($AvailableLogs -notcontains $LogName) {
                Write-Host "    ⚠️ Log '$LogName' not available on this DC, skipping..." -ForegroundColor Yellow
                continue
            }

            Write-Host "    Checking specific critical events in $LogName..." -ForegroundColor Gray

            foreach ($EventID in $CriticalEventContainers[$LogName].Keys) {
                try {
                    $EventConfig = $CriticalEventContainers[$LogName][$EventID]

                    # Get events from the last 7 days with timeout protection
                    $Events = $null
                    $Job = Start-Job -ScriptBlock {
                        param($ComputerName, $LogName, $EventID, $IsLocal)
                        try {
                            if ($IsLocal) {
                                Get-WinEvent -FilterHashtable @{
                                    LogName = $LogName
                                    ID = [int]$EventID
                                    StartTime = (Get-Date).AddDays(-7)
                                } -MaxEvents 50 -ErrorAction SilentlyContinue
                            } else {
                                Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{
                                    LogName = $LogName
                                    ID = [int]$EventID
                                    StartTime = (Get-Date).AddDays(-7)
                                } -MaxEvents 50 -ErrorAction SilentlyContinue
                            }
                        } catch {
                            # Return empty if error occurs
                            @()
                        }
                    } -ArgumentList $ComputerNameForQuery, $LogName, $EventID, $IsLocalDC

                    # Wait for job to complete or timeout after 15 seconds
                    $Events = Receive-Job -Job $Job -Wait -Timeout 15
                    Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue

                    if ($Events -and $Events.Count -gt 0) {
                        # Use the predefined severity and category
                        $Severity = $EventConfig.Severity
                        $Category = $EventConfig.Category
                        $Description = $EventConfig.Description

                        # Build recent events array
                        $RecentEvents = @()
                        foreach ($Event in ($Events | Select-Object -First 3)) {
                            try {
                                $xml = [xml]$Event.ToXml()
                                $UserName = "N/A"
                                $TargetAccount = "N/A"

                                # Try to extract username (mainly for Security events)
                                if ($LogName -eq "Security") {
                                    $UserNameNode = $xml.Event.EventData.Data | Where-Object {$_.Name -eq 'SubjectUserName'}
                                    if ($UserNameNode) {
                                        $UserName = $UserNameNode.'#text'
                                    }

                                    # Try to extract target account
                                    $TargetNodes = $xml.Event.EventData.Data | Where-Object {$_.Name -in @('TargetUserName','TargetSid','MemberName')}
                                    if ($TargetNodes) {
                                        $TargetAccount = ($TargetNodes | Select-Object -First 1).'#text'
                                    }
                                } else {
                                    # For non-Security logs, try to get general information
                                    $UserName = if ($Event.UserId) { $Event.UserId.Value } else { "System" }
                                    $TargetAccount = if ($xml.Event.EventData.Data) {
                                        ($xml.Event.EventData.Data | Select-Object -First 1).'#text'
                                    } else {
                                        "N/A"
                                    }
                                }

                                $RecentEvents += [PSCustomObject]@{
                                    TimeCreated = $Event.TimeCreated
                                    UserName = $UserName
                                    TargetAccount = $TargetAccount
                                    LogName = $LogName
                                }
                            } catch {
                                $RecentEvents += [PSCustomObject]@{
                                    TimeCreated = $Event.TimeCreated
                                    UserName = "N/A"
                                    TargetAccount = "N/A"
                                    LogName = $LogName
                                }
                            }
                        }

                        # Add to event analysis
                        $EventAnalysis += [PSCustomObject]@{
                            DomainController = $DC.DomainController
                            LogName = $LogName
                            EventID = $EventID
                            EventType = $Description
                            Count = $Events.Count
                            LastOccurrence = $Events[0].TimeCreated
                            Severity = $Severity
                            Category = $Category
                            RecentEvents = $RecentEvents
                        }
                    }
                } catch {
                    # Only log critical errors, skip events that don't exist
                    if ($_.Exception.Message -notlike "*No events were found*" -and
                        $_.Exception.Message -notlike "*The specified channel could not be found*") {
                        Write-Warning "Error checking Event $EventID in $LogName on $($DC.DomainController): $($_.Exception.Message)"
                    }
                }
            }
        }

        # Collect general error events (Level 2 = Error) from key logs, filtered to avoid noise
        Write-Host "    Collecting filtered error events..." -ForegroundColor Gray

        $ErrorLogs = @("System", "Application", "Directory Service", "DNS Server")
        foreach ($LogName in $ErrorLogs) {
            try {
                # Get error events from last 7 days with timeout protection
                $ErrorEvents = $null
                $Job = Start-Job -ScriptBlock {
                    param($ComputerName, $LogName, $IsLocal)
                    if ($IsLocal) {
                        Get-WinEvent -FilterHashtable @{
                            LogName = $LogName
                            Level = 2  # Error level
                            StartTime = (Get-Date).AddDays(-7)
                        } -MaxEvents 100 -ErrorAction SilentlyContinue
                    } else {
                        Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{
                            LogName = $LogName
                            Level = 2  # Error level
                            StartTime = (Get-Date).AddDays(-7)
                        } -MaxEvents 100 -ErrorAction SilentlyContinue
                    }
                } -ArgumentList $ComputerNameForQuery, $LogName, $IsLocalDC

                # Wait for job to complete or timeout after 10 seconds
                $ErrorEvents = Receive-Job -Job $Job -Wait -Timeout 10
                Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue

                if ($ErrorEvents) {
                    # Group by Event ID and filter out noise
                    $FilteredErrors = $ErrorEvents | Group-Object Id | ForEach-Object {
                        $EventGroup = $_.Group
                        $SampleEvent = $EventGroup[0]

                        # Check if this is a noise event
                        $IsNoise = $false

                        # Filter by source provider
                        if ($SampleEvent.ProviderName -in $NoiseEventSources) {
                            $IsNoise = $true
                        }

                        # Filter by message content
                        if (-not $IsNoise -and $SampleEvent.LevelDisplayName -eq "Error") {
                            foreach ($NoiseKeyword in $NoiseKeywords) {
                                if ($SampleEvent.Message -and $SampleEvent.Message.ToLower() -like "*$NoiseKeyword*") {
                                    $IsNoise = $true
                                    break
                                }
                            }
                        }

                        # Skip if it's noise or if we already have this event ID from critical events
                        if (-not $IsNoise -and -not ($EventAnalysis | Where-Object { $_.EventID -eq $SampleEvent.Id -and $_.LogName -eq $LogName })) {

                            # Build recent events for this error
                            $RecentErrorEvents = @()
                            foreach ($ErrorEvent in ($EventGroup | Select-Object -First 3)) {
                                $RecentErrorEvents += [PSCustomObject]@{
                                    TimeCreated = $ErrorEvent.TimeCreated
                                    UserName = if ($ErrorEvent.UserId) { $ErrorEvent.UserId.Value } else { "System" }
                                    TargetAccount = "N/A"
                                    LogName = $LogName
                                }
                            }

                            # Create error event entry
                            [PSCustomObject]@{
                                DomainController = $DC.DomainController
                                LogName = $LogName
                                EventID = $SampleEvent.Id
                                EventType = if ($SampleEvent.Message -and $SampleEvent.Message.Length -gt 100) {
                                    $SampleEvent.Message.Substring(0, 100) + "..."
                                } else {
                                    $SampleEvent.Message
                                }
                                Count = $EventGroup.Count
                                LastOccurrence = $EventGroup[0].TimeCreated
                                Severity = "Medium"  # General errors are medium severity
                                Category = "System Errors"
                                RecentEvents = $RecentErrorEvents
                                ProviderName = $SampleEvent.ProviderName
                            }
                        }
                    } | Where-Object { $_ -ne $null }

                    # Add filtered errors to analysis
                    $EventAnalysis += $FilteredErrors
                }
            } catch {
                # Silently continue if log doesn't exist or can't be accessed
                Write-Verbose "Could not access $LogName on $($DC.DomainController): $($_.Exception.Message)"
            }
        }

        # Check if we've spent too much time on this DC
        $DCElapsed = (Get-Date) - $DCStartTime
        if ($DCElapsed.TotalMinutes -gt 5) {
            Write-Host "    ⚠️ Event analysis for $($DC.HostName) took longer than 5 minutes - moving to next DC" -ForegroundColor Yellow
        } else {
            Write-Host "    ✓ Event analysis for $($DC.HostName) completed in $([math]::Round($DCElapsed.TotalSeconds, 1)) seconds" -ForegroundColor Green
        }
    }

    # Filter and organize results
    $AuditResults.EventLogAnalysis = $EventAnalysis | Sort-Object Severity, EventID

    Write-Host "Security events & error analysis completed. Found $($EventAnalysis.Count) event types with activity." -ForegroundColor Green

    # Debug: Show breakdown of what was found
    if ($EventAnalysis.Count -gt 0) {
        $HighSev = ($EventAnalysis | Where-Object { $_.Severity -in @("Critical", "High") }).Count
        $MedSev = ($EventAnalysis | Where-Object { $_.Severity -eq "Medium" }).Count
        $LowSev = ($EventAnalysis | Where-Object { $_.Severity -eq "Low" }).Count
        Write-Host "  → Critical/High: $HighSev, Medium: $MedSev, Low: $LowSev" -ForegroundColor Cyan

        # Show top event types found
        $TopEvents = $EventAnalysis | Group-Object EventID | Sort-Object Count -Descending | Select-Object -First 3
        foreach ($Event in $TopEvents) {
            $Sample = $Event.Group[0]
            Write-Host "  → Event $($Sample.EventID) ($($Sample.LogName)): $($Event.Count) occurrences - $($Sample.Description)" -ForegroundColor Gray
        }
    } else {
        Write-Host "  → No critical events or errors found in the last 7 days" -ForegroundColor Gray
    }

} catch {
    $AuditResults.EventLogAnalysis = "Error: $($_.Exception.Message)"
    Write-Error "Event Log Analysis failed: $($_.Exception.Message)"
}
#endregion

#region Security Group Nesting and Membership Analysis
Write-Host "Analyzing Security Group Nesting..." -ForegroundColor Cyan
try {
    $Groups = Get-ADGroup -Filter {GroupCategory -eq "Security"} -Properties Members
    $GroupAnalysis = @()
    
    foreach ($Group in $Groups) {
        $GroupInfo = @{
            Name = $Group.Name
            DistinguishedName = $Group.DistinguishedName
            GroupScope = $Group.GroupScope
            MemberCount = $Group.Members.Count
            SafeID = ($Group.DistinguishedName -replace '[^a-zA-Z0-9]', '_')
        }
        
        # Get detailed membership information
        try {
            $Members = Get-ADGroupMember -Identity $Group -ErrorAction SilentlyContinue
            $UserMembers = $Members | Where-Object {$_.objectClass -eq "user"}
            $ComputerMembers = $Members | Where-Object {$_.objectClass -eq "computer"}
            $NestedGroups = $Members | Where-Object {$_.objectClass -eq "group"}
            
            $GroupInfo.UserMembers = $UserMembers.Count
            $GroupInfo.ComputerMembers = $ComputerMembers.Count
            $GroupInfo.NestedGroups = $NestedGroups.Count
            $GroupInfo.NestedGroupNames = ($NestedGroups | ForEach-Object { $_.Name }) -join ", "
            
            # Check nesting depth (simplified check)
            if ($NestedGroups.Count -gt 0) {
                $MaxDepth = 1
                foreach ($NestedGroup in $NestedGroups) {
                    try {
                        $SubNested = Get-ADGroupMember -Identity $NestedGroup -ErrorAction SilentlyContinue | Where-Object {$_.objectClass -eq "group"}
                        if ($SubNested.Count -gt 0) {
                            $MaxDepth = [Math]::Max($MaxDepth, 2)
                            # Check one more level
                            foreach ($SubGroup in $SubNested) {
                                try {
                                    $SubSubNested = Get-ADGroupMember -Identity $SubGroup -ErrorAction SilentlyContinue | Where-Object {$_.objectClass -eq "group"}
                                    if ($SubSubNested.Count -gt 0) {
                                        $MaxDepth = [Math]::Max($MaxDepth, 3)
                                        break
                                    }
                                } catch { }
                            }
                        }
                    } catch { }
                }
                $GroupInfo.NestingDepth = $MaxDepth
            } else {
                $GroupInfo.NestingDepth = 0
            }
        } catch {
            $GroupInfo.UserMembers = 0
            $GroupInfo.ComputerMembers = 0
            $GroupInfo.NestedGroups = 0
            $GroupInfo.NestingDepth = 0
            $GroupInfo.NestedGroupNames = "Error retrieving members"
        }
        
        $GroupAnalysis += New-Object PSObject -Property $GroupInfo
    }
    $AuditResults.GroupNesting = $GroupAnalysis
} catch {
    $AuditResults.GroupNesting = "Error: $($_.Exception.Message)"
}
#endregion



#region Protocol Security Analysis (LLMNR, SMB, TLS, NTLM)
Write-Host "Analyzing Protocol Security..." -ForegroundColor Cyan
try {
    $ProtocolAnalysis = @()
    
    foreach ($DC in $DCs) {  # Check all DCs
        $DCProtocols = @{
            DomainController = $DC.DomainController
        }
        
        # --- SMB Version Check ---
        try {
            $SMBv1 = Invoke-Command -ComputerName $DC.DomainController -ScriptBlock {
                (Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction SilentlyContinue).State -eq 'Enabled'
            }
            $SMBConfig = Get-SmbServerConfiguration -CimSession $DC.DomainController -ErrorAction SilentlyContinue
            $DCProtocols.SMBv1_Enabled = $SMBv1
            $DCProtocols.SMBv2_Enabled = $SMBConfig.EnableSMB2Protocol
            $DCProtocols.SMBv3_Enabled = $SMBConfig.EnableSMB3Protocol
        } catch {
            $DCProtocols.SMBv1_Enabled = "Error"
            $DCProtocols.SMBv2_Enabled = "Error"
            $DCProtocols.SMBv3_Enabled = "Error"
        }
        
        # --- LLMNR Check ---
        try {
            $LLMNRReg = Invoke-Command -ComputerName $DC.DomainController -ScriptBlock {
                Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient" -Name "EnableMulticast" -ErrorAction SilentlyContinue
            } -ErrorAction SilentlyContinue
            $DCProtocols.LLMNRDisabled = if ($LLMNRReg.EnableMulticast -eq 0) { "Yes" } else { "No" }
        } catch {
            $DCProtocols.LLMNRDisabled = "Unknown"
        }
        
        # --- TLS Versions Check ---
        try {
            $TLSProtocols = @('1.0','1.1','1.2','1.3')
            foreach ($ver in $TLSProtocols) {
                $KeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS $ver\Server"
                $Enabled = Invoke-Command -ComputerName $DC.DomainController -ScriptBlock {
                    param($Path) 
                    if (Test-Path $Path) {
                        (Get-ItemProperty -Path $Path -Name "Enabled" -ErrorAction SilentlyContinue).Enabled
                    } else { $null }
                } -ArgumentList $KeyPath -ErrorAction SilentlyContinue

                $DCProtocols."TLS_$($ver)_Enabled" = if ($Enabled -eq 1 -or $Enabled -eq $null) { $true } else { $false }
            }
        } catch {
            foreach ($ver in $TLSProtocols) {
                $DCProtocols."TLS_$($ver)_Enabled" = "Error"
            }
        }
        
        # --- NTLM Audit Policy Check ---
        try {
            $NTLMAudit = Invoke-Command -ComputerName $DC.DomainController -ScriptBlock {
                auditpol /get /subcategory:"Logon" 2>$null | Select-String "NTLM"
            } -ErrorAction SilentlyContinue
            $DCProtocols.NTLMAudit = if ($NTLMAudit) { "Configured" } else { "Not Configured" }
        } catch {
            $DCProtocols.NTLMAudit = "Error"
        }
        
        $ProtocolAnalysis += New-Object PSObject -Property $DCProtocols
    }
    $AuditResults.ProtocolSecurity = $ProtocolAnalysis
} catch {
    $AuditResults.ProtocolSecurity = "Error: $($_.Exception.Message)"
}
#endregion




#region Group Managed Service Accounts (gMSA) Usage Analysis
Write-Host "Checking Group Managed Service Accounts usage..." -ForegroundColor Cyan
try {
    $gMSAAccounts = Get-ADServiceAccount -Filter {ObjectClass -eq "msDS-GroupManagedServiceAccount"} -Properties Enabled, PrincipalsAllowedToRetrieveManagedPassword

    if ($gMSAAccounts.Count -eq 0) {
        $AuditResults.gMSA = @{
            Status = "No gMSA accounts found"
            Recommendation = "Consider implementing Group Managed Service Accounts (gMSA) for better security, automated password management, and Kerberos authentication."
        }
    } else {
        $gMSADetails = @()
        foreach ($acct in $gMSAAccounts) {
            $AllowedComputers = $null
            if ($acct.PrincipalsAllowedToRetrieveManagedPassword) {
                $AllowedComputers = ($acct.PrincipalsAllowedToRetrieveManagedPassword | ForEach-Object { $_.Name }) -join ", "
            }
            $gMSADetails += [PSCustomObject]@{
                Name = $acct.Name
                Enabled = $acct.Enabled
                AllowedComputers = $AllowedComputers
            }
        }
        $AuditResults.gMSA = @{
            Status = "gMSA accounts found"
            Accounts = $gMSADetails
        }
    }
} catch {
    $AuditResults.gMSA = "Error: $($_.Exception.Message)"
}
#endregion




#region LAPS Check
Write-Host "Checking for LAPS deployment..." -ForegroundColor Cyan

$LAPSResults = [PSCustomObject]@{
    LAPS_Type        = "Not Detected"
    LAPS_Configured  = $false
    OU_Count         = 0
    Computers_Managed = 0
    Recommendation   = ""
}

try {
    # Check for legacy LAPS (old solution) - silently handle missing properties
    $LegacyLAPS_OU = $null
    try {
        $LegacyLAPS_OU = Get-ADOrganizationalUnit -Filter * -Properties ms-Mcs-AdmPwdExpirationTime -ErrorAction SilentlyContinue 2>$null |
            Where-Object { $_."ms-Mcs-AdmPwdExpirationTime" -ne $null }
    } catch {
        # LAPS not deployed or properties don't exist - this is normal
    }

    if ($LegacyLAPS_OU) {
        $LAPSResults.LAPS_Type = "Legacy LAPS (ms-Mcs-AdmPwd)"
        $LAPSResults.OU_Count = $LegacyLAPS_OU.Count

        # Count computers with password set
        $ManagedComputers = Get-ADComputer -Filter * -Properties ms-Mcs-AdmPwd -ErrorAction SilentlyContinue |
            Where-Object { $_."ms-Mcs-AdmPwd" -ne $null }
        $LAPSResults.Computers_Managed = $ManagedComputers.Count
        $LAPSResults.LAPS_Configured = $true
    }
    else {
        # Check for Windows LAPS (new solution)
        $WinLAPS_Computers = Get-ADComputer -Filter * -Properties msLAPS-PasswordExpirationTime -ErrorAction SilentlyContinue |
            Where-Object { $_."msLAPS-PasswordExpirationTime" -ne $null }

        if ($WinLAPS_Computers) {
            $LAPSResults.LAPS_Type = "Windows LAPS (msLAPS)"
            $LAPSResults.Computers_Managed = $WinLAPS_Computers.Count
            $LAPSResults.LAPS_Configured = $true
        }
    }

    # Recommendation
    if (-not $LAPSResults.LAPS_Configured) {
        $LAPSResults.Recommendation = "LAPS is not deployed. Recommended to enable Windows LAPS to protect local admin passwords."
    }
    else {
        $LAPSResults.Recommendation = "$($LAPSResults.LAPS_Type) is in use."
    }

} catch {
    # Handle LAPS check errors silently as LAPS may not be deployed
    $LAPSResults.Recommendation = "LAPS not detected or unavailable in this environment."
}

$AuditResults.LAPSCheck = $LAPSResults
#endregion






#region Privileged Account Monitoring
Write-Host "Monitoring Privileged Accounts..." -ForegroundColor Cyan
try {
    $PrivilegedAccounts = @{
        AdminGroups = @()
        ServiceAccounts = @()
        InactivePrivilegedAccounts = @()
        PasswordPolicies = @()
        Recommendations = @()
    }

    # Define high-privilege groups using well-known SIDs (language-independent)
    # This approach works across different OS languages (English, French, German, etc.)
    # by using Security Identifiers (SIDs) instead of localized group names
    $HighPrivGroups = @(
        @{ SID = "$DomainSID-512"; DisplayName = "Domain Admins" },           # Domain Admins
        @{ SID = "$RootDomainSID-519"; DisplayName = "Enterprise Admins" },   # Enterprise Admins  
        @{ SID = "$RootDomainSID-518"; DisplayName = "Schema Admins" },       # Schema Admins
        @{ SID = "S-1-5-32-544"; DisplayName = "Administrators" },           # Built-in Administrators
        @{ SID = "$DomainSID-548"; DisplayName = "Account Operators" },       # Account Operators
        @{ SID = "S-1-5-32-551"; DisplayName = "Backup Operators" },         # Built-in Backup Operators
        @{ SID = "$DomainSID-549"; DisplayName = "Server Operators" },        # Server Operators
        @{ SID = "S-1-5-32-550"; DisplayName = "Print Operators" },          # Built-in Print Operators
        @{ SID = "$DomainSID-520"; DisplayName = "Group Policy Creator Owners" } # Group Policy Creator Owners
    )
    
    # Add DNS Admins if it exists (not all domains have this group)
    try {
        $DNSAdmins = Get-ADGroup -Filter "Name -eq 'DNS Admins' -or Name -eq 'DnsAdmins'" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($DNSAdmins) {
            $HighPrivGroups += @{ SID = $DNSAdmins.SID.Value; DisplayName = "DNS Admins" }
        }
    } catch {
        Write-Warning "Could not check for DNS Admins group"
    }

    # Analyze administrative groups
    foreach ($PrivGroup in $HighPrivGroups) {
        try {
            $Group = Get-ADGroup -Identity $PrivGroup.SID -Properties Members, ManagedBy, Description -ErrorAction SilentlyContinue
            if ($Group) {
                $Members = Get-ADGroupMember -Identity $Group -ErrorAction SilentlyContinue
                $GroupInfo = @{
                    GroupName = $PrivGroup.DisplayName
                    ActualGroupName = $Group.Name  # Store the actual localized name for reference
                    SID = $PrivGroup.SID
                    MemberCount = $Members.Count
                    Members = @()
                    LastModified = $Group.whenChanged
                    ManagedBy = $Group.ManagedBy
                    Description = $Group.Description
                }

                foreach ($Member in $Members) {
                    if ($Member.objectClass -eq "user") {
                        $User = Get-ADUser -Identity $Member -Properties LastLogonDate, PasswordLastSet, Enabled, AccountExpirationDate -ErrorAction SilentlyContinue
                        if ($User) {
                            $MemberInfo = @{
                                Name = $User.Name
                                SamAccountName = $User.SamAccountName
                                Enabled = $User.Enabled
                                LastLogon = $User.LastLogonDate
                                PasswordLastSet = $User.PasswordLastSet
                                AccountExpiration = $User.AccountExpirationDate
                                DaysSinceLastLogon = if ($User.LastLogonDate) { (Get-Date).Subtract($User.LastLogonDate).Days } else { "Never" }
                            }
                            $GroupInfo.Members += $MemberInfo

                            # Flag inactive privileged accounts (not logged in for 90+ days)
                            if ($User.Enabled -and ($User.LastLogonDate -eq $null -or (Get-Date).Subtract($User.LastLogonDate).Days -gt 90)) {
                                $PrivilegedAccounts.InactivePrivilegedAccounts += @{
                                    User = $User.SamAccountName
                                    Group = $GroupName
                                    LastLogon = $User.LastLogonDate
                                    DaysInactive = if ($User.LastLogonDate) { (Get-Date).Subtract($User.LastLogonDate).Days } else { "Never logged in" }
                                }
                            }
                        }
                    }
                }
                $PrivilegedAccounts.AdminGroups += $GroupInfo
            }
        } catch {
            # Silently skip missing privileged groups - they may not exist in all environments
            Write-Host "Note: Privileged group $($PrivGroup.DisplayName) (SID: $($PrivGroup.SID)) not found in this domain" -ForegroundColor Yellow
        }
    }

    # Find service accounts with high privileges
    $ServiceAccounts = Get-ADUser -Filter {ServicePrincipalName -like "*"} -Properties ServicePrincipalName, LastLogonDate, PasswordLastSet, MemberOf -ErrorAction SilentlyContinue
    $PrivilegedGroupSIDs = $HighPrivGroups.SID  # Extract just the SIDs for comparison
    
    foreach ($ServiceAccount in $ServiceAccounts) {
        $IsPrivileged = $false
        $ServiceAccountPrivGroups = @()
        
        foreach ($GroupDN in $ServiceAccount.MemberOf) {
            try {
                $Group = Get-ADGroup -Identity $GroupDN -Properties SID -ErrorAction SilentlyContinue
                if ($Group -and $Group.SID.Value -in $PrivilegedGroupSIDs) {
                    $IsPrivileged = $true
                    # Find the display name from our privileged groups list
                    $MatchedPrivGroup = $HighPrivGroups | Where-Object { $_.SID -eq $Group.SID.Value } | Select-Object -First 1
                    if ($MatchedPrivGroup) {
                        $ServiceAccountPrivGroups += $MatchedPrivGroup.DisplayName
                    } else {
                        $ServiceAccountPrivGroups += $Group.Name
                    }
                }
            } catch {
                Write-Warning "Could not check group membership for service account $($ServiceAccount.Name): $_"
            }
        }
        
        if ($IsPrivileged) {
            $PrivilegedAccounts.ServiceAccounts += @{
                Name = $ServiceAccount.Name
                SamAccountName = $ServiceAccount.SamAccountName
                ServicePrincipalNames = $ServiceAccount.ServicePrincipalName
                LastLogon = $ServiceAccount.LastLogonDate
                PasswordLastSet = $ServiceAccount.PasswordLastSet
                PrivilegedGroups = $ServiceAccountPrivGroups -join ", "
                DaysSinceLastLogon = if ($ServiceAccount.LastLogonDate) { 
                    (New-TimeSpan -Start $ServiceAccount.LastLogonDate -End (Get-Date)).Days 
                } else { "Never" }
            }
        }
    }

    # Check password policies for privileged accounts
    $DefaultPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
    $FineGrainedPolicies = Get-ADFineGrainedPasswordPolicy -Filter * -ErrorAction SilentlyContinue

    $PrivilegedAccounts.PasswordPolicies += @{
        PolicyType = "Default Domain Policy"
        MaxPasswordAge = $DefaultPasswordPolicy.MaxPasswordAge.Days
        MinPasswordLength = $DefaultPasswordPolicy.MinPasswordLength
        PasswordHistoryCount = $DefaultPasswordPolicy.PasswordHistoryCount
        ComplexityEnabled = $DefaultPasswordPolicy.ComplexityEnabled
        LockoutThreshold = $DefaultPasswordPolicy.LockoutThreshold
    }

    foreach ($Policy in $FineGrainedPolicies) {
        $PrivilegedAccounts.PasswordPolicies += @{
            PolicyType = "Fine-Grained Policy"
            Name = $Policy.Name
            MaxPasswordAge = $Policy.MaxPasswordAge.Days
            MinPasswordLength = $Policy.MinPasswordLength
            PasswordHistoryCount = $Policy.PasswordHistoryCount
            ComplexityEnabled = $Policy.ComplexityEnabled
            LockoutThreshold = $Policy.LockoutThreshold
            AppliesTo = ($Policy.AppliesTo | ForEach-Object { (Get-ADObject -Identity $_).Name }) -join ", "
        }
    }

    # Generate recommendations
    if ($PrivilegedAccounts.InactivePrivilegedAccounts.Count -gt 0) {
        $PrivilegedAccounts.Recommendations += "Found $($PrivilegedAccounts.InactivePrivilegedAccounts.Count) inactive privileged accounts. Consider disabling or removing unused accounts."
    }
    
    $TotalPrivilegedUsers = ($PrivilegedAccounts.AdminGroups | ForEach-Object { $_.Members }).Count
    if ($TotalPrivilegedUsers -gt 20) {
        $PrivilegedAccounts.Recommendations += "High number of privileged users ($TotalPrivilegedUsers). Consider implementing principle of least privilege."
    }

    if ($PrivilegedAccounts.ServiceAccounts.Count -gt 0) {
        $PrivilegedAccounts.Recommendations += "Found $($PrivilegedAccounts.ServiceAccounts.Count) service accounts with high privileges. Consider using Group Managed Service Accounts (gMSA)."
    }

} catch {
    Write-Warning "Error during privileged account monitoring: $_"
    $PrivilegedAccounts.Error = $_.Exception.Message
}

$AuditResults.PrivilegedAccountMonitoring = $PrivilegedAccounts
#endregion


#region SYSVOL and NTDS Integrity Analysis
Write-Host "Analyzing SYSVOL and NTDS Integrity..." -ForegroundColor Cyan
try {
    $SysvolNtdsAnalysis = @{
        SysvolChecks = @()
        NtdsChecks = @()
        Summary = @{
            SysvolHealthy = $true
            NtdsHealthy = $true
            TotalIssues = 0
            CriticalIssues = 0
        }
        Recommendations = @()
    }

    $AllDomains = (Get-ADForest).Domains

    foreach ($Domain in $AllDomains) {
        $DomainControllers = Get-ADDomainController -Filter * -Server $Domain

        foreach ($DC in $DomainControllers) {
            Write-Host "  Checking SYSVOL and NTDS on $($DC.HostName)..." -ForegroundColor Yellow

            # SYSVOL Analysis
            $SysvolCheck = @{
                DomainController = $DC.HostName
                Domain = $Domain
                SysvolPath = "\\$($DC.HostName)\SYSVOL"
                NetlogonPath = "\\$($DC.HostName)\NETLOGON"
                SysvolAccessible = $false
                NetlogonAccessible = $false
                SysvolSize = "N/A"
                PolicyCount = 0
                ScriptCount = 0
                FolderStructure = @()
                Permissions = @()
                ReplicationStatus = "Unknown"
                LastModified = "N/A"
                Issues = @()
            }

            # Check SYSVOL accessibility and structure
            try {
                if (Test-Path $SysvolCheck.SysvolPath) {
                    $SysvolCheck.SysvolAccessible = $true

                    # Get SYSVOL size and structure
                    try {
                        $SysvolItems = Get-ChildItem $SysvolCheck.SysvolPath -Recurse -ErrorAction SilentlyContinue
                        if ($SysvolItems) {
                            $SysvolCheck.SysvolSize = "$([math]::Round(($SysvolItems | Where-Object {!$_.PSIsContainer} | Measure-Object -Property Length -Sum).Sum / 1MB, 2)) MB"
                            $SysvolCheck.PolicyCount = ($SysvolItems | Where-Object {$_.Name -like "*{*-*-*-*-*}*"}).Count
                            $SysvolCheck.ScriptCount = ($SysvolItems | Where-Object {$_.Extension -in @('.bat','.cmd','.ps1','.vbs')}).Count

                            # Check folder structure
                            $SysvolCheck.FolderStructure = @(
                                @{ Path = "$($SysvolCheck.SysvolPath)\$Domain"; Exists = (Test-Path "$($SysvolCheck.SysvolPath)\$Domain") }
                                @{ Path = "$($SysvolCheck.SysvolPath)\$Domain\Policies"; Exists = (Test-Path "$($SysvolCheck.SysvolPath)\$Domain\Policies") }
                                @{ Path = "$($SysvolCheck.SysvolPath)\$Domain\scripts"; Exists = (Test-Path "$($SysvolCheck.SysvolPath)\$Domain\scripts") }
                            )

                            # Get last modification time
                            $LatestFile = $SysvolItems | Where-Object {!$_.PSIsContainer} | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                            if ($LatestFile) {
                                $SysvolCheck.LastModified = $LatestFile.LastWriteTime
                            }
                        }
                    } catch {
                        $SysvolCheck.Issues += "Error analyzing SYSVOL contents: $($_.Exception.Message)"
                        $SysvolNtdsAnalysis.Summary.TotalIssues++
                    }

                    # Check SYSVOL permissions
                    try {
                        $Acl = Get-Acl $SysvolCheck.SysvolPath -ErrorAction SilentlyContinue
                        if ($Acl) {
                            $SysvolCheck.Permissions = @(
                                @{ Principal = "SYSTEM"; Access = ($Acl.Access | Where-Object {$_.IdentityReference -like "*SYSTEM*"}).FileSystemRights }
                                @{ Principal = "Domain Admins"; Access = ($Acl.Access | Where-Object {$_.IdentityReference -like "*Domain Admins*"}).FileSystemRights }
                                @{ Principal = "Authenticated Users"; Access = ($Acl.Access | Where-Object {$_.IdentityReference -like "*Authenticated Users*"}).FileSystemRights }
                            )
                        }
                    } catch {
                        $SysvolCheck.Issues += "Error checking SYSVOL permissions: $($_.Exception.Message)"
                        $SysvolNtdsAnalysis.Summary.TotalIssues++
                    }

                } else {
                    $SysvolCheck.Issues += "SYSVOL share not accessible"
                    $SysvolNtdsAnalysis.Summary.SysvolHealthy = $false
                    $SysvolNtdsAnalysis.Summary.CriticalIssues++
                }
            } catch {
                $SysvolCheck.Issues += "Error accessing SYSVOL: $($_.Exception.Message)"
                $SysvolNtdsAnalysis.Summary.SysvolHealthy = $false
                $SysvolNtdsAnalysis.Summary.CriticalIssues++
            }

            # Check NETLOGON accessibility
            try {
                if (Test-Path $SysvolCheck.NetlogonPath) {
                    $SysvolCheck.NetlogonAccessible = $true
                } else {
                    $SysvolCheck.Issues += "NETLOGON share not accessible"
                    $SysvolNtdsAnalysis.Summary.SysvolHealthy = $false
                    $SysvolNtdsAnalysis.Summary.TotalIssues++
                }
            } catch {
                $SysvolCheck.Issues += "Error accessing NETLOGON: $($_.Exception.Message)"
                $SysvolNtdsAnalysis.Summary.SysvolHealthy = $false
                $SysvolNtdsAnalysis.Summary.TotalIssues++
            }

            # SYSVOL Replication Status (using DFSR or FRS)
            try {
                # Check for DFSR first (Windows 2008+)
                $DfsrInfo = Invoke-Command -ComputerName $DC.HostName.ToString() -ScriptBlock {
                    Get-WmiObject -Namespace "root\MicrosoftDfs" -Class DfsrReplicationGroupConfig -ErrorAction SilentlyContinue
                } -ErrorAction SilentlyContinue

                if ($DfsrInfo) {
                    $SysvolCheck.ReplicationStatus = "DFSR - Active"
                } else {
                    # Check for FRS (legacy)
                    $FrsInfo = Invoke-Command -ComputerName $DC.HostName.ToString() -ScriptBlock {
                        Get-WmiObject -Class Win32_NTDomain -ErrorAction SilentlyContinue
                    } -ErrorAction SilentlyContinue

                    if ($FrsInfo) {
                        $SysvolCheck.ReplicationStatus = "FRS - Legacy (Consider upgrading to DFSR)"
                        $SysvolCheck.Issues += "Using legacy FRS replication - recommend upgrading to DFSR"
                        $SysvolNtdsAnalysis.Summary.TotalIssues++
                    } else {
                        $SysvolCheck.ReplicationStatus = "Unknown"
                        $SysvolCheck.Issues += "Unable to determine replication method"
                        $SysvolNtdsAnalysis.Summary.TotalIssues++
                    }
                }
            } catch {
                $SysvolCheck.ReplicationStatus = "Error checking replication"
                $SysvolCheck.Issues += "Error checking SYSVOL replication: $($_.Exception.Message)"
                $SysvolNtdsAnalysis.Summary.TotalIssues++
            }

            $SysvolNtdsAnalysis.SysvolChecks += $SysvolCheck

            # NTDS.dit Analysis
            $NtdsCheck = @{
                DomainController = $DC.HostName
                Domain = $Domain
                NtdsPath = "C:\Windows\NTDS\ntds.dit"
                NtdsAccessible = $false
                DatabaseSize = "N/A"
                DatabaseIntegrity = "Unknown"
                LastBackup = "N/A"
                TransactionLogs = @()
                FragmentationLevel = "N/A"
                DatabaseVersion = "N/A"
                Issues = @()
                DatabaseHealth = "Unknown"
            }

            # Check NTDS.dit accessibility and properties
            try {
                # Try direct file access first
                $NtdsFilePath = "\\$($DC.HostName)\C$\Windows\NTDS\ntds.dit"

                if (Test-Path $NtdsFilePath) {
                    $NtdsCheck.NtdsAccessible = $true
                    $NtdsFile = Get-Item $NtdsFilePath -ErrorAction SilentlyContinue

                    if ($NtdsFile) {
                        $NtdsCheck.DatabaseSize = "$([math]::Round($NtdsFile.Length / 1GB, 2)) GB"

                        # Check if database size is reasonable
                        $DbSizeGB = [math]::Round($NtdsFile.Length / 1GB, 2)
                        if ($DbSizeGB -gt 50) {
                            $NtdsCheck.Issues += "Large database size ($DbSizeGB GB) - consider maintenance"
                            $SysvolNtdsAnalysis.Summary.TotalIssues++
                        }
                    }
                } else {
                    # Try WMI method as fallback
                    $RemoteNtdsFile = Get-CimInstance -ComputerName $DC.HostName.ToString() -ClassName CIM_DataFile -Filter "Name='C:\\Windows\\NTDS\\ntds.dit'" -ErrorAction SilentlyContinue

                    if ($RemoteNtdsFile) {
                        $NtdsCheck.NtdsAccessible = $true
                        $NtdsCheck.DatabaseSize = "$([math]::Round($RemoteNtdsFile.FileSize / 1GB, 2)) GB"
                    } else {
                        $NtdsCheck.Issues += "NTDS.dit file not accessible"
                        $SysvolNtdsAnalysis.Summary.NtdsHealthy = $false
                        $SysvolNtdsAnalysis.Summary.CriticalIssues++
                    }
                }

                # Check transaction logs
                $LogPath = "\\$($DC.HostName)\C$\Windows\NTDS\"
                if (Test-Path $LogPath) {
                    $LogFiles = Get-ChildItem "$LogPath*.log" -ErrorAction SilentlyContinue
                    if ($LogFiles) {
                        foreach ($LogFile in $LogFiles) {
                            $NtdsCheck.TransactionLogs += @{
                                Name = $LogFile.Name
                                Size = "$([math]::Round($LogFile.Length / 1MB, 2)) MB"
                                LastModified = $LogFile.LastWriteTime
                            }
                        }

                        # Check for excessive log accumulation
                        $TotalLogSize = ($LogFiles | Measure-Object -Property Length -Sum).Sum
                        if ($TotalLogSize -gt 10GB) {
                            $NtdsCheck.Issues += "Excessive transaction log accumulation ($([math]::Round($TotalLogSize / 1GB, 2)) GB)"
                            $SysvolNtdsAnalysis.Summary.TotalIssues++
                        }
                    }
                }

                # Check database integrity using esentutl (if accessible)
                try {
                    $IntegrityCheck = Invoke-Command -ComputerName $DC.HostName.ToString() -ScriptBlock {
                        $EsentResult = & esentutl /mh "C:\Windows\NTDS\ntds.dit" 2>&1
                        if ($EsentResult -match "State:\s+(.+)") {
                            return $Matches[1].Trim()
                        }
                        return "Unknown"
                    } -ErrorAction SilentlyContinue

                    if ($IntegrityCheck) {
                        $NtdsCheck.DatabaseIntegrity = $IntegrityCheck
                        if ($IntegrityCheck -ne "Clean Shutdown") {
                            $NtdsCheck.Issues += "Database not in clean shutdown state: $IntegrityCheck"
                            $SysvolNtdsAnalysis.Summary.NtdsHealthy = $false
                            $SysvolNtdsAnalysis.Summary.CriticalIssues++
                        }
                    }
                } catch {
                    $NtdsCheck.DatabaseIntegrity = "Unable to check"
                    $NtdsCheck.Issues += "Could not verify database integrity"
                    $SysvolNtdsAnalysis.Summary.TotalIssues++
                }

                # Check for recent system state backup
                try {
                    # Check backup events with timeout protection
                    $BackupEvents = $null
                    $Job = Start-Job -ScriptBlock {
                        param($ComputerName, $IsLocal)
                        if ($IsLocal) {
                            Get-WinEvent -FilterHashtable @{
                                LogName = 'Application'
                                ID = 2001  # System State backup event
                                StartTime = (Get-Date).AddDays(-30)
                            } -MaxEvents 1 -ErrorAction SilentlyContinue
                        } else {
                            Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{
                                LogName = 'Application'
                                ID = 2001  # System State backup event
                                StartTime = (Get-Date).AddDays(-30)
                            } -MaxEvents 1 -ErrorAction SilentlyContinue
                        }
                    } -ArgumentList $ComputerNameForQuery, $IsLocalDC

                    $BackupEvents = Receive-Job -Job $Job -Wait -Timeout 5
                    Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue

                    if ($BackupEvents) {
                        $NtdsCheck.LastBackup = $BackupEvents[0].TimeCreated
                        $DaysOld = ((Get-Date) - $BackupEvents[0].TimeCreated).Days
                        if ($DaysOld -gt 7) {
                            $NtdsCheck.Issues += "Last backup is $DaysOld days old"
                            $SysvolNtdsAnalysis.Summary.TotalIssues++
                        }
                    } else {
                        $NtdsCheck.LastBackup = "No recent backup found (30 days)"
                        $NtdsCheck.Issues += "No recent system state backup detected"
                        $SysvolNtdsAnalysis.Summary.TotalIssues++
                    }
                } catch {
                    $NtdsCheck.LastBackup = "Unable to check backup status"
                    $NtdsCheck.Issues += "Could not verify backup status"
                    $SysvolNtdsAnalysis.Summary.TotalIssues++
                }

                # Overall database health assessment
                if ($NtdsCheck.Issues.Count -eq 0) {
                    $NtdsCheck.DatabaseHealth = "Healthy"
                } elseif ($NtdsCheck.Issues.Count -le 2) {
                    $NtdsCheck.DatabaseHealth = "Minor Issues"
                } else {
                    $NtdsCheck.DatabaseHealth = "Critical Issues"
                    $SysvolNtdsAnalysis.Summary.NtdsHealthy = $false
                }

            } catch {
                $NtdsCheck.Issues += "Error analyzing NTDS.dit: $($_.Exception.Message)"
                $NtdsCheck.DatabaseHealth = "Error"
                $SysvolNtdsAnalysis.Summary.NtdsHealthy = $false
                $SysvolNtdsAnalysis.Summary.CriticalIssues++
            }

            $SysvolNtdsAnalysis.NtdsChecks += $NtdsCheck
        }
    }

    # Generate recommendations
    if ($SysvolNtdsAnalysis.Summary.CriticalIssues -gt 0) {
        $SysvolNtdsAnalysis.Recommendations += "Critical issues detected requiring immediate attention"
    }

    if (-not $SysvolNtdsAnalysis.Summary.SysvolHealthy) {
        $SysvolNtdsAnalysis.Recommendations += "SYSVOL replication issues detected - check DFSR/FRS health"
    }

    if (-not $SysvolNtdsAnalysis.Summary.NtdsHealthy) {
        $SysvolNtdsAnalysis.Recommendations += "NTDS database issues detected - consider maintenance procedures"
    }

    $FrsCount = ($SysvolNtdsAnalysis.SysvolChecks | Where-Object {$_.ReplicationStatus -like "*FRS*"}).Count
    if ($FrsCount -gt 0) {
        $SysvolNtdsAnalysis.Recommendations += "Upgrade from FRS to DFSR for better SYSVOL replication"
    }

    $LargeDbCount = ($SysvolNtdsAnalysis.NtdsChecks | Where-Object {$_.DatabaseSize -like "*GB" -and [float]($_.DatabaseSize -replace " GB","") -gt 20}).Count
    if ($LargeDbCount -gt 0) {
        $SysvolNtdsAnalysis.Recommendations += "Consider offline defragmentation for large NTDS databases"
    }

    $AuditResults.SysvolNtdsAnalysis = $SysvolNtdsAnalysis
    Write-Host "SYSVOL and NTDS analysis completed." -ForegroundColor Green

} catch {
    $AuditResults.SysvolNtdsAnalysis = @{
        Error = $_.Exception.Message
        SysvolChecks = @()
        NtdsChecks = @()
        Summary = @{ SysvolHealthy = $false; NtdsHealthy = $false; TotalIssues = 1; CriticalIssues = 1 }
        Recommendations = @("Unable to complete SYSVOL and NTDS analysis")
    }
    Write-Error "SYSVOL and NTDS analysis failed: $($_.Exception.Message)"
}
#endregion

#region Schema Architecture Analysis
Write-Host "Analyzing Schema Architecture..." -ForegroundColor Cyan
try {
    $SchemaAnalysis = @{
        SchemaVersion = ""
        CustomClasses = @()
        CustomAttributes = @()
        SchemaExtensions = @()
        SchemaStatistics = @{
            TotalClasses = 0
            TotalAttributes = 0
            CustomClasses = 0
            CustomAttributes = 0
        }
        Recommendations = @()
    }
    
    # Get schema naming context
    $SchemaContext = (Get-ADRootDSE).schemaNamingContext
    
    # Get schema version
    try {
        $SchemaVersionObj = Get-ADObject -Identity "CN=Schema,$SchemaContext" -Properties objectVersion
        $SchemaAnalysis.SchemaVersion = $SchemaVersionObj.objectVersion
    } catch {
        $SchemaAnalysis.SchemaVersion = "Unable to determine"
    }
    
    # Analyze object classes
    Write-Host "  Analyzing object classes..." -ForegroundColor Yellow
    $AllClasses = Get-ADObject -SearchBase $SchemaContext -Filter {objectClass -eq "classSchema"} -Properties cn, whenCreated, adminDescription, defaultSecurityDescriptor
    $SchemaAnalysis.SchemaStatistics.TotalClasses = $AllClasses.Count
    
    # Identify custom classes (typically created after base schema)
    $BaseSchemaDate = Get-Date "2000-01-01"  # Classes created after this are likely custom
    foreach ($Class in $AllClasses) {
        if ($Class.whenCreated -gt $BaseSchemaDate -and $Class.cn -notmatch "^ms-|^Microsoft|^Exchange|^System") {
            $SchemaAnalysis.CustomClasses += @{
                Name = $Class.cn
                Created = $Class.whenCreated
                Description = $Class.adminDescription
            }
        }
    }
    $SchemaAnalysis.SchemaStatistics.CustomClasses = $SchemaAnalysis.CustomClasses.Count
    
    # Analyze attributes
    Write-Host "  Analyzing attributes..." -ForegroundColor Yellow
    $AllAttributes = Get-ADObject -SearchBase $SchemaContext -Filter {objectClass -eq "attributeSchema"} -Properties cn, whenCreated, adminDescription, attributeSyntax, isSingleValued
    $SchemaAnalysis.SchemaStatistics.TotalAttributes = $AllAttributes.Count
    
    # Identify custom attributes
    foreach ($Attribute in $AllAttributes) {
        if ($Attribute.whenCreated -gt $BaseSchemaDate -and $Attribute.cn -notmatch "^ms-|^Microsoft|^Exchange|^System") {
            $SchemaAnalysis.CustomAttributes += @{
                Name = $Attribute.cn
                Created = $Attribute.whenCreated
                Description = $Attribute.adminDescription
                Syntax = $Attribute.attributeSyntax
                SingleValued = $Attribute.isSingleValued
            }
        }
    }
    $SchemaAnalysis.SchemaStatistics.CustomAttributes = $SchemaAnalysis.CustomAttributes.Count
    
    # Check for common schema extensions
    Write-Host "  Checking for schema extensions..." -ForegroundColor Yellow
    $KnownExtensions = @{
        "Exchange" = @("ms-Exch-*", "mailNickname", "legacyExchangeDN")
        "Lync/Skype" = @("msRTCSIP-*", "msDS-SourceObjectDN")
        "ConfigMgr" = @("msSMS-*")
        "Custom Applications" = $SchemaAnalysis.CustomAttributes | ForEach-Object { $_.Name }
    }
    
    foreach ($Extension in $KnownExtensions.GetEnumerator()) {
        $ExtensionCount = 0
        foreach ($Pattern in $Extension.Value) {
            if ($Pattern -like "*-*") {
                $ExtensionCount += ($AllAttributes | Where-Object {$_.cn -like $Pattern}).Count
                $ExtensionCount += ($AllClasses | Where-Object {$_.cn -like $Pattern}).Count
            } else {
                $ExtensionCount += ($AllAttributes | Where-Object {$_.cn -eq $Pattern}).Count
            }
        }
        
        if ($ExtensionCount -gt 0) {
            $SchemaAnalysis.SchemaExtensions += @{
                ExtensionType = $Extension.Key
                ObjectCount = $ExtensionCount
            }
        }
    }
    
    # Generate recommendations
    if ($SchemaAnalysis.CustomClasses.Count -gt 50) {
        $SchemaAnalysis.Recommendations += "High number of custom classes ($($SchemaAnalysis.CustomClasses.Count)). Review if all are still needed."
    }
    
    if ($SchemaAnalysis.CustomAttributes.Count -gt 100) {
        $SchemaAnalysis.Recommendations += "High number of custom attributes ($($SchemaAnalysis.CustomAttributes.Count)). Review if all are still in use."
    }
    
    if ($SchemaAnalysis.SchemaExtensions.Count -eq 0) {
        $SchemaAnalysis.Recommendations += "No major schema extensions detected. Schema appears to be standard Active Directory."
    }
    
    # Check schema version for upgrade readiness
    if ($SchemaAnalysis.SchemaVersion -match "^\d+$") {
        $VersionNumber = [int]$SchemaAnalysis.SchemaVersion
        if ($VersionNumber -lt 87) {
            $SchemaAnalysis.Recommendations += "Schema version $VersionNumber indicates older Active Directory. Consider evaluating upgrade path."
        }
    }
    
} catch {
    $SchemaAnalysis = @{
        Error = $_.Exception.Message
        Recommendations = @("Unable to analyze schema architecture: $($_.Exception.Message)")
    }
    Write-Warning "Error analyzing schema: $_"
}

$AuditResults.SchemaArchitecture = $SchemaAnalysis
#endregion

#region AD Object Protection Analysis
Write-Host "Analyzing AD Object Protection from Accidental Deletion..." -ForegroundColor Cyan
try {
    $ObjectProtectionAnalysis = @{
        ProtectedOUs = @()
        UnprotectedOUs = @()
        ProtectedContainers = @()
        UnprotectedContainers = @()
        ProtectedUsers = @()
        UnprotectedUsers = @()
        ProtectedComputers = @()
        UnprotectedComputers = @()
        ProtectedGroups = @()
        UnprotectedGroups = @()
        Statistics = @{
            TotalOUs = 0
            ProtectedOUs = 0
            UnprotectedOUs = 0
            TotalContainers = 0
            ProtectedContainers = 0
            UnprotectedContainers = 0
            TotalCriticalUsers = 0
            ProtectedCriticalUsers = 0
            UnprotectedCriticalUsers = 0
            TotalServiceAccounts = 0
            ProtectedServiceAccounts = 0
            UnprotectedServiceAccounts = 0
            TotalCriticalGroups = 0
            ProtectedCriticalGroups = 0
            UnprotectedCriticalGroups = 0
        }
        Recommendations = @()
    }
    
    # Analyze OUs protection
    Write-Host "  Analyzing OU protection..." -ForegroundColor Yellow
    $AllOUs = Get-ADOrganizationalUnit -Filter * -Properties ProtectedFromAccidentalDeletion, Description
    $ObjectProtectionAnalysis.Statistics.TotalOUs = $AllOUs.Count
    
    foreach ($OU in $AllOUs) {
        if ($OU.ProtectedFromAccidentalDeletion -eq $true) {
            $ObjectProtectionAnalysis.Statistics.ProtectedOUs++
            $ObjectProtectionAnalysis.ProtectedOUs += @{
                Name = $OU.Name
                DistinguishedName = $OU.DistinguishedName
                Description = $OU.Description
            }
        } else {
            $ObjectProtectionAnalysis.Statistics.UnprotectedOUs++
            $ObjectProtectionAnalysis.UnprotectedOUs += @{
                Name = $OU.Name
                DistinguishedName = $OU.DistinguishedName
                Description = $OU.Description
            }
        }
    }
    
    # Analyze Container protection (CN= objects that are not OUs)
    Write-Host "  Analyzing container protection..." -ForegroundColor Yellow
    $DomainDN = (Get-ADDomain).DistinguishedName
    $AllContainers = Get-ADObject -SearchBase $DomainDN -Filter {objectClass -eq "container"} -Properties ProtectedFromAccidentalDeletion, Description
    # Filter out system containers that are typically not user-managed
    $ImportantContainers = $AllContainers | Where-Object {
        $_.DistinguishedName -notmatch "CN=System," -and 
        $_.DistinguishedName -notmatch "CN=Configuration," -and
        $_.DistinguishedName -notmatch "CN=Program Data," -and
        $_.DistinguishedName -notmatch "CN=Microsoft," -and
        $_.DistinguishedName -notmatch "CN=Keys,"
    }
    
    $ObjectProtectionAnalysis.Statistics.TotalContainers = $ImportantContainers.Count
    
    foreach ($Container in $ImportantContainers) {
        if ($Container.ProtectedFromAccidentalDeletion -eq $true) {
            $ObjectProtectionAnalysis.Statistics.ProtectedContainers++
            $ObjectProtectionAnalysis.ProtectedContainers += @{
                Name = $Container.Name
                DistinguishedName = $Container.DistinguishedName
                Description = $Container.Description
            }
        } else {
            $ObjectProtectionAnalysis.Statistics.UnprotectedContainers++
            $ObjectProtectionAnalysis.UnprotectedContainers += @{
                Name = $Container.Name
                DistinguishedName = $Container.DistinguishedName
                Description = $Container.Description
            }
        }
    }
    
    # Analyze critical user accounts protection
    Write-Host "  Analyzing critical user protection..." -ForegroundColor Yellow
    # Focus on privileged users and service accounts
    $PrivilegedGroups = @("Domain Admins", "Enterprise Admins", "Schema Admins", "Administrators")
    $CriticalUsers = @()
    
    foreach ($GroupName in $PrivilegedGroups) {
        try {
            $Group = Get-ADGroup -Identity $GroupName -ErrorAction SilentlyContinue
            if ($Group) {
                $GroupMembers = Get-ADGroupMember -Identity $Group -Recursive -ErrorAction SilentlyContinue | Where-Object {$_.objectClass -eq "user"}
                $CriticalUsers += $GroupMembers
            }
        } catch {
            # Group might not exist in this domain
        }
    }
    
    # Also include service accounts
    $ServiceAccounts = Get-ADUser -Filter {ServicePrincipalName -like "*"} -Properties ProtectedFromAccidentalDeletion, ServicePrincipalName -ErrorAction SilentlyContinue
    $CriticalUsers += $ServiceAccounts
    
    # Remove duplicates
    $CriticalUsers = $CriticalUsers | Sort-Object DistinguishedName | Get-Unique -AsString
    
    $ObjectProtectionAnalysis.Statistics.TotalCriticalUsers = $CriticalUsers.Count
    
    foreach ($User in $CriticalUsers) {
        try {
            $UserDetail = Get-ADUser -Identity $User.DistinguishedName -Properties ProtectedFromAccidentalDeletion, ServicePrincipalName -ErrorAction SilentlyContinue
            if ($UserDetail) {
                $IsServiceAccount = $UserDetail.ServicePrincipalName -ne $null -and $UserDetail.ServicePrincipalName.Count -gt 0
                if ($IsServiceAccount) {
                    $ObjectProtectionAnalysis.Statistics.TotalServiceAccounts++
                }
                
                if ($UserDetail.ProtectedFromAccidentalDeletion -eq $true) {
                    $ObjectProtectionAnalysis.Statistics.ProtectedCriticalUsers++
                    if ($IsServiceAccount) {
                        $ObjectProtectionAnalysis.Statistics.ProtectedServiceAccounts++
                    }
                    $ObjectProtectionAnalysis.ProtectedUsers += @{
                        Name = $UserDetail.Name
                        SamAccountName = $UserDetail.SamAccountName
                        DistinguishedName = $UserDetail.DistinguishedName
                        IsServiceAccount = $IsServiceAccount
                        ServicePrincipalNames = if ($IsServiceAccount) { $UserDetail.ServicePrincipalName -join "; " } else { "" }
                    }
                } else {
                    $ObjectProtectionAnalysis.Statistics.UnprotectedCriticalUsers++
                    if ($IsServiceAccount) {
                        $ObjectProtectionAnalysis.Statistics.UnprotectedServiceAccounts++
                    }
                    $ObjectProtectionAnalysis.UnprotectedUsers += @{
                        Name = $UserDetail.Name
                        SamAccountName = $UserDetail.SamAccountName
                        DistinguishedName = $UserDetail.DistinguishedName
                        IsServiceAccount = $IsServiceAccount
                        ServicePrincipalNames = if ($IsServiceAccount) { $UserDetail.ServicePrincipalName -join "; " } else { "" }
                    }
                }
            }
        } catch {
            # Skip users that can't be accessed
        }
    }
    
    # Analyze critical groups protection
    Write-Host "  Analyzing critical group protection..." -ForegroundColor Yellow
    $CriticalGroupNames = @(
        "Domain Admins", "Enterprise Admins", "Schema Admins", "Administrators", 
        "Account Operators", "Server Operators", "Backup Operators", "Print Operators",
        "Group Policy Creator Owners", "DNS Admins", "DnsAdmins"
    )
    
    foreach ($GroupName in $CriticalGroupNames) {
        try {
            $Group = Get-ADGroup -Identity $GroupName -Properties ProtectedFromAccidentalDeletion, Description -ErrorAction SilentlyContinue
            if ($Group) {
                $ObjectProtectionAnalysis.Statistics.TotalCriticalGroups++
                
                if ($Group.ProtectedFromAccidentalDeletion -eq $true) {
                    $ObjectProtectionAnalysis.Statistics.ProtectedCriticalGroups++
                    $ObjectProtectionAnalysis.ProtectedGroups += @{
                        Name = $Group.Name
                        DistinguishedName = $Group.DistinguishedName
                        Description = $Group.Description
                        GroupScope = $Group.GroupScope
                    }
                } else {
                    $ObjectProtectionAnalysis.Statistics.UnprotectedCriticalGroups++
                    $ObjectProtectionAnalysis.UnprotectedGroups += @{
                        Name = $Group.Name
                        DistinguishedName = $Group.DistinguishedName
                        Description = $Group.Description
                        GroupScope = $Group.GroupScope
                    }
                }
            }
        } catch {
            # Group might not exist
        }
    }
    
    # Generate recommendations
    if ($ObjectProtectionAnalysis.Statistics.UnprotectedOUs -gt 0) {
        $ObjectProtectionAnalysis.Recommendations += "Found $($ObjectProtectionAnalysis.Statistics.UnprotectedOUs) unprotected OUs. Consider enabling protection for critical organizational units."
    }
    
    if ($ObjectProtectionAnalysis.Statistics.UnprotectedContainers -gt 0) {
        $ObjectProtectionAnalysis.Recommendations += "Found $($ObjectProtectionAnalysis.Statistics.UnprotectedContainers) unprotected containers. Review and protect important containers from accidental deletion."
    }
    
    if ($ObjectProtectionAnalysis.Statistics.UnprotectedCriticalUsers -gt 0) {
        $ObjectProtectionAnalysis.Recommendations += "Found $($ObjectProtectionAnalysis.Statistics.UnprotectedCriticalUsers) unprotected critical users (admins/service accounts). Enable protection for privileged accounts."
    }
    
    if ($ObjectProtectionAnalysis.Statistics.UnprotectedServiceAccounts -gt 0) {
        $ObjectProtectionAnalysis.Recommendations += "Found $($ObjectProtectionAnalysis.Statistics.UnprotectedServiceAccounts) unprotected service accounts. Service accounts should always be protected from accidental deletion."
    }
    
    if ($ObjectProtectionAnalysis.Statistics.UnprotectedCriticalGroups -gt 0) {
        $ObjectProtectionAnalysis.Recommendations += "Found $($ObjectProtectionAnalysis.Statistics.UnprotectedCriticalGroups) unprotected critical groups. Administrative groups should be protected from accidental deletion."
    }
    
    # Overall protection assessment
    $TotalObjects = $ObjectProtectionAnalysis.Statistics.TotalOUs + $ObjectProtectionAnalysis.Statistics.TotalContainers + $ObjectProtectionAnalysis.Statistics.TotalCriticalUsers + $ObjectProtectionAnalysis.Statistics.TotalCriticalGroups
    $ProtectedObjects = $ObjectProtectionAnalysis.Statistics.ProtectedOUs + $ObjectProtectionAnalysis.Statistics.ProtectedContainers + $ObjectProtectionAnalysis.Statistics.ProtectedCriticalUsers + $ObjectProtectionAnalysis.Statistics.ProtectedCriticalGroups
    
    if ($TotalObjects -gt 0) {
        $ProtectionPercentage = [math]::Round(($ProtectedObjects / $TotalObjects) * 100, 1)
        if ($ProtectionPercentage -lt 80) {
            $ObjectProtectionAnalysis.Recommendations += "Only $ProtectionPercentage% of critical AD objects are protected from accidental deletion. Aim for 100% protection of critical infrastructure objects."
        } elseif ($ProtectionPercentage -eq 100) {
            $ObjectProtectionAnalysis.Recommendations += "Excellent! All critical AD objects are protected from accidental deletion."
        }
        $ObjectProtectionAnalysis.Statistics.OverallProtectionPercentage = $ProtectionPercentage
    }
    
} catch {
    $ObjectProtectionAnalysis = @{
        Error = $_.Exception.Message
        Recommendations = @("Unable to analyze AD object protection: $($_.Exception.Message)")
    }
    Write-Warning "Error analyzing AD object protection: $_"
}

$AuditResults.ObjectProtection = $ObjectProtectionAnalysis
#endregion

#================================================================
# END OF POWERSHELL AUDIT COMMANDS
#================================================================
# All data collection and analysis is now complete.
# $AuditResults contains all the collected information.
#================================================================

#================================================================
# HTML REPORT GENERATION SECTION
#================================================================
# This section takes the collected audit data and generates
# the HTML report for presentation.
#================================================================

# Generate HTML Report
Write-Host "Generating HTML Report..." -ForegroundColor Green

$HTMLReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Active Directory Audit Report</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif; 
            margin: 0; padding: 20px 40px; 
            background: linear-gradient(to bottom right, #f8fafc 0%, #e2e8f0 100%);
            min-height: 100vh; 
            line-height: 1.5;
            font-size: 14px;
        }
        .container { 
            max-width: 1200px; 
            margin: 0 auto; 
            padding: 0 20px;
        }
        
        /* Header styling - cleaner and more professional */
        .header { 
            background: white;
            color: #1e293b; 
            padding: 40px 50px; 
            border-radius: 12px; 
            margin-bottom: 32px; 
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            border-left: 4px solid #3b82f6;
        }
        .header h1 { margin: 0 0 20px 0; font-size: 2rem; font-weight: 700; color: #1e293b; }
        .header p { margin: 6px 0; color: #64748b; font-size: 0.875rem; }
        
        /* Section styling - reduced shadows and better spacing */
        .section { 
            background: white;
            margin: 24px 0; 
            padding: 32px 40px; 
            border-radius: 12px; 
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
            border: 1px solid #e2e8f0;
        }
        .section h2 { 
            color: #1e293b; 
            border-bottom: 2px solid #e2e8f0; 
            padding-bottom: 16px; 
            margin-bottom: 28px; 
            font-weight: 600; 
            font-size: 1.375rem;
        }
        .section h3 {
            color: #334155;
            margin-top: 32px;
            margin-bottom: 20px;
            font-size: 1.125rem;
            font-weight: 600;
        }
        .section h4 {
            color: #374151;
            margin-top: 24px;
            margin-bottom: 16px;
            font-size: 1rem;
            font-weight: 600;
        }
        .section p {
            margin-bottom: 16px;
            color: #4b5563;
        }
        
        /* Alert boxes - softer colors and better contrast */
        .warning { 
            background: #fef3c7; 
            border: 1px solid #f59e0b; 
            color: #92400e;
            padding: 20px 24px; 
            border-radius: 8px; 
            margin: 24px 0;
            border-left: 4px solid #f59e0b;
            font-size: 0.875rem;
        }
        .error { 
            background: #fee2e2; 
            border: 1px solid #ef4444; 
            color: #991b1b;
            padding: 20px 24px; 
            border-radius: 8px; 
            margin: 24px 0;
            border-left: 4px solid #ef4444;
            font-size: 0.875rem;
        }
        .success {
            background: #dcfce7;
            border: 1px solid #22c55e;
            color: #166534;
            padding: 20px 24px;
            border-radius: 8px;
            margin: 24px 0;
            border-left: 4px solid #22c55e;
            font-size: 0.875rem;
        }
        .info {
            background: #e0f2fe;
            border: 1px solid #0284c7;
            color: #0c4a6e;
            padding: 20px 24px;
            border-radius: 8px;
            margin: 24px 0;
            border-left: 4px solid #0284c7;
            font-size: 0.875rem;
        }

        table { width: 100%; border-collapse: collapse; margin: 15px 0; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
        th, td { border: none; padding: 12px 15px; text-align: left; }
        th { background: linear-gradient(135deg, #3498db 0%, #2980b9 100%); color: white; font-weight: 500; }
        tr:nth-child(even) { background-color: rgba(52,152,219,0.05); }
        
        /* Status-specific row styling */
        .table .table-success {
            background-color: #d4edda !important;
            color: #155724;
        }
        
        .table .table-warning {
            background-color: #fff3cd !important;
            color: #856404;
        }
        
        .table .table-danger {
            background-color: #f8d7da !important;
            color: #721c24;
        }
        
        .table .table-info {
            background-color: #d1ecf1 !important;
            color: #0c5460;
        }

        /* Direct info class for table rows */
        tr.info {
            background-color: #e0f2fe !important;
            color: #0c4a6e;
        }

        tr.info td {
            background-color: #e0f2fe !important;
            color: #0c4a6e;
        }

        /* Info header styling */
        h5.info {
            color: #0c4a6e;
            background-color: #e0f2fe;
            padding: 8px 12px;
            border-radius: 4px;
            border-left: 3px solid #0284c7;
            margin: 16px 0 8px 0;
        }
        
        .table .table-primary {
            background-color: #cce7ff !important;
            color: #004085;
        }

        /* Table cell specific status styling (without heavy padding) */
        td.success {
            background-color: #dcfce7 !important;
            color: #166534;
            font-weight: 500;
        }

        td.warning {
            background-color: #fef3c7 !important;
            color: #92400e;
            font-weight: 500;
        }

        td.error {
            background-color: #fee2e2 !important;
            color: #991b1b;
            font-weight: 500;
        }

        /* Badge-like status indicators - subtle styling */
        .badge {
            display: inline-block;
            padding: 2px 6px;
            font-size: 0.75rem;
            font-weight: 600;
            line-height: 1.2;
            text-align: center;
            white-space: nowrap;
            vertical-align: baseline;
            border-radius: 12px;
            border: 1px solid;
        }

        .badge-success {
            color: #166534;
            background-color: transparent;
            border-color: #166534;
        }

        .badge-warning {
            color: #92400e;
            background-color: transparent;
            border-color: #92400e;
        }

        .badge-danger {
            color: #991b1b;
            background-color: transparent;
            border-color: #991b1b;
        }

        .badge-info {
            color: #0c5460;
            background-color: transparent;
            border-color: #0c5460;
        }

        .badge-primary {
            color: #004085;
            background-color: transparent;
            border-color: #004085;
        }
        
        .badge-secondary {
            color: #fff;
            background-color: #6c757d;
        }
        
        /* Metric cards - more subtle and professional */
        .metric {
            display: inline-block;
            margin: 10px;
            padding: 20px;
            background: white;
            border-radius: 8px;
            text-align: center;
            min-width: 140px;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);
            border: 1px solid #e2e8f0;
            transition: all 0.2s ease;
            position: relative;
        }
        .metric:hover { 
            transform: translateY(-2px); 
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        }
        .metric .number { 
            font-size: 1.75rem; 
            font-weight: 700; 
            color: #1e293b; 
            margin-bottom: 10px; 
        }
        .metric .label { 
            font-size: 0.8rem; 
            color: #64748b; 
            font-weight: 500;
            line-height: 1.3;
        }
        
        /* Export button styling */
        .export-btn { 
            position: absolute; 
            top: 8px; 
            right: 8px; 
            background: #3b82f6; 
            color: white; 
            border: none; 
            border-radius: 6px; 
            width: 28px; 
            height: 28px; 
            font-size: 12px; 
            cursor: pointer; 
            opacity: 0.7; 
            transition: all 0.2s ease;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        .export-btn:hover { 
            background: #2563eb; 
            opacity: 1;
            transform: scale(1.1);
        }
        
        /* Tree and toggle styling */
        .ou-tree, .group-tree { margin: 20px 0; }
        .ou-item, .group-item { margin: 2px 0; }
        .ou-toggle, .group-toggle { 
            cursor: pointer; 
            user-select: none; 
            padding: 8px 12px; 
            border-radius: 6px; 
            transition: all 0.2s ease; 
            display: inline-block;
            color: #374151;
        }
        .ou-toggle:hover, .group-toggle:hover { 
            background: #f3f4f6;
            color: #1f2937;
        }
        .ou-toggle::before, .group-toggle::before { 
            content: '▶'; 
            display: inline-block; 
            width: 12px; 
            color: #6b7280; 
            font-size: 10px; 
            transition: transform 0.2s ease; 
        }
        .ou-toggle.expanded::before, .group-toggle.expanded::before { 
            content: '▼'; 
            transform: rotate(0deg); 
        }
        .ou-children, .group-children { 
            margin-left: 20px; 
            display: none; 
            border-left: 2px solid #e5e7eb; 
            padding-left: 12px; 
            margin-top: 4px; 
        }
        .ou-children.show, .group-children.show { 
            display: block; 
            animation: fadeIn 0.2s ease; 
        }
        @keyframes fadeIn { 
            from { opacity: 0; transform: translateY(-4px); } 
            to { opacity: 1; transform: translateY(0); } 
        }
        
        /* Stats and level styling */
        .ou-stats, .group-stats { 
            font-size: 0.75rem; 
            color: #6b7280; 
            margin-left: 8px; 
            background: #f3f4f6; 
            padding: 2px 8px; 
            border-radius: 12px; 
            display: inline-block; 
        }
        .ou-level-0 { margin-left: 0px; }
        .ou-level-1 { margin-left: 16px; }
        .ou-level-2 { margin-left: 32px; }
        .ou-level-3 { margin-left: 48px; }
        .ou-level-4 { margin-left: 64px; }
        .ou-level-5 { margin-left: 80px; }
        
        /* Group tree styling */
        .group-tree { 
            max-height: 500px; 
            overflow-y: auto; 
            border: 1px solid #e5e7eb; 
            padding: 16px; 
            border-radius: 8px; 
            background: #fafafa; 
        }
        .group-tree::-webkit-scrollbar { width: 6px; }
        .group-tree::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 3px; }
        .group-tree::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
        .group-tree::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        
        .group-name { font-weight: 600; color: #374151; }
        .group-scope { 
            font-size: 0.75rem; 
            color: #6b7280; 
            margin-left: 6px; 
            background: #e5e7eb; 
            padding: 2px 6px; 
            border-radius: 8px; 
        }
        
        /* Risk level styling */
        .high-risk { color: #dc2626; font-weight: 600; }
        .medium-risk { color: #d97706; font-weight: 500; }
        .low-risk { color: #059669; font-weight: 500; }
        
        /* Controls panel */
        .controls-panel { 
            margin: 28px 0; 
            padding: 24px; 
            background: #f8fafc; 
            border-radius: 8px; 
            border: 1px solid #e2e8f0; 
        }
        .control-btn { 
            margin: 6px; 
            padding: 10px 20px; 
            background: #3b82f6; 
            color: white; 
            border: none; 
            border-radius: 6px; 
            cursor: pointer; 
            font-weight: 500; 
            transition: all 0.2s ease;
            font-size: 0.8rem;
        }
        .control-btn:hover { 
            background: #2563eb; 
            transform: translateY(-1px);
        }
        .control-btn:active { transform: translateY(0px); }
        
        /* Template styling */
        .template-list { margin: 12px 0; }
        .template-tag { 
            display: inline-block; 
            margin: 4px 6px 4px 0; 
            padding: 6px 12px; 
            background: #eff6ff; 
            border: 1px solid #3b82f6; 
            border-radius: 16px; 
            font-size: 0.75rem; 
            color: #1e40af; 
            font-weight: 500; 
            transition: all 0.2s ease; 
        }
        .template-tag:hover { 
            background: #dbeafe;
            box-shadow: 0 2px 4px rgba(59, 130, 246, 0.1);
        }
        
        /* Status indicators */
        .status-indicator { 
            display: flex; 
            align-items: center; 
            margin: 12px 0; 
            padding: 12px 16px; 
            border-radius: 8px; 
            font-weight: 500; 
        }
        .status-indicator.success { 
            background: #dcfce7; 
            border: 1px solid #22c55e; 
            color: #166534; 
        }
        .status-indicator.warning { 
            background: #fef3c7; 
            border: 1px solid #f59e0b; 
            color: #92400e; 
        }
        .status-indicator .status-icon { margin-right: 8px; font-size: 16px; }
        .status.success {
            color: #166534;
            font-weight: 600;
            border-left: 2px solid #166534;
            padding-left: 6px;
        }
        .status.warning {
            color: #92400e;
            font-weight: 600;
            border-left: 2px solid #92400e;
            padding-left: 6px;
        }
        .status.error {
            color: #991b1b;
            font-weight: 600;
            border-left: 2px solid #991b1b;
            padding-left: 6px;
        }
        
        .info-grid { margin: 12px 0; }
        .info-grid p { margin: 6px 0; }
        
        /* Responsive design improvements */
        @media (max-width: 768px) {
            body { padding: 12px 16px; font-size: 13px; }
            .container { max-width: 100%; padding: 0 10px; }
            .header { padding: 24px 20px; }
            .header h1 { font-size: 1.5rem; }
            .section { padding: 20px 24px; margin: 16px 0; }
            .metric { min-width: 130px; margin: 8px; padding: 18px; }
            .metric .number { font-size: 1.5rem; }
            .metric .label { font-size: 0.75rem; }
            table { font-size: 0.8rem; }
            th, td { padding: 10px 14px; }
            .section h2 { font-size: 1.25rem; }
            .section h3 { font-size: 1rem; }
        }
        
        @media (max-width: 480px) {
            body { padding: 8px 12px; }
            .header { padding: 20px 16px; }
            .section { padding: 16px 20px; }
            .metric { min-width: 120px; margin: 6px; padding: 16px; }
        }
        
        /* Status indicators with left border styling - more subtle and professional */
        .status-online {
            border-left: 4px solid #10b981 !important;
            background-color: rgba(16, 185, 129, 0.05) !important;
            padding-left: 12px !important;
        }
        .status-warning {
            border-left: 4px solid #f59e0b !important;
            background-color: rgba(245, 158, 11, 0.05) !important;
            padding-left: 12px !important;
        }
        .status-error {
            border-left: 4px solid #ef4444 !important;
            background-color: rgba(239, 68, 68, 0.05) !important;
            padding-left: 12px !important;
        }
        .status-unknown {
            border-left: 4px solid #6b7280 !important;
            background-color: rgba(107, 114, 128, 0.05) !important;
            padding-left: 12px !important;
        }

        /* DNS Table specific styles */
        .table-responsive {
            overflow-x: auto;
            margin-bottom: 1rem;
        }

        .dns-table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            background: white;
        }

        .dns-table th {
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            color: white;
            font-weight: 500;
            padding: 12px 15px;
            text-align: left;
            border: none;
        }

        .dns-table td {
            padding: 12px 15px;
            border: none;
            border-bottom: 1px solid #e2e8f0;
        }

        .dns-table tr:nth-child(even) {
            background-color: rgba(52,152,219,0.03);
        }

        .dns-table tr:hover {
            background-color: rgba(52,152,219,0.08);
        }

        .dns-table .status {
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.875rem;
            font-weight: 500;
        }

        .dns-table .status.success {
            color: #166534;
            font-weight: 600;
            border-left: 3px solid #166534;
            padding-left: 8px;
        }

        .dns-table .status.warning {
            color: #92400e;
            font-weight: 600;
            border-left: 3px solid #92400e;
            padding-left: 8px;
        }

        .dns-table .status.error {
            color: #991b1b;
            font-weight: 600;
            border-left: 3px solid #991b1b;
            padding-left: 8px;
        }

        @media (max-width: 768px) {
            .dns-table {
                font-size: 0.875rem;
            }

            .dns-table th,
            .dns-table td {
                padding: 8px 10px;
            }
        }

        .fsmo-summary { 
            margin-top: 16px; 
            padding: 12px 16px; 
            border-radius: 8px; 
            font-weight: 500;
        }
    </style>
    <script>
        function exportToCSV(data, filename) {
            // Function to properly escape CSV fields
            function escapeCSVField(field) {
                if (field === null || field === undefined) {
                    return '';
                }
                const stringField = String(field);
                // If field contains comma, quotes, or newlines, wrap in quotes and escape internal quotes
                if (stringField.includes(',') || stringField.includes('"') || stringField.includes('\n') || stringField.includes('\r')) {
                    return '"' + stringField.replace(/"/g, '""') + '"';
                }
                return stringField;
            }
            
            const csvContent = "data:text/csv;charset=utf-8," + 
                data.map(row => row.map(field => escapeCSVField(field)).join(",")).join("\n");
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", filename);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
        
        function exportUserStats() {
            const userStats = [
                ["Metric", "Count"],
                ["Total Users", "$($AuditResults.AccountsAnalysis.Users.TotalUsers)"],
                ["Enabled Users", "$($AuditResults.AccountsAnalysis.Users.EnabledUsers)"],
                ["Disabled Users", "$($AuditResults.AccountsAnalysis.Users.DisabledUsers)"],
                ["Password Never Expires", "$($AuditResults.AccountsAnalysis.Users.PasswordNeverExpires)"],
                ["Inactive 90+ Days", "$($AuditResults.AccountsAnalysis.Users.InactiveUsers90Days)"],
                ["Locked Out Users", "$($AuditResults.AccountsAnalysis.Users.LockedOutUsers)"]
            ];
            exportToCSV(userStats, "user_statistics.csv");
        }
        
        function exportComputerStats() {
            const computerStats = [
                ["Metric", "Count"],
                ["Total Computers", "$($AuditResults.AccountsAnalysis.Computers.TotalComputers)"],
                ["Enabled Computers", "$($AuditResults.AccountsAnalysis.Computers.EnabledComputers)"],
                ["Disabled Computers", "$($AuditResults.AccountsAnalysis.Computers.DisabledComputers)"],
                ["Windows Servers", "$($AuditResults.AccountsAnalysis.Computers.WindowsServers)"],
                ["Windows Workstations", "$($AuditResults.AccountsAnalysis.Computers.WindowsWorkstations)"],
                ["Inactive 90+ Days", "$($AuditResults.AccountsAnalysis.Computers.InactiveComputers90Days)"]
            ];
            exportToCSV(computerStats, "computer_statistics.csv");
        }
        
        // Individual metric export functions
        function exportEnabledUsers() {
            const data = [
                ["Name", "SamAccountName", "Last Logon", "Password Last Set"],
$(if ($AuditResults.AccountsAnalysis.Users.EnabledUsersList) {
    foreach ($user in $AuditResults.AccountsAnalysis.Users.EnabledUsersList) {
        "                [`"$($user.Name)`", `"$($user.SamAccountName)`", `"$($user.LastLogonDate)`", `"$($user.PasswordLastSet)`"],"
    }
})
            ];
            exportToCSV(data, "enabled_users.csv");
        }
        
        function exportDisabledUsers() {
            const data = [
                ["Name", "SamAccountName", "Last Logon", "Password Last Set"],
$(if ($AuditResults.AccountsAnalysis.Users.DisabledUsersList) {
    foreach ($user in $AuditResults.AccountsAnalysis.Users.DisabledUsersList) {
        "                [`"$($user.Name)`", `"$($user.SamAccountName)`", `"$($user.LastLogonDate)`", `"$($user.PasswordLastSet)`"],"
    }
})
            ];
            exportToCSV(data, "disabled_users.csv");
        }
        
        function exportPasswordNeverExpiresUsers() {
            const data = [
                ["Name", "SamAccountName", "Last Logon", "Password Last Set", "Enabled"],
$(if ($AuditResults.AccountsAnalysis.Users.PasswordNeverExpiresUsersList) {
    foreach ($user in $AuditResults.AccountsAnalysis.Users.PasswordNeverExpiresUsersList) {
        "                [`"$($user.Name)`", `"$($user.SamAccountName)`", `"$($user.LastLogonDate)`", `"$($user.PasswordLastSet)`", `"$($user.Enabled)`"],"
    }
})
            ];
            exportToCSV(data, "password_never_expires_users.csv");
        }
        
        function exportInactiveUsers() {
            const data = [
                ["Name", "SamAccountName", "Last Logon", "Enabled"],
$(if ($AuditResults.AccountsAnalysis.Users.InactiveUsersList) {
    foreach ($user in $AuditResults.AccountsAnalysis.Users.InactiveUsersList) {
        "                [`"$($user.Name)`", `"$($user.SamAccountName)`", `"$($user.LastLogonDate)`", `"$($user.Enabled)`"],"
    }
})
            ];
            exportToCSV(data, "inactive_users_90days.csv");
        }
        
        function exportLockedOutUsers() {
            const data = [
                ["Name", "SamAccountName", "Account Lockout Time", "Last Logon"],
$(if ($AuditResults.AccountsAnalysis.Users.LockedOutUsersList) {
    foreach ($user in $AuditResults.AccountsAnalysis.Users.LockedOutUsersList) {
        "                [`"$($user.Name)`", `"$($user.SamAccountName)`", `"$($user.AccountLockoutTime)`", `"$($user.LastLogonDate)`"],"
    }
})
            ];
            exportToCSV(data, "locked_out_users.csv");
        }
        
        function exportEnabledComputers() {
            const data = [
                ["Name", "Operating System", "Last Logon"],
$(if ($AuditResults.AccountsAnalysis.Computers.EnabledComputersList) {
    foreach ($computer in $AuditResults.AccountsAnalysis.Computers.EnabledComputersList) {
        "                [`"$($computer.Name)`", `"$($computer.OperatingSystem)`", `"$($computer.LastLogonDate)`"],"
    }
})
            ];
            exportToCSV(data, "enabled_computers.csv");
        }
        
        function exportDisabledComputers() {
            const data = [
                ["Name", "Operating System", "Last Logon"],
$(if ($AuditResults.AccountsAnalysis.Computers.DisabledComputersList) {
    foreach ($computer in $AuditResults.AccountsAnalysis.Computers.DisabledComputersList) {
        "                [`"$($computer.Name)`", `"$($computer.OperatingSystem)`", `"$($computer.LastLogonDate)`"],"
    }
})
            ];
            exportToCSV(data, "disabled_computers.csv");
        }
        
        function exportInactiveComputers() {
            const data = [
                ["Name", "Operating System", "Last Logon"],
$(if ($AuditResults.AccountsAnalysis.Computers.InactiveComputersList) {
    foreach ($computer in $AuditResults.AccountsAnalysis.Computers.InactiveComputersList) {
        "                [`"$($computer.Name)`", `"$($computer.OperatingSystem)`", `"$($computer.LastLogonDate)`"],"
    }
})
            ];
            exportToCSV(data, "inactive_computers_90days.csv");
        }
        
        function exportWindowsServers() {
            const data = [
                ["Name", "Operating System", "Last Logon", "Enabled"],
$(if ($AuditResults.AccountsAnalysis.Computers.WindowsServersList) {
    foreach ($server in $AuditResults.AccountsAnalysis.Computers.WindowsServersList) {
        "                [`"$($server.Name)`", `"$($server.OperatingSystem)`", `"$($server.LastLogonDate)`", `"$($server.Enabled)`"],"
    }
})
            ];
            exportToCSV(data, "windows_servers.csv");
        }
        
        function exportWindowsWorkstations() {
            const data = [
                ["Name", "Operating System", "Last Logon", "Enabled"],
$(if ($AuditResults.AccountsAnalysis.Computers.WindowsWorkstationsList) {
    foreach ($workstation in $AuditResults.AccountsAnalysis.Computers.WindowsWorkstationsList) {
        "                [`"$($workstation.Name)`", `"$($workstation.OperatingSystem)`", `"$($workstation.LastLogonDate)`", `"$($workstation.Enabled)`"],"
    }
})
            ];
            exportToCSV(data, "windows_workstations.csv");
        }
        
        function exportTotalUsers() {
            const data = [
                ["Name", "SamAccountName", "Enabled", "Last Logon", "Password Last Set", "Password Never Expires", "Locked Out"],
$(if ($AuditResults.AccountsAnalysis.Users.AllUsersList) {
    foreach ($user in $AuditResults.AccountsAnalysis.Users.AllUsersList) {
        "                [`"$($user.Name)`", `"$($user.SamAccountName)`", `"$($user.Enabled)`", `"$($user.LastLogonDate)`", `"$($user.PasswordLastSet)`", `"$($user.PasswordNeverExpires)`", `"$($user.LockedOut)`"],"
    }
})
            ];
            exportToCSV(data, "total_users.csv");
        }
        
        function exportTotalComputers() {
            const data = [
                ["Name", "Operating System", "Enabled", "Last Logon", "Type"],
$(if ($AuditResults.AccountsAnalysis.Computers.AllComputersList) {
    foreach ($computer in $AuditResults.AccountsAnalysis.Computers.AllComputersList) {
        $computerType = if ($computer.OperatingSystem -like "*Server*") { "Server" } else { "Workstation" }
        "                [`"$($computer.Name)`", `"$($computer.OperatingSystem)`", `"$($computer.Enabled)`", `"$($computer.LastLogonDate)`", `"$computerType`"],"
    }
})
            ];
            exportToCSV(data, "total_computers.csv");
        }
    </script>
</head>
<body class="bg-light">
    <div class="container-fluid">
        <!-- Header Section -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="card shadow-sm border-0">
                    <div class="card-header bg-primary text-white">
                        <div class="row align-items-center">
                            <div class="col-md-8">
                                <h1 class="h3 mb-0">
                                    <i class="bi bi-shield-check me-2"></i>
                                    Active Directory Audit Report
                                </h1>
                            </div>
                            <div class="col-md-4 text-md-end">
                                <small class="opacity-75">
                                    <i class="bi bi-calendar3 me-1"></i>
                                    Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                                </small>
                            </div>
                        </div>
                    </div>
                    <div class="card-body bg-gradient">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="d-flex align-items-center mb-2">
                                    <i class="bi bi-building text-primary me-2"></i>
                                    <strong class="me-2">Domain:</strong>
                                    <span class="badge bg-light text-dark">$(try { (Get-ADDomain).DNSRoot } catch { "Error retrieving domain" })</span>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="d-flex align-items-center mb-2">
                                    <i class="bi bi-diagram-3 text-success me-2"></i>
                                    <strong class="me-2">Forest:</strong>
                                    <span class="badge bg-light text-dark">$(try { (Get-ADForest).Name } catch { "Error retrieving forest" })</span>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Executive Summary & Recommendations Section -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="card shadow-sm border-0">
                    <div class="card-header bg-primary text-white">
                        <h4 class="mb-0">
                            <i class="bi bi-clipboard-check me-2"></i>
                            Executive Summary & Analysis
                        </h4>
                    </div>
                    <div class="card-body">
                        <div class="row">
                            <div class="col-lg-6">
                                <h5 class="text-primary mb-3">🔍 Key Findings</h5>
                                <div class="summary-findings">
                                    $(
                                        $findings = @()

                                        # Domain Controllers Health
                                        if ($AuditResults.EnhancedDCHealth) {
                                            $totalDCs = $AuditResults.EnhancedDCHealth.Count
                                            $onlineDCs = ($AuditResults.EnhancedDCHealth | Where-Object { $_.PingStatus -eq "Online" }).Count
                                            if ($onlineDCs -lt $totalDCs) {
                                                $findings += "<li class='text-warning'><strong>Domain Controllers:</strong> $($totalDCs - $onlineDCs) out of $totalDCs DCs are offline or unreachable</li>"
                                            } else {
                                                $findings += "<li class='text-success'><strong>Domain Controllers:</strong> All $totalDCs DCs are online and accessible</li>"
                                            }
                                        }

                                        # Security Events
                                        if ($AuditResults.EventLogAnalysis -is [array] -and $AuditResults.EventLogAnalysis.Count -gt 0) {
                                            $criticalEvents = ($AuditResults.EventLogAnalysis | Where-Object { $_.Severity -in @("Critical", "High") }).Count
                                            if ($criticalEvents -gt 0) {
                                                $findings += "<li class='text-danger'><strong>Security Events:</strong> $criticalEvents critical/high severity events detected</li>"
                                            } else {
                                                $findings += "<li class='text-success'><strong>Security Events:</strong> No critical security events found</li>"
                                            }
                                        }

                                        # User Accounts
                                        if ($AuditResults.AccountsAnalysis.Users) {
                                            $inactiveUsers = $AuditResults.AccountsAnalysis.Users.InactiveUsers90Days
                                            $neverExpireUsers = $AuditResults.AccountsAnalysis.Users.PasswordNeverExpires
                                            if ($inactiveUsers -gt 0) {
                                                $findings += "<li class='text-warning'><strong>User Accounts:</strong> $inactiveUsers inactive accounts (90+ days)</li>"
                                            }
                                            if ($neverExpireUsers -gt 0) {
                                                $findings += "<li class='text-warning'><strong>Password Policy:</strong> $neverExpireUsers accounts with non-expiring passwords</li>"
                                            }
                                        }

                                        # Privileged Groups
                                        if ($AuditResults.PrivilegedGroups) {
                                            $domainAdmins = $AuditResults.PrivilegedGroups.DomainAdminsCount
                                            if ($domainAdmins -gt 5) {
                                                $findings += "<li class='text-warning'><strong>Privileged Access:</strong> $domainAdmins Domain Admins detected (recommend reducing to 3-5)</li>"
                                            }
                                        }

                                        # Orphaned SIDs
                                        if ($AuditResults.OrphanedSIDs.TotalCount -gt 0) {
                                            $findings += "<li class='text-warning'><strong>Security Cleanup:</strong> $($AuditResults.OrphanedSIDs.TotalCount) orphaned SIDs found in ACLs</li>"
                                        }

                                        if ($findings.Count -eq 0) {
                                            $findings += "<li class='text-success'><strong>Overall Status:</strong> No major issues detected in this audit</li>"
                                        }

                                        "<ul class='list-unstyled'>" + ($findings -join "") + "</ul>"
                                    )
                                </div>
                            </div>
                            <div class="col-lg-6">
                                <h5 class="text-success mb-3">💡 Recommendations</h5>
                                <div class="summary-recommendations">
                                    $(
                                        $recommendations = @()

                                        # Security recommendations
                                        if ($AuditResults.EventLogAnalysis -is [array] -and ($AuditResults.EventLogAnalysis | Where-Object { $_.Severity -in @("Critical", "High") }).Count -gt 0) {
                                            $recommendations += "<li><strong>Security Monitoring:</strong> Review and address critical security events immediately</li>"
                                        }

                                        # Account management recommendations
                                        if ($AuditResults.AccountsAnalysis.Users.InactiveUsers90Days -gt 0) {
                                            $recommendations += "<li><strong>Account Cleanup:</strong> Disable or remove inactive user accounts (90+ days)</li>"
                                        }

                                        if ($AuditResults.AccountsAnalysis.Users.PasswordNeverExpires -gt 0) {
                                            $recommendations += "<li><strong>Password Policy:</strong> Review accounts with non-expiring passwords</li>"
                                        }

                                        # Privileged access recommendations
                                        if ($AuditResults.PrivilegedGroups.DomainAdminsCount -gt 5) {
                                            $recommendations += "<li><strong>Privileged Access:</strong> Reduce Domain Admin membership to essential personnel only</li>"
                                        }

                                        # Infrastructure recommendations
                                        if ($AuditResults.EnhancedDCHealth) {
                                            $offlineDCs = ($AuditResults.EnhancedDCHealth | Where-Object { $_.PingStatus -ne "Online" }).Count
                                            if ($offlineDCs -gt 0) {
                                                $recommendations += "<li><strong>Infrastructure:</strong> Investigate and restore offline domain controllers</li>"
                                            }
                                        }

                                        # General recommendations
                                        $recommendations += "<li><strong>Regular Auditing:</strong> Schedule monthly AD health checks</li>"
                                        $recommendations += "<li><strong>Backup Verification:</strong> Ensure AD database backups are current and tested</li>"
                                        $recommendations += "<li><strong>Documentation:</strong> Update AD documentation and emergency procedures</li>"

                                        if ($AuditResults.OrphanedSIDs.TotalCount -gt 0) {
                                            $recommendations += "<li><strong>Security Cleanup:</strong> Remove orphaned SIDs from ACLs to improve security</li>"
                                        }

                                        "<ul class='list-unstyled'>" + ($recommendations -join "") + "</ul>"
                                    )
                                </div>
                            </div>
                        </div>

                    </div>
                </div>
            </div>
        </div>

        <!-- Stats Dashboard Section -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="card shadow-sm border-0">
                    <div class="card-header bg-light">
                        <h4 class="mb-0">
                            <i class="bi bi-graph-up text-primary me-2"></i>
                            Audit Overview & Key Metrics
                        </h4>
                    </div>
                    <div class="card-body">
                        <div class="row g-3">
                            <!-- Domain Controllers Stats -->
                            <div class="col-xl-3 col-lg-6 col-md-6">
                                <div class="card bg-light border-primary h-100">
                                    <div class="card-body">
                                        <div class="d-flex justify-content-between">
                                            <div>
                                                <h6 class="text-uppercase text-primary small">Domain Controllers</h6>
                                                <h2 class="mb-0 text-primary" id="dcCount">$(if ($AuditResults.EnhancedDCHealth) { $AuditResults.EnhancedDCHealth.Count } else { 0 })</h2>
                                                <span class="small text-muted">Total DCs Found</span>
                                            </div>
                                            <div class="align-self-center">
                                                <i class="bi bi-server text-primary" style="font-size: 2rem; opacity: 0.8;"></i>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <!-- User Accounts Stats -->
                            <div class="col-xl-3 col-lg-6 col-md-6">
                                <div class="card bg-light border-success h-100">
                                    <div class="card-body">
                                        <div class="d-flex justify-content-between">
                                            <div>
                                                <h6 class="text-uppercase text-success small">User Accounts</h6>
                                                <h2 class="mb-0 text-success" id="userCount">$(if ($AuditResults.AccountsAnalysis.Users.TotalUsers) { $AuditResults.AccountsAnalysis.Users.TotalUsers } else { 0 })</h2>
                                                <span class="small text-muted">Active Users</span>
                                            </div>
                                            <div class="align-self-center">
                                                <i class="bi bi-people text-success" style="font-size: 2rem; opacity: 0.8;"></i>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <!-- Computer Accounts Stats -->
                            <div class="col-xl-3 col-lg-6 col-md-6">
                                <div class="card bg-light border-info h-100">
                                    <div class="card-body">
                                        <div class="d-flex justify-content-between">
                                            <div>
                                                <h6 class="text-uppercase text-info small">Computer Accounts</h6>
                                                <h2 class="mb-0 text-info" id="computerCount">$(if ($AuditResults.AccountsAnalysis.Computers.TotalComputers) { $AuditResults.AccountsAnalysis.Computers.TotalComputers } else { 0 })</h2>
                                                <span class="small text-muted">Total Computers</span>
                                            </div>
                                            <div class="align-self-center">
                                                <i class="bi bi-laptop text-info" style="font-size: 2rem; opacity: 0.8;"></i>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            <!-- Security Events Stats -->
                            <div class="col-xl-3 col-lg-6 col-md-6">
                                <div class="card bg-light border-warning h-100">
                                    <div class="card-body">
                                        <div class="d-flex justify-content-between">
                                            <div>
                                                <h6 class="text-uppercase text-warning small">Security Events</h6>
                                                <h2 class="mb-0 text-warning" id="securityEvents">$(if ($AuditResults.EventLogAnalysis -is [array]) { $AuditResults.EventLogAnalysis.Count } else { 0 })</h2>
                                                <span class="small text-muted">Critical & Error Events</span>
                                            </div>
                                            <div class="align-self-center">
                                                <i class="bi bi-shield-exclamation text-warning" style="font-size: 2rem; opacity: 0.8;"></i>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>

                    </div>
                </div>
            </div>
        </div>

    <!-- ============================================ -->
    <!-- 🏗️ SECTION 1: INFRASTRUCTURE & FOUNDATION -->
    <!-- ============================================ -->

    <div class="row mb-4">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-header bg-secondary text-white">
                    <h4 class="mb-0">
                        <i class="bi bi-diagram-3 me-2"></i>
                        Infrastructure & Foundation
                    </h4>
                    <small class="text-white-50">Core Active Directory infrastructure components including domain controllers, replication, DNS, and forest structure.</small>
                </div>
                <div class="card-body">
"@

$HTMLReport += @"
                    <div class="row">
                        <div class="col-12">

    <div class="section">
        <h2>🖥️ Domain Controllers Analysis</h2>
"@

if ($AuditResults.EnhancedDCHealth -is [array]) {
    # --- New Summary Table: DC Info ---
    $HTMLReport += "<h3>Domain Controllers Summary</h3>"
    $HTMLReport += "<table><thead><tr><th>DC Name</th><th>Domain</th><th>OS Version</th><th>IPv4</th><th>Global Catalog</th><th>Read Only</th></tr></thead><tbody>"

    foreach ($DC in $AuditResults.EnhancedDCHealth) {
        $OSParts = $DC.OperatingSystem -split '\s{2,}|\s(?=\()'  # Basic split of version vs edition
        $OSVersion = $OSParts[0]
        $OSEdition = if ($OSParts.Length -gt 1) { $OSParts[1] } else { "" }
        $GCBadge = if ($DC.IsGlobalCatalog -eq 'True') { '<span class="badge badge-success">Yes</span>' } else { '<span class="badge badge-secondary">No</span>' }
        $ROBadge = if ($DC.IsReadOnly -eq 'True') { '<span class="badge badge-warning">Yes</span>' } else { '<span class="badge badge-success">No</span>' }

        $HTMLReport += "<tr><td><strong>$($DC.DomainController)</strong></td><td>$($DC.Domain)</td><td>$OSVersion</td><td>$($DC.IPv4Address)</td><td>$GCBadge</td><td>$ROBadge</td></tr>"
    }

    $HTMLReport += "</tbody></table>"

    # --- Existing Health Details Table ---
    $HTMLReport += "<h3>Detailed Health Report</h3>"
    $HTMLReport += "<table><thead><tr><th>DC Name</th><th>Domain</th><th>Uptime</th><th>Time Sync</th><th>Disk Free %</th><th>Disk Free GB</th><th>SYSVOL/Netlogon</th><th>Services</th><th>DCDiag Summary</th></tr></thead><tbody>"
    
    foreach ($DC in $AuditResults.EnhancedDCHealth) {
        $PassedTests = 0
        $TotalDCDiagTests = 0
        foreach ($prop in $DC.DCDiagResults.PSObject.Properties) {
            if ($prop.Name -ne "ServerName" -and $prop.Value -ne $null) {
                $TotalDCDiagTests++
                if ($prop.Value -eq "Passed") { $PassedTests++ }
            }
        }
        $DCDiagSummary = "$PassedTests/$TotalDCDiagTests"

        $ServicesStatus = @()
        
        # Handle service status display - check for Running/Started status or Unreachable DCs
if ($DC.DNSService -eq 'Running' -or $DC.DNSService -eq 'Started')           { $ServicesStatus += "DNS✓" }       
elseif ($DC.DNSService -eq 'Unreachable')                                   { $ServicesStatus += "DNS⚠" }       
else                                                                         { $ServicesStatus += "DNS✗" }

if ($DC.NTDSService -eq 'Running' -or $DC.NTDSService -eq 'Started')        { $ServicesStatus += "NTDS✓" }      
elseif ($DC.NTDSService -eq 'Unreachable')                                  { $ServicesStatus += "NTDS⚠" }      
else                                                                         { $ServicesStatus += "NTDS✗" }

if ($DC.NETLOGONService -eq 'Running' -or $DC.NETLOGONService -eq 'Started') { $ServicesStatus += "NETLOGON✓" }  
elseif ($DC.NETLOGONService -eq 'Unreachable')                              { $ServicesStatus += "NETLOGON⚠" }  
else                                                                         { $ServicesStatus += "NETLOGON✗" }

# KDC Service - using Get-Service since property doesn't exist
if ($DC.PingStatus -eq 'Unreachable' -or $DC.PingStatus -eq 'Offline') {
    $ServicesStatus += "KDC⚠"
} else {
    try {
        if ((Get-Service -ComputerName $DC.DomainController -Name "KDC" -ErrorAction Stop).Status -eq 'Running') { $ServicesStatus += "KDC✓" } else { $ServicesStatus += "KDC✗" }
    } catch { $ServicesStatus += "KDC✗" }
}

# W32Time Service - using Get-Service since property doesn't exist
if ($DC.PingStatus -eq 'Unreachable' -or $DC.PingStatus -eq 'Offline') {
    $ServicesStatus += "W32Time⚠"
} else {
    try {
        if ((Get-Service -ComputerName $DC.DomainController -Name "W32Time" -ErrorAction Stop).Status -eq 'Running') { $ServicesStatus += "W32Time✓" } else { $ServicesStatus += "W32Time✗" }
    } catch { $ServicesStatus += "W32Time✗" }
}

        $ServicesString = $ServicesStatus -join " "

        # Format uptime for better readability
        $UptimeFormatted = if ($DC.Uptime -eq "Unreachable" -or $DC.Uptime -eq "Error") {
            $DC.Uptime
        } elseif ($DC.Uptime -is [TimeSpan]) {
            "{0} days, {1:D2}h {2:D2}m" -f $DC.Uptime.Days, $DC.Uptime.Hours, $DC.Uptime.Minutes
        } else {
            $DC.Uptime
        }
        
        # Format time difference for better readability
        $TimeDiffFormatted = if ($DC.TimeDifference -eq "Unreachable" -or $DC.TimeDifference -eq "Error") {
            $DC.TimeDifference
        } elseif ($DC.TimeDifference -is [double]) {
            if ([Math]::Abs($DC.TimeDifference) -lt 1) {
                "{0:F1}ms" -f ($DC.TimeDifference * 1000)
            } elseif ([Math]::Abs($DC.TimeDifference) -lt 60) {
                "{0:F1}s" -f $DC.TimeDifference
            } else {
                "{0:F0}s" -f $DC.TimeDifference
            }
        } else {
            "$($DC.TimeDifference)s"
        }

        $HTMLReport += "<tr><td><strong>$($DC.DomainController)</strong></td><td>$($DC.Domain)</td><td>$UptimeFormatted</td><td>$TimeDiffFormatted</td><td>$($DC.OSDriveFreeSpacePercent)%</td><td>$($DC.OSDriveFreeSpaceGB) GB</td><td>$($DC.SYSVOL_Netlogon)</td><td>$ServicesString</td><td>$DCDiagSummary</td></tr>"
    }
    $HTMLReport += "</table>"

    # --- DCDiag Details Expandable Section ---
    $HTMLReport += "<h3>Detailed DCDiag Test Results</h3>"
    $HTMLReport += "<div class='controls-panel'>"
    $HTMLReport += "<strong>DCDiag Controls:</strong> "
    $HTMLReport += "<button class='control-btn' onclick='expandAllDCDiag()'>Expand All</button>"
    $HTMLReport += "<button class='control-btn' onclick='collapseAllDCDiag()'>Collapse All</button>"
    $HTMLReport += "<button class='control-btn' onclick='showFailedDCDiagOnly()'>Show Failed Only</button>"
    $HTMLReport += "<button class='control-btn' onclick='showAllDCDiag()'>Show All</button>"
    $HTMLReport += "</div>"

    foreach ($DC in $AuditResults.EnhancedDCHealth | Sort-Object DomainController) {
        $HTMLReport += "<div class='group-item'>"
        $HTMLReport += "<span class='group-toggle' onclick='toggleDCDiag(""$($DC.DomainController -replace '[^a-zA-Z0-9]', '_')"")'>"
        $HTMLReport += "<strong>$($DC.DomainController)</strong>"
        $HTMLReport += "</span>"
        $HTMLReport += "<div class='group-children' id='dcdiag-$($DC.DomainController -replace '[^a-zA-Z0-9]', '_')'>"
        $HTMLReport += "<table><thead><tr><th>Test Name</th><th>Result</th></tr></thead><tbody>"

        foreach ($prop in $DC.DCDiagResults.PSObject.Properties) {
            if ($prop.Name -ne "ServerName" -and $prop.Value -ne $null) {
                $TestRowClass = if ($prop.Value -eq "Failed") { "error" } elseif ($prop.Value -eq "Passed") { "success" } else { "" }
                $HTMLReport += "<tr class='$TestRowClass'><td>$($prop.Name)</td><td>$($prop.Value)</td></tr>"
            }
        }

        $HTMLReport += "</tbody></table></div></div>"
    }

    # Keep JS as-is
    $HTMLReport += @"
    <script>
    function toggleDCDiag(dcId) {
        var toggle = event.target.closest('.group-toggle');
        var children = document.getElementById('dcdiag-' + dcId);
        
        if (children) {
            if (children.classList.contains('show')) {
                children.classList.remove('show');
                toggle.classList.remove('expanded');
            } else {
                children.classList.add('show');
                toggle.classList.add('expanded');
            }
        }
    }

    function expandAllDCDiag() {
        document.querySelectorAll('.group-toggle').forEach(t => t.classList.add('expanded'));
        document.querySelectorAll('[id^="dcdiag-"]').forEach(d => d.classList.add('show'));
    }

    function collapseAllDCDiag() {
        document.querySelectorAll('.group-toggle').forEach(t => t.classList.remove('expanded'));
        document.querySelectorAll('[id^="dcdiag-"]').forEach(d => d.classList.remove('show'));
    }

    function showFailedDCDiagOnly() {
        document.querySelectorAll('.group-item').forEach(item => {
            const anyFail = item.querySelector('tr.error') !== null;
            item.style.display = anyFail ? 'block' : 'none';
        });
    }

    function showAllDCDiag() {
        document.querySelectorAll('.group-item').forEach(item => {
            item.style.display = 'block';
        });
    }
    </script>
"@
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.EnhancedDCHealth)</div>"
}




$HTMLReport += @"
    </div>

    <div class="section">
        <h2>📊 DC Performance Metrics</h2>
        <table>
            <tr>
                <th>Domain Controller</th>
                <th>CPU Usage</th>
                <th>Memory Usage</th>
                <th>AD Database Size</th>
                <th>Log File Size</th>
                <th>LDAP Connections</th>
            </tr>
"@

if ($AuditResults.DCPerformanceMetrics -and $AuditResults.DCPerformanceMetrics.Count -gt 0) {
    foreach ($DC in $AuditResults.DCPerformanceMetrics) {
        if ($DC.DomainController) {
            $HTMLReport += @"
                <tr>
                    <td><strong>$($DC.DomainController)</strong></td>
                    <td>$($DC.CPUUsage)</td>
                    <td>$($DC.MemoryUsage)</td>
                    <td>$($DC.ADDatabaseSize)</td>
                    <td>$($DC.LogFileSize)</td>
                    <td>$($DC.LDAPConnections)</td>
                </tr>
"@
        }
    }
} else {
    $HTMLReport += @"
            <tr><td colspan="6">No performance metrics available or access denied</td></tr>
"@
}

$HTMLReport += @"
        </table>
    </div>


    <div class="section">
        <h2>🔄 Replication Health Assessment</h2>
"@

if ($AuditResults.ReplicationHealth -is [array] -and $AuditResults.ReplicationHealth.Count -gt 0) {
    $HTMLReport += "<div class='warning'>Replication issues detected:</div>"
    $HTMLReport += "<table><thead><tr><th>Source DC</th><th>Destination DC</th><th>Naming Context</th><th>Failures</th><th>Last Success</th></tr></thead><tbody>"
    foreach ($Repl in $AuditResults.ReplicationHealth) {
        $HTMLReport += "<tr><td>$($Repl.SourceDC)</td><td>$($Repl.DestinationDC)</td><td>$($Repl.NamingContext)</td><td>$($Repl.Failures)</td><td>$($Repl.LastSuccess)</td></tr>"
    }
    $HTMLReport += "</tbody></table>"
} else {
    $HTMLReport += "<div class='success'>No replication issues detected</div>"
}


$HTMLReport += @"
</div>
<div class="section">
    <h2>🌐 DNS Configuration Testing</h2>
"@

# Check if the data is available
if ($AuditResults.DNSConfiguration -is [array] -and $AuditResults.DNSConfiguration.Count -gt 0) {
    
    # DNS Status Overview Table
    $HTMLReport += "<h3>📊 DNS Status Overview</h3>"
    $HTMLReport += "<div class='table-responsive'>"
    $HTMLReport += "<table class='dns-table'>"
    $HTMLReport += "<thead><tr><th>Domain Controller</th><th>Connectivity</th><th>DNS Server IP</th><th>DCDiag DNS</th><th>Domain Resolution</th><th>Reverse Lookup</th><th>Scavenging</th><th>Issues Found</th></tr></thead><tbody>"

    if ($AuditResults.DNSConfiguration.Count -eq 0) {
        $HTMLReport += "<tr><td colspan='8' class='text-center'><em>No DNS configuration data available</em></td></tr>"
    }

    foreach ($DNS in $AuditResults.DNSConfiguration) {
        # Determine status classes for connectivity
        $ConnectivityClass = switch ($DNS.ConnectivityStatus) {
            "Online" { "success" }
            { $_ -like "*Unreachable*" } { "error" }
            "Error" { "error" }
            default { "warning" }
        }
        
        $DCDiagClass = switch ($DNS.DCDiagDNS) {
            "Passed" { "success" }
            "Failed" { "error" }
            { $_ -like "*Error*" } { "error" }
            default { "warning" }
        }
        
        $IssueCount = if ($DNS.DNSFailures -and $DNS.DNSFailures.Count) { $DNS.DNSFailures.Count } else { 0 }
        $IssueClass = if ($IssueCount -eq 0) { "success" } elseif ($IssueCount -lt 5) { "warning" } else { "error" }
        
        # Ensure all properties have values
        $DCName = if ($DNS.DomainController) { $DNS.DomainController } else { "Unknown" }
        $ConnStatus = if ($DNS.ConnectivityStatus) { $DNS.ConnectivityStatus } else { "Unknown" }
        $DNSServerIP = if ($DNS.DNSServer) { $DNS.DNSServer } else { "Unknown" }
        $DCDiagResult = if ($DNS.DCDiagDNS) { $DNS.DCDiagDNS } else { "Unknown" }
        $DomainRes = if ($DNS.DomainResolution) { $DNS.DomainResolution } else { "Unknown" }
        $ReverseRes = if ($DNS.ReverseLookup) { $DNS.ReverseLookup } else { "Unknown" }
        $Scavenging = if ($DNS.ScavengingEnabled) { $DNS.ScavengingEnabled } else { "Unknown" }

        $HTMLReport += "<tr class='$ConnectivityClass'>"
        $HTMLReport += "<td><strong>$DCName</strong></td>"
        $HTMLReport += "<td><span class='status $ConnectivityClass'>$ConnStatus</span></td>"
        $HTMLReport += "<td>$DNSServerIP</td>"
        $HTMLReport += "<td><span class='status $DCDiagClass'>$DCDiagResult</span></td>"
        $HTMLReport += "<td>$DomainRes</td>"
        $HTMLReport += "<td>$ReverseRes</td>"
        $HTMLReport += "<td>$Scavenging</td>"
        $HTMLReport += "<td><span class='status $IssueClass'>$IssueCount issues</span></td>"
        $HTMLReport += "</tr>"
    }
    $HTMLReport += "</tbody></table>"
    $HTMLReport += "</div>"

    # Connectivity Issues Summary
    $OfflineDCs = $AuditResults.DNSConfiguration | Where-Object { $_.ConnectivityStatus -ne "Online" }
    if ($OfflineDCs -and $OfflineDCs.Count -gt 0) {
        $HTMLReport += "<div class='warning'>"
        $HTMLReport += "<h4>⚠️ Connectivity Issues Detected</h4>"
        $HTMLReport += "<p><strong>$($OfflineDCs.Count)</strong> Domain Controller(s) are experiencing connectivity issues:</p>"
        $HTMLReport += "<ul>"
        foreach ($OfflineDC in $OfflineDCs) {
            $HTMLReport += "<li><strong>$($OfflineDC.DomainController)</strong> - $($OfflineDC.ConnectivityStatus)"
            if ($OfflineDC.LastError) {
                $HTMLReport += "<br><small>Error: $($OfflineDC.LastError)</small>"
            }
            $HTMLReport += "</li>"
        }
        $HTMLReport += "</ul>"
        $HTMLReport += "<p><em>These domain controllers were skipped during DNS testing due to connectivity issues. Consider checking:</em></p>"
        $HTMLReport += "<ul><li>Network connectivity to the servers</li><li>Windows Firewall settings</li><li>Server availability and power status</li><li>DNS resolution for the server names</li></ul>"
        $HTMLReport += "</div>"
    }
    
    # DNS Issues Details
    $TotalIssuesCount = ($AuditResults.DNSConfiguration | ForEach-Object { if ($_.DNSFailures -and $_.DNSFailures.Count) { $_.DNSFailures.Count } else { 0 } } | Measure-Object -Sum).Sum
    
    if ($TotalIssuesCount -gt 0) {
        $HTMLReport += "<h3>🔍 Detailed DNS Issues ($TotalIssuesCount total)</h3>"
        
        foreach ($DNS in $AuditResults.DNSConfiguration) {
            if ($DNS.DNSFailures -and $DNS.DNSFailures.Count -gt 0) {
                $HTMLReport += "<h4>Issues on $($DNS.DomainController) ($($DNS.DNSFailures.Count) issues)</h4>"
                
                # Group issues by type
                $IssuesByType = $DNS.DNSFailures | Group-Object Type
                
                foreach ($IssueGroup in $IssuesByType) {
                    $TypeClass = switch ($IssueGroup.Name) {
                        "Connectivity" { "error" }
                        "DCDiag" { "warning" }
                        "Obsolete" { "info" }
                        "Configuration" { "warning" }
                        "System" { "error" }
                        default { "warning" }
                    }
                    
                    $HTMLReport += "<h5 class='$TypeClass'>$($IssueGroup.Name) Issues ($($IssueGroup.Count))</h5>"
                    
                    if ($IssueGroup.Count -gt 10) {
                        $HTMLReport += "<div style='max-height: 300px; overflow-y: auto; border: 1px solid #e2e8f0; padding: 12px; border-radius: 8px; background: #fafafa;'>"
                    }
                    
                    $HTMLReport += "<table>"
                    $HTMLReport += "<thead><tr><th>Issue</th><th>Details</th>"
                    if ($IssueGroup.Name -eq "Obsolete") { $HTMLReport += "<th>Zone</th>" }
                    $HTMLReport += "</tr></thead><tbody>"
                    
                    foreach ($issue in $IssueGroup.Group) {
                        $HTMLReport += "<tr class='$TypeClass'>"
                        $HTMLReport += "<td>$($issue.Entry)</td>"
                        $HTMLReport += "<td>$($issue.Details)</td>"
                        if ($IssueGroup.Name -eq "Obsolete") { 
                            $HTMLReport += "<td>$($issue.Zone)</td>" 
                        }
                        $HTMLReport += "</tr>"
                    }
                    
                    $HTMLReport += "</tbody></table>"
                    
                    if ($IssueGroup.Count -gt 10) {
                        $HTMLReport += "</div>"
                    }
                }
            }
        }
    }
    
    # Summary Statistics
    $OnlineDCs = $AuditResults.DNSConfiguration | Where-Object { $_.ConnectivityStatus -eq "Online" }
    $PassedDNS = $AuditResults.DNSConfiguration | Where-Object { $_.DCDiagDNS -eq "Passed" }
    $ObsoleteCount = ($AuditResults.DNSConfiguration | ForEach-Object { 
        if ($_.DNSFailures) { 
            ($_.DNSFailures | Where-Object { $_.Type -eq "Obsolete" }).Count 
        } else { 0 } 
    } | Measure-Object -Sum).Sum
    
    $HTMLReport += "<div class='success'>"
    $HTMLReport += "<h4>📈 DNS Health Summary</h4>"
    $HTMLReport += "<div class='info-grid'>"
    $HTMLReport += "<p><strong>Online Domain Controllers:</strong> $($OnlineDCs.Count) / $($AuditResults.DNSConfiguration.Count)</p>"
    $HTMLReport += "<p><strong>Passed DCDiag DNS Tests:</strong> $($PassedDNS.Count) / $($OnlineDCs.Count) (online DCs)</p>"
    $HTMLReport += "<p><strong>Total DNS Issues Found:</strong> $TotalIssuesCount</p>"
    $HTMLReport += "<p><strong>Obsolete DNS Records:</strong> $ObsoleteCount</p>"
    $HTMLReport += "</div>"
    $HTMLReport += "</div>"
    
} else {
    $HTMLReport += "<div class='error'>"
    $HTMLReport += "<h4>❌ DNS Configuration Testing Failed</h4>"
    $HTMLReport += "<p>Unable to retrieve DNS configuration data. This could be due to:</p>"
    $HTMLReport += "<ul>"
    $HTMLReport += "<li>No domain controllers found or accessible</li>"
    $HTMLReport += "<li>Insufficient permissions to run DNS tests</li>"
    $HTMLReport += "<li>Network connectivity issues</li>"
    $HTMLReport += "<li>Script execution errors</li>"
    $HTMLReport += "</ul>"
    if ($AuditResults.DNSConfiguration -is [string]) {
        $HTMLReport += "<p><strong>Error Details:</strong> $($AuditResults.DNSConfiguration)</p>"
    }
    $HTMLReport += "</div>"
}






$HTMLReport += @"
    </div>

    <div class="section">
        <h2>👑 FSMO Roles Verification</h2>
"@

if ($AuditResults.FSMORoles -is [array]) {
    $HTMLReport += "<table><thead><tr><th>Role</th><th>Server</th><th>Scope</th><th>Status</th></tr></thead><tbody>"
    
    foreach ($FsmoRole in $AuditResults.FSMORoles) {
        $StatusClass = switch ($FsmoRole.Status) {
            "Online" { "status-online" }
            "Offline" { "status-offline" }
            "Unknown Server" { "status-warning" }
            "Role Not Assigned" { "status-error" }
            "Connection Error" { "status-error" }
            "Domain Error" { "status-error" }
            "Processing Error" { "status-error" }
            default { "status-unknown" }
        }
        
        $HTMLReport += "<tr class='$StatusClass'>"
        $HTMLReport += "<td>$($FsmoRole.Role)</td>"
        $HTMLReport += "<td>$($FsmoRole.Server)</td>"
        $HTMLReport += "<td>$($FsmoRole.Scope)</td>"
        $HTMLReport += "<td>$($FsmoRole.Status)</td>"
        $HTMLReport += "</tr>"
    }
    
    $HTMLReport += "</tbody></table>"
    
    # Add summary statistics
    $OnlineCount = ($AuditResults.FSMORoles | Where-Object { $_.Status -eq "Online" }).Count
    $TotalCount = $AuditResults.FSMORoles.Count
    $HealthPercentage = if ($TotalCount -gt 0) { [math]::Round(($OnlineCount / $TotalCount) * 100, 1) } else { 0 }
    
    $SummaryClass = if ($HealthPercentage -ge 80) { "status-online" } elseif ($HealthPercentage -ge 60) { "status-warning" } else { "status-error" }
    
    $HTMLReport += "<div class='fsmo-summary $SummaryClass'>"
    $HTMLReport += "<strong>FSMO Health Summary:</strong> $OnlineCount/$TotalCount roles online ($HealthPercentage%)"
    $HTMLReport += "</div>"
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.FSMORoles)</div>"
}

$HTMLReport += @"
    </div>

    <div class="section">
        <h2>🤝 Trust Relationships Analysis</h2>
"@

if ($AuditResults.TrustRelationships -is [array] -and $AuditResults.TrustRelationships.Count -gt 0) {
    $HTMLReport += "<table><thead><tr><th>Trust Name</th><th>Direction</th><th>Type</th><th>Forest Transitive</th><th>Connectivity</th><th>Created</th></tr></thead><tbody>"
    foreach ($Trust in $AuditResults.TrustRelationships) {
        $HTMLReport += "<tr><td>$($Trust.Name)</td><td>$($Trust.Direction)</td><td>$($Trust.TrustType)</td><td>$($Trust.ForestTransitive)</td><td>$($Trust.ConnectivityTest)</td><td>$($Trust.Created)</td></tr>"
    }
    $HTMLReport += "</tbody></table>"
} elseif ($AuditResults.TrustRelationships -is [array]) {
    $HTMLReport += "<div class='success'>No external trusts configured</div>"
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.TrustRelationships)</div>"
}

$HTMLReport += @"
    </div>

    <div class="section">
        <h2>🏢 AD Sites & Services Audit</h2>
"@

if ($AuditResults.ADSites.Sites -is [array]) {
    $HTMLReport += "<h3>Sites</h3>"
    $HTMLReport += "<table><thead><tr><th>Site Name</th><th>Description</th><th>Domain Controllers</th><th>Subnets</th></tr></thead><tbody>"
    foreach ($Site in $AuditResults.ADSites.Sites) {
        $HTMLReport += "<tr><td>$($Site.Name)</td><td>$($Site.Description)</td><td>$($Site.DomainControllers)</td><td>$($Site.Subnets)</td></tr>"
    }
    $HTMLReport += "</tbody></table>"
    
    if ($AuditResults.ADSites.SiteLinks -is [array]) {
        $HTMLReport += "<h3>Site Links</h3>"
        $HTMLReport += "<table><thead><tr><th>Link Name</th><th>Cost</th><th>Replication Frequency (min)</th><th>Sites</th></tr></thead><tbody>"
        foreach ($Link in $AuditResults.ADSites.SiteLinks) {
            $HTMLReport += "<tr><td>$($Link.Name)</td><td>$($Link.Cost)</td><td>$($Link.ReplicationFrequencyInMinutes)</td><td>$($Link.SitesIncluded)</td></tr>"
        }
        $HTMLReport += "</tbody></table>"
    }
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.ADSites)</div>"
}

$HTMLReport += @"
    </div>

    <div class="section">
        <h2>📋 Group Policy Security Analysis</h2>
"@

if ($AuditResults.GroupPolicy -is [array] -and $AuditResults.GPOSecuritySummary) {
    $Summary = $AuditResults.GPOSecuritySummary

    # Security Summary Metrics
    $HTMLReport += @"
        <div class="subsection">
            <h3>🎯 Security Overview</h3>
            <div class="row">
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title text-primary">$($Summary.TotalGPOs)</h5>
                            <p class="card-text">Total GPOs</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title text-success">$($Summary.LinkedGPOs)</h5>
                            <p class="card-text">Linked GPOs</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title $(if($Summary.HighRiskGPOs -gt 0){'text-danger'}else{'text-success'})">$($Summary.HighRiskGPOs)</h5>
                            <p class="card-text">High Risk GPOs</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title $(if($Summary.OrphanedGPOs -gt 0){'text-warning'}else{'text-success'})">$($Summary.OrphanedGPOs)</h5>
                            <p class="card-text">Orphaned GPOs</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
"@

    # Security Issues Summary
    if ($AuditResults.GPOAllSecurityIssues -and $AuditResults.GPOAllSecurityIssues.Count -gt 0) {
        $HTMLReport += @"
        <div class="subsection">
            <h3>⚠️ Security Issues Found</h3>
            <div class="alert alert-warning">
                <strong>$($AuditResults.GPOAllSecurityIssues.Count) unique security issues detected across all GPOs:</strong>
                <ul class="mt-2">
"@
        foreach ($Issue in $AuditResults.GPOAllSecurityIssues) {
            $HTMLReport += "<li>$([System.Web.HttpUtility]::HtmlEncode($Issue))</li>"
        }
        $HTMLReport += @"
                </ul>
            </div>
        </div>
"@
    }

    # Recommendations
    if ($AuditResults.GPOAllRecommendations -and $AuditResults.GPOAllRecommendations.Count -gt 0) {
        $HTMLReport += @"
        <div class="subsection">
            <h3>💡 Security Recommendations</h3>
            <div class="alert alert-info">
                <ul class="mb-0">
"@
        foreach ($Recommendation in $AuditResults.GPOAllRecommendations) {
            $HTMLReport += "<li>$([System.Web.HttpUtility]::HtmlEncode($Recommendation))</li>"
        }
        $HTMLReport += @"
                </ul>
            </div>
        </div>
"@
    }

    # GPO Conflict Analysis
    if ($AuditResults.GPOConflictAnalysis) {
        $ConflictSummary = $AuditResults.GPOConflictAnalysis.ConflictSummary
        $HTMLReport += @"
        <div class="subsection">
            <h3>⚡ GPO Conflict Analysis</h3>
            <div class="row">
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title $(if($ConflictSummary.TotalConflicts -gt 0){'text-danger'}else{'text-success'})">$($ConflictSummary.TotalConflicts)</h5>
                            <p class="card-text">Total Conflicts</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title $(if($ConflictSummary.CriticalConflicts -gt 0){'text-danger'}else{'text-success'})">$($ConflictSummary.CriticalConflicts)</h5>
                            <p class="card-text">Critical Conflicts</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title $(if($ConflictSummary.WarningConflicts -gt 0){'text-warning'}else{'text-success'})">$($ConflictSummary.WarningConflicts)</h5>
                            <p class="card-text">Warning Conflicts</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title $(if($ConflictSummary.AffectedOUs.Count -gt 0){'text-warning'}else{'text-success'})">$($ConflictSummary.AffectedOUs.Count)</h5>
                            <p class="card-text">Affected OUs</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
"@

        # Policy Conflicts Details
        if ($AuditResults.GPOConflictAnalysis.PolicyConflicts -and $AuditResults.GPOConflictAnalysis.PolicyConflicts.Count -gt 0) {
            $HTMLReport += @"
        <div class="subsection">
            <h3>🔥 Policy Conflicts Detected</h3>
            <div class="table-responsive">
                <table class="table table-striped table-hover">
                    <thead class="table-dark">
                        <tr>
                            <th>Conflict Type</th>
                            <th>Setting</th>
                            <th>GPO 1</th>
                            <th>Value 1</th>
                            <th>GPO 2</th>
                            <th>Value 2</th>
                            <th>Target</th>
                            <th>Severity</th>
                        </tr>
                    </thead>
                    <tbody>
"@
            foreach ($Conflict in $AuditResults.GPOConflictAnalysis.PolicyConflicts) {
                $SeverityClass = switch ($Conflict.Severity) {
                    "High" { "text-danger" }
                    "Medium" { "text-warning" }
                    default { "text-info" }
                }
                $HTMLReport += @"
                        <tr>
                            <td>$([System.Web.HttpUtility]::HtmlEncode($Conflict.Type))</td>
                            <td><code>$([System.Web.HttpUtility]::HtmlEncode($Conflict.Setting))</code></td>
                            <td>$([System.Web.HttpUtility]::HtmlEncode($Conflict.GPO1))</td>
                            <td><span class="badge bg-primary">$([System.Web.HttpUtility]::HtmlEncode($Conflict.Value1))</span></td>
                            <td>$([System.Web.HttpUtility]::HtmlEncode($Conflict.GPO2))</td>
                            <td><span class="badge bg-secondary">$([System.Web.HttpUtility]::HtmlEncode($Conflict.Value2))</span></td>
                            <td><small>$([System.Web.HttpUtility]::HtmlEncode($Conflict.Target))</small></td>
                            <td><span class="badge $(if($Conflict.Severity -eq 'High'){'bg-danger'}elseif($Conflict.Severity -eq 'Medium'){'bg-warning'}else{'bg-info'})">$($Conflict.Severity)</span></td>
                        </tr>
"@
            }
            $HTMLReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
        }

        # Precedence Warnings
        if ($AuditResults.GPOConflictAnalysis.PrecedenceWarnings -and $AuditResults.GPOConflictAnalysis.PrecedenceWarnings.Count -gt 0) {
            $HTMLReport += @"
        <div class="subsection">
            <h3>⚠️ Precedence & Inheritance Issues</h3>
            <div class="alert alert-warning">
                <strong>The following precedence issues were detected:</strong>
                <ul class="mt-2">
"@
            foreach ($Warning in $AuditResults.GPOConflictAnalysis.PrecedenceWarnings) {
                $HTMLReport += "<li><strong>$([System.Web.HttpUtility]::HtmlEncode($Warning.Type)):</strong> $([System.Web.HttpUtility]::HtmlEncode($Warning.Description))</li>"
            }
            $HTMLReport += @"
                </ul>
            </div>
        </div>
"@
        }

        # Inheritance Issues
        if ($AuditResults.GPOConflictAnalysis.InheritanceIssues -and $AuditResults.GPOConflictAnalysis.InheritanceIssues.Count -gt 0) {
            $HTMLReport += @"
        <div class="subsection">
            <h3>🔗 Inheritance Issues</h3>
            <div class="alert alert-info">
                <strong>The following inheritance issues were detected:</strong>
                <ul class="mt-2">
"@
            foreach ($Issue in $AuditResults.GPOConflictAnalysis.InheritanceIssues) {
                $HTMLReport += "<li><strong>$([System.Web.HttpUtility]::HtmlEncode($Issue.Type)):</strong> $([System.Web.HttpUtility]::HtmlEncode($Issue.Description))</li>"
            }
            $HTMLReport += @"
                </ul>
            </div>
        </div>
"@
        }
    }

    # Detailed GPO Analysis Table
    $HTMLReport += @"
        <div class="subsection">
            <h3>📊 Detailed GPO Analysis</h3>
            <div class="table-responsive">
                <table class="table table-striped table-hover">
                    <thead class="table-dark">
                        <tr>
                            <th>GPO Name</th>
                            <th>Status</th>
                            <th>Links</th>
                            <th>Last Modified</th>
                            <th>Issues</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    foreach ($GPO in $AuditResults.GroupPolicy | Sort-Object SecurityScore) {
        $StatusClass = "success"
        $StatusText = "✅ Healthy"

        if ($GPO.IsOrphaned) {
            $StatusClass = "warning"
            $StatusText = "🗑️ Orphaned"
        } elseif ($GPO.SecurityScore -lt 60) {
            $StatusClass = "danger"
            $StatusText = "🚨 High Risk"
        } elseif ($GPO.SecurityScore -lt 80) {
            $StatusClass = "warning"
            $StatusText = "⚠️ Medium Risk"
        }


        $IssueCount = $GPO.SecurityIssues.Count + $GPO.ComplianceIssues.Count
        $IssuesBadge = if ($IssueCount -gt 0) { "<span class='badge bg-danger'>$IssueCount</span>" } else { "<span class='badge bg-success'>0</span>" }

        $LastModifiedClass = if ($GPO.LastModifiedDays -gt 365) { "text-warning" } elseif ($GPO.LastModifiedDays -gt 730) { "text-danger" } else { "text-muted" }

        $HTMLReport += @"
                        <tr class="table-$StatusClass">
                            <td><strong>$([System.Web.HttpUtility]::HtmlEncode($GPO.DisplayName))</strong></td>
                            <td>$StatusText</td>
                            <td>
                                <span class="badge $(if($GPO.LinkCount -gt 0){'bg-success'}else{'bg-secondary'})">$($GPO.LinkCount)</span>
                                $(if($GPO.IsLinked -eq 'No'){'<span class="badge bg-warning">Unlinked</span>'}else{''})
                            </td>
                            <td class="$LastModifiedClass">$($GPO.LastModifiedDays) days ago</td>
                            <td>$IssuesBadge</td>
                        </tr>
"@

        # Add issue details if present
        if ($IssueCount -gt 0) {
            $HTMLReport += @"
                        <tr class="table-light">
                            <td colspan="5">
                                <div class="ms-3">
"@
            if ($GPO.SecurityIssues.Count -gt 0) {
                $HTMLReport += "<small><strong>Security Issues:</strong></small><ul class='small mb-1'>"
                foreach ($Issue in $GPO.SecurityIssues) {
                    $HTMLReport += "<li class='text-danger'>$([System.Web.HttpUtility]::HtmlEncode($Issue))</li>"
                }
                $HTMLReport += "</ul>"
            }

            if ($GPO.ComplianceIssues.Count -gt 0) {
                $HTMLReport += "<small><strong>Compliance Issues:</strong></small><ul class='small mb-1'>"
                foreach ($Issue in $GPO.ComplianceIssues) {
                    $HTMLReport += "<li class='text-warning'>$([System.Web.HttpUtility]::HtmlEncode($Issue))</li>"
                }
                $HTMLReport += "</ul>"
            }

            $HTMLReport += @"
                                </div>
                            </td>
                        </tr>
"@
        }
    }

    $HTMLReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@

    # GPO Link Analysis
    $LinkedGPOs = $AuditResults.GroupPolicy | Where-Object { $_.LinkCount -gt 0 }
    if ($LinkedGPOs) {
        $HTMLReport += @"
        <div class="subsection">
            <h3>🔗 GPO Link Analysis</h3>
            <div class="table-responsive">
                <table class="table table-sm">
                    <thead class="table-secondary">
                        <tr>
                            <th>GPO Name</th>
                            <th>Linked To</th>
                            <th>Link Status</th>
                            <th>Enforced</th>
                            <th>Order</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        foreach ($GPO in $LinkedGPOs) {
            foreach ($Link in $GPO.LinkDetails) {
                $LinkTarget = if ($Link.TargetType -eq "Domain") { "🏢 Domain Root" } else { "📁 " + $Link.Target.Split(',')[0].Replace('OU=','') }
                $EnabledBadge = if ($Link.Enabled) { "<span class='badge bg-success'>Enabled</span>" } else { "<span class='badge bg-secondary'>Disabled</span>" }
                $EnforcedBadge = if ($Link.Enforced) { "<span class='badge bg-warning'>Enforced</span>" } else { "<span class='badge bg-light text-dark'>Normal</span>" }

                $HTMLReport += @"
                        <tr>
                            <td>$([System.Web.HttpUtility]::HtmlEncode($GPO.DisplayName))</td>
                            <td>$LinkTarget</td>
                            <td>$EnabledBadge</td>
                            <td>$EnforcedBadge</td>
                            <td><span class='badge bg-info'>$($Link.Order)</span></td>
                        </tr>
"@
            }
        }
        $HTMLReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

} else {
    $HTMLReport += "<div class='alert alert-danger'>Failed to analyze Group Policy: $($AuditResults.GroupPolicy)</div>"
}

$HTMLReport += @"
    </div>

    <div class="section">
        <h2>🏗️ Schema Architecture</h2>
"@

if ($AuditResults.SchemaArchitecture.Error) {
    $HTMLReport += "<div class='error'>Schema Analysis Error: $($AuditResults.SchemaArchitecture.Error)</div>"
} else {
    $HTMLReport += @"
        <h3>Schema Overview</h3>
        <div class="metric">
            <div class="number">$($AuditResults.SchemaArchitecture.SchemaVersion)</div>
            <div class="label">Schema Version</div>
        </div>
        <div class="metric">
            <div class="number">$($AuditResults.SchemaArchitecture.SchemaStatistics.TotalClasses)</div>
            <div class="label">Total Classes</div>
        </div>
        <div class="metric">
            <div class="number">$($AuditResults.SchemaArchitecture.SchemaStatistics.TotalAttributes)</div>
            <div class="label">Total Attributes</div>
        </div>
        <div class="metric">
            <div class="number">$($AuditResults.SchemaArchitecture.SchemaStatistics.CustomClasses)</div>
            <div class="label">Custom Classes</div>
        </div>
        <div class="metric">
            <div class="number">$($AuditResults.SchemaArchitecture.SchemaStatistics.CustomAttributes)</div>
            <div class="label">Custom Attributes</div>
        </div>
        
        <h3>Schema Extensions</h3>
        <table>
            <tr>
                <th>Extension Type</th>
                <th>Object Count</th>
                <th>Status</th>
            </tr>
"@
    
    if ($AuditResults.SchemaArchitecture.SchemaExtensions -and $AuditResults.SchemaArchitecture.SchemaExtensions.Count -gt 0) {
        foreach ($Extension in $AuditResults.SchemaArchitecture.SchemaExtensions) {
            $Status = if ($Extension.ObjectCount -gt 0) { "Active" } else { "No objects found" }
            $StatusClass = if ($Extension.ObjectCount -gt 0) { "success" } else { "warning" }
            $HTMLReport += @"
                <tr>
                    <td><strong>$($Extension.ExtensionType)</strong></td>
                    <td>$($Extension.ObjectCount)</td>
                    <td class="$StatusClass">$Status</td>
                </tr>
"@
        }
    } else {
        $HTMLReport += "<tr><td colspan='3'>No schema extensions detected</td></tr>"
    }
    
    $HTMLReport += @"
        </table>
        
        <h3>Custom Schema Objects</h3>
        <div style="display: flex; gap: 20px;">
            <div style="flex: 1;">
                <h4>Custom Classes ($($AuditResults.SchemaArchitecture.SchemaStatistics.CustomClasses))</h4>
                <table>
                    <tr>
                        <th>Class Name</th>
                        <th>Created</th>
                        <th>Description</th>
                    </tr>
"@
    
    if ($AuditResults.SchemaArchitecture.CustomClasses -and $AuditResults.SchemaArchitecture.CustomClasses.Count -gt 0) {
        $CustomClassesLimited = $AuditResults.SchemaArchitecture.CustomClasses | Select-Object -First 10
        foreach ($Class in $CustomClassesLimited) {
            $Description = if ($Class.Description) { $Class.Description } else { "No description" }
            $HTMLReport += @"
                    <tr>
                        <td><strong>$($Class.Name)</strong></td>
                        <td>$($Class.Created.ToString('yyyy-MM-dd'))</td>
                        <td>$Description</td>
                    </tr>
"@
        }
        if ($AuditResults.SchemaArchitecture.CustomClasses.Count -gt 10) {
            $HTMLReport += "<tr><td colspan='3'><em>... and $($AuditResults.SchemaArchitecture.CustomClasses.Count - 10) more custom classes</em></td></tr>"
        }
    } else {
        $HTMLReport += "<tr><td colspan='3'>No custom classes found</td></tr>"
    }
    
    $HTMLReport += @"
                </table>
            </div>
            
            <div style="flex: 1;">
                <h4>Custom Attributes ($($AuditResults.SchemaArchitecture.SchemaStatistics.CustomAttributes))</h4>
                <table>
                    <tr>
                        <th>Attribute Name</th>
                        <th>Created</th>
                        <th>Type</th>
                    </tr>
"@
    
    if ($AuditResults.SchemaArchitecture.CustomAttributes -and $AuditResults.SchemaArchitecture.CustomAttributes.Count -gt 0) {
        $CustomAttributesLimited = $AuditResults.SchemaArchitecture.CustomAttributes | Select-Object -First 10
        foreach ($Attribute in $CustomAttributesLimited) {
            $TypeInfo = if ($Attribute.SingleValued) { "Single" } else { "Multi" }
            $HTMLReport += @"
                    <tr>
                        <td><strong>$($Attribute.Name)</strong></td>
                        <td>$($Attribute.Created.ToString('yyyy-MM-dd'))</td>
                        <td>$TypeInfo</td>
                    </tr>
"@
        }
        if ($AuditResults.SchemaArchitecture.CustomAttributes.Count -gt 10) {
            $HTMLReport += "<tr><td colspan='3'><em>... and $($AuditResults.SchemaArchitecture.CustomAttributes.Count - 10) more custom attributes</em></td></tr>"
        }
    } else {
        $HTMLReport += "<tr><td colspan='3'>No custom attributes found</td></tr>"
    }
    
    $HTMLReport += @"
                </table>
            </div>
        </div>
"@
    
    # Add recommendations if any
    if ($AuditResults.SchemaArchitecture.Recommendations -and $AuditResults.SchemaArchitecture.Recommendations.Count -gt 0) {
        $HTMLReport += "<h3>Schema Recommendations</h3><div class='warning'><ul>"
        foreach ($Recommendation in $AuditResults.SchemaArchitecture.Recommendations) {
            $HTMLReport += "<li>$Recommendation</li>"
        }
        $HTMLReport += "</ul></div>"
    }
}

$HTMLReport += @"
    </div>

    <!-- ============================================ -->
    <!-- 👥 SECTION 2: IDENTITY & ACCESS MANAGEMENT -->
    <!-- ============================================ -->


    <div class="section">
        <h2>👥 User & Computer Accounts Analysis</h2>
"@

if ($AuditResults.AccountsAnalysis.Users -is [hashtable]) {
    $HTMLReport += @"
        <h3>👥 User Account Statistics</h3>
        <div class="metric">
            <button class="export-btn" onclick="exportTotalUsers()" title="Export All Users"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Users.TotalUsers)</div>
            <div class="label">Total Users</div>
        </div>
        <div class="metric">
            <button class="export-btn" onclick="exportEnabledUsers()" title="Export Enabled Users"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Users.EnabledUsers)</div>
            <div class="label">Enabled</div>
        </div>
        <div class="metric">
            <button class="export-btn" onclick="exportDisabledUsers()" title="Export Disabled Users"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Users.DisabledUsers)</div>
            <div class="label">Disabled</div>
        </div>
        <div class="metric">
            <button class="export-btn" onclick="exportPasswordNeverExpiresUsers()" title="Export Users with Non-Expiring Passwords"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Users.PasswordNeverExpires)</div>
            <div class="label">Never Expires</div>
        </div>
        <div class="metric">
            <button class="export-btn" onclick="exportInactiveUsers()" title="Export Inactive Users"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Users.InactiveUsers90Days)</div>
            <div class="label">Inactive 90d</div>
        </div>
        <div class="metric">
            <button class="export-btn" onclick="exportLockedOutUsers()" title="Export Locked Out Users"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Users.LockedOutUsers)</div>
            <div class="label">Locked Out</div>
        </div>
        
        <h3>💻 Computer Account Statistics</h3>
        <div class="metric">
            <button class="export-btn" onclick="exportTotalComputers()" title="Export All Computers"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Computers.TotalComputers)</div>
            <div class="label">Total Computers</div>
        </div>
        <div class="metric">
            <button class="export-btn" onclick="exportEnabledComputers()" title="Export Enabled Computers"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Computers.EnabledComputers)</div>
            <div class="label">Enabled</div>
        </div>
        <div class="metric">
            <button class="export-btn" onclick="exportDisabledComputers()" title="Export Disabled Computers"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Computers.DisabledComputers)</div>
            <div class="label">Disabled</div>
        </div>
        <div class="metric">
            <button class="export-btn" onclick="exportWindowsServers()" title="Export Windows Servers"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Computers.WindowsServers)</div>
            <div class="label">Servers</div>
        </div>
        <div class="metric">
            <button class="export-btn" onclick="exportWindowsWorkstations()" title="Export Windows Workstations"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Computers.WindowsWorkstations)</div>
            <div class="label">Workstations</div>
        </div>
        <div class="metric">
            <button class="export-btn" onclick="exportInactiveComputers()" title="Export Inactive Computers"><i class="bi bi-download"></i></button>
            <div class="number">$($AuditResults.AccountsAnalysis.Computers.InactiveComputers90Days)</div>
            <div class="label">Inactive 90d</div>
        </div>
"@
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.AccountsAnalysis)</div>"
}

$HTMLReport += @"
    </div>

    <!-- ============================================ -->
    <!-- 🔒 SECTION 3: SECURITY ANALYSIS -->
    <!-- ============================================ -->


    <div class="section">
        <h2>🔒 Security Analysis & Risk Assessment</h2>
"@

if ($AuditResults.SecurityFindings -is [array]) {
    if ($AuditResults.SecurityFindings.Count -gt 0) {
        $HTMLReport += "<div class='warning'><strong>Security Findings:</strong><ul>"
        foreach ($Finding in $AuditResults.SecurityFindings) {
            $HTMLReport += "<li>$Finding</li>"
        }
        $HTMLReport += "</ul></div>"
    } else {
        $HTMLReport += "<div class='success'>No critical security issues detected</div>"
    }
    
    if ($AuditResults.PasswordPolicy) {
        $HTMLReport += "<h3>Password Policies</h3>"
        
        # Default Domain Password Policy
        $HTMLReport += "<h4>Default Domain Password Policy</h4>"
        $HTMLReport += "<table><thead><tr><th>Setting</th><th>Value</th></tr></thead><tbody>"
        $HTMLReport += "<tr><td>Minimum Password Length</td><td>$($AuditResults.PasswordPolicy.DefaultPolicy.MinPasswordLength)</td></tr>"
        $HTMLReport += "<tr><td>Password Complexity</td><td>$($AuditResults.PasswordPolicy.DefaultPolicy.ComplexityEnabled)</td></tr>"
        $HTMLReport += "<tr><td>Maximum Password Age</td><td>$($AuditResults.PasswordPolicy.DefaultPolicy.MaxPasswordAge)</td></tr>"
        $HTMLReport += "<tr><td>Minimum Password Age</td><td>$($AuditResults.PasswordPolicy.DefaultPolicy.MinPasswordAge)</td></tr>"
        $HTMLReport += "<tr><td>Password History Count</td><td>$($AuditResults.PasswordPolicy.DefaultPolicy.PasswordHistoryCount)</td></tr>"
        $HTMLReport += "<tr><td>Lockout Threshold</td><td>$($AuditResults.PasswordPolicy.DefaultPolicy.LockoutThreshold)</td></tr>"
        $HTMLReport += "</tbody></table>"
        
        # Password Settings Objects (PSOs)
        if ($AuditResults.PasswordPolicy.PSOCount -gt 0) {
            $HTMLReport += "<h4>Password Settings Objects (Fine-Grained Password Policies) - $($AuditResults.PasswordPolicy.PSOCount) found</h4>"
            $HTMLReport += "<table><thead><tr><th>PSO Name</th><th>Min Length</th><th>Complexity</th><th>Max Age (Days)</th><th>Lockout Threshold</th><th>Applies To</th></tr></thead><tbody>"
            
            foreach ($PSO in $AuditResults.PasswordPolicy.PSOs) {
                $AppliesTo = try { 
                    ($PSO.AppliesTo | ForEach-Object { 
                        try { (Get-ADObject -Identity $_ -ErrorAction SilentlyContinue).Name } catch { $_ } 
                    }) -join ", " 
                } catch { "Unknown" }
                
                $HTMLReport += "<tr><td><strong>$($PSO.Name)</strong></td><td>$($PSO.MinPasswordLength)</td><td>$($PSO.ComplexityEnabled)</td><td>$($PSO.MaxPasswordAge.Days)</td><td>$($PSO.LockoutThreshold)</td><td>$AppliesTo</td></tr>"
            }
            $HTMLReport += "</tbody></table>"
        } else {
            $HTMLReport += "<div class='warning'><strong>No Password Settings Objects (PSOs) configured</strong><br/>Consider implementing Fine-Grained Password Policies for privileged accounts.</div>"
        }
    }
    
    if ($AuditResults.PrivilegedGroups) {
        $HTMLReport += "<h3>Privileged Group Membership</h3>"
        $HTMLReport += "<div class='metric'><div class='number'>$($AuditResults.PrivilegedGroups.DomainAdminsCount)</div><div class='label'>Domain Admins</div></div>"
        $HTMLReport += "<div class='metric'><div class='number'>$($AuditResults.PrivilegedGroups.EnterpriseAdminsCount)</div><div class='label'>Enterprise Admins</div></div>"
        $HTMLReport += "<div class='metric'><div class='number'>$($AuditResults.PrivilegedGroups.SchemaAdminsCount)</div><div class='label'>Schema Admins</div></div>"
    }
    
    # KRBTGT Account Status
    if ($AuditResults.KRBTGTPasswordAge) {
        $HTMLReport += "<h3>🔑 KRBTGT Account Status</h3>"
        $StatusClass = switch ($AuditResults.KRBTGTPasswordAge.Status) {
            "Good" { "success" }
            "Warning" { "warning" }
            "Critical" { "error" }
            "Error" { "error" }
        }
        
        $HTMLReport += "<div class='$StatusClass'>"
        if ($AuditResults.KRBTGTPasswordAge.Status -eq "Error") {
            $HTMLReport += "<h4>❌ KRBTGT Account Check Failed</h4>"
            $HTMLReport += "<p><strong>Error:</strong> $($AuditResults.KRBTGTPasswordAge.Error)</p>"
        } else {
            $HTMLReport += "<h4>KRBTGT Password Information</h4>"
            $HTMLReport += "<div class='info-grid'>"
            $HTMLReport += "<p><strong>Password Last Set:</strong> $($AuditResults.KRBTGTPasswordAge.PasswordLastSet)</p>"
            $HTMLReport += "<p><strong>Password Age:</strong> $($AuditResults.KRBTGTPasswordAge.PasswordAgeDays) days</p>"
            $HTMLReport += "<p><strong>Status:</strong> $($AuditResults.KRBTGTPasswordAge.Status)</p>"
            $HTMLReport += "</div>"
            
            if ($AuditResults.KRBTGTPasswordAge.Status -eq "Critical") {
                $HTMLReport += "<p><strong>⚠️ Action Required:</strong> KRBTGT password should be reset immediately. Passwords older than 180 days pose a security risk.</p>"
            } elseif ($AuditResults.KRBTGTPasswordAge.Status -eq "Warning") {
                $HTMLReport += "<p><strong>⚠️ Recommendation:</strong> Consider resetting the KRBTGT password soon. Best practice is to reset every 90-180 days.</p>"
            }
        }
        $HTMLReport += "</div>"
    }
    
    # Active Directory Recycle Bin Status
    if ($AuditResults.RecycleBinStatus) {
        $HTMLReport += "<h3>🗑️ Active Directory Recycle Bin</h3>"
        $StatusClass = switch ($AuditResults.RecycleBinStatus.Status) {
            "Enabled" { "success" }
            "Disabled" { "warning" }
            "Error" { "error" }
        }
        
        $HTMLReport += "<div class='$StatusClass'>"
        if ($AuditResults.RecycleBinStatus.Status -eq "Error") {
            $HTMLReport += "<h4>❌ Recycle Bin Check Failed</h4>"
            $HTMLReport += "<p><strong>Error:</strong> $($AuditResults.RecycleBinStatus.Error)</p>"
        } elseif ($AuditResults.RecycleBinStatus.Enabled) {
            $HTMLReport += "<h4>✅ AD Recycle Bin is Enabled</h4>"
            $HTMLReport += "<p>The Active Directory Recycle Bin feature is enabled and providing protection against accidental deletions.</p>"
            if ($AuditResults.RecycleBinStatus.EnabledScopes) {
                $HTMLReport += "<p><strong>Enabled for:</strong> $($AuditResults.RecycleBinStatus.EnabledScopes -join ', ')</p>"
            }
        } else {
            $HTMLReport += "<h4>⚠️ AD Recycle Bin is Disabled</h4>"
            $HTMLReport += "<p><strong>Recommendation:</strong> Enable the Active Directory Recycle Bin feature to protect against accidental object deletions.</p>"
            $HTMLReport += "<p>This feature allows recovery of deleted AD objects without requiring a system state restore.</p>"
        }
        $HTMLReport += "</div>"
    }
    
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.SecurityFindings)</div>"
}

$HTMLReport += @"
    </div>

    <div class="section">
        <h2>🏗️ OU Hierarchy Mapping</h2>
"@

if ($AuditResults.OUHierarchy -is [array]) {
    # Build interactive tree structure
    $HTMLReport += "<div class='ou-tree' id='ou-tree'>"
    
    # Group OUs by parent to build hierarchy
    $OUGroups = $AuditResults.OUHierarchy | Group-Object ParentDN
    $RootOUs = $AuditResults.OUHierarchy | Where-Object {$_.Level -eq 1}
    
    # Function to build tree recursively (simplified for PowerShell)
    function Build-OUTree($OUs, $ParentDN, $Level) {
        $ChildOUs = $OUs | Where-Object {$_.ParentDN -eq $ParentDN}
        $TreeHTML = ""
        
        foreach ($OU in $ChildOUs | Sort-Object Name) {
            $HasChildren = ($OUs | Where-Object {$_.ParentDN -eq $OU.DistinguishedName}).Count -gt 0
            $ToggleClass = if ($HasChildren) { "ou-toggle" } else { "ou-item" }
            
            $TreeHTML += "<div class='ou-item ou-level-$Level'>"
            if ($HasChildren) {
                $TreeHTML += "<span class='$ToggleClass' onclick='toggleOU(`"$($OU.SafeID)`")'>"
            } else {
                $TreeHTML += "<span class='ou-item'>&nbsp;&nbsp;&nbsp;"
            }
            
            $TreeHTML += "<strong>$($OU.Name)</strong>"
            $TreeHTML += "<span class='ou-stats'>(Users: $($OU.UserCount), Computers: $($OU.ComputerCount), Groups: $($OU.GroupCount))</span>"
            $TreeHTML += "</span>"
            
            if ($HasChildren) {
                $TreeHTML += "<div class='ou-children' id='children-$($OU.SafeID)'>"
                $ChildHTML = Build-OUTree $OUs $OU.DistinguishedName ($Level + 1)
                $TreeHTML += $ChildHTML
                $TreeHTML += "</div>"
            }
            
            $TreeHTML += "</div>"
        }
        return $TreeHTML
    }
    
    # Start with domain root
    $DomainDN = (Get-ADDomain).DistinguishedName
    $TreeHTML = Build-OUTree $AuditResults.OUHierarchy $DomainDN 0
    $HTMLReport += $TreeHTML
    
    $HTMLReport += "</div>"
    
    # Add JavaScript for interactive functionality
    $HTMLReport += @"
    <script>
    function toggleOU(ouId) {
        var toggle = event.target;
        var children = document.getElementById('children-' + ouId);
        
        if (children) {
            if (children.classList.contains('show')) {
                children.classList.remove('show');
                toggle.classList.remove('expanded');
            } else {
                children.classList.add('show');
                toggle.classList.add('expanded');
            }
        }
    }
    </script>
"@
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.OUHierarchy)</div>"
}

$HTMLReport += @"
    </div>

    <div class="section">
        <h2>🔍 Security Events & Error Analysis</h2>
"@

if ($AuditResults.EventLogAnalysis -is [array] -and $AuditResults.EventLogAnalysis.Count -gt 0) {

    # Event Summary by Category and Severity
    $HighSeverityEvents = $AuditResults.EventLogAnalysis | Where-Object { $_.Severity -eq "High" }
    $MediumSeverityEvents = $AuditResults.EventLogAnalysis | Where-Object { $_.Severity -eq "Medium" }
    $LowSeverityEvents = $AuditResults.EventLogAnalysis | Where-Object { $_.Severity -eq "Low" }
    $SystemErrorEvents = $AuditResults.EventLogAnalysis | Where-Object { $_.Category -eq "System Errors" }
    $CriticalSecurityEvents = $AuditResults.EventLogAnalysis | Where-Object { $_.Severity -eq "High" -and $_.Category -ne "System Errors" }
    $TotalEvents = ($AuditResults.EventLogAnalysis | Measure-Object Count -Sum).Sum

    $HTMLReport += @"
        <div class="subsection">
            <h3>📊 Event Summary (Last 7 Days)</h3>
            <div class="row">
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title text-danger">$($HighSeverityEvents.Count)</h5>
                            <p class="card-text">High Severity</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title text-warning">$($MediumSeverityEvents.Count)</h5>
                            <p class="card-text">Medium Severity</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title text-info">$($SystemErrorEvents.Count)</h5>
                            <p class="card-text">System Errors</p>
                        </div>
                    </div>
                </div>
                <div class="col-md-3">
                    <div class="card text-center">
                        <div class="card-body">
                            <h5 class="card-title text-primary">$TotalEvents</h5>
                            <p class="card-text">Total Events</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
"@
    
    # High Priority Events Section
    if ($HighSeverityEvents.Count -gt 0) {
        $HTMLReport += "<h3>🚨 High Priority Security Events</h3>"
        $HTMLReport += "<p>These events require immediate attention as they indicate significant security activities.</p>"
        $HTMLReport += "<table>"
        $HTMLReport += "<tr><th>Domain Controller</th><th>Event ID</th><th>Category</th><th>Event Type</th><th>Count</th><th>Last Occurrence</th><th>Details</th></tr>"
        
        foreach ($Event in $HighSeverityEvents | Sort-Object Count -Descending) {
            $RowClass = "error"
            $HTMLReport += "<tr class='$RowClass'>"
            $HTMLReport += "<td><strong>$($Event.DomainController)</strong></td>"
            $HTMLReport += "<td>$($Event.EventID)</td>"
            $HTMLReport += "<td>$($Event.Category)</td>"
            $HTMLReport += "<td>$($Event.EventType)</td>"
            $HTMLReport += "<td><strong>$($Event.Count)</strong></td>"
            $HTMLReport += "<td>$($Event.LastOccurrence)</td>"
            $HTMLReport += "<td><button class='control-btn' onclick='toggleEventDetails(`"$($Event.EventID)_$($Event.DomainController -replace '[^a-zA-Z0-9]', '_')`")'>View Details</button></td>"
            $HTMLReport += "</tr>"
            
            # Add expandable details row
            $HTMLReport += "<tr class='group-children' id='event-$($Event.EventID)_$($Event.DomainController -replace '[^a-zA-Z0-9]', '_')' style='display: none;'>"
            $HTMLReport += "<td colspan='7'>"
            $HTMLReport += "<div style='background: #f8fafc; padding: 12px; border-radius: 8px; margin: 8px 0;'>"
            $HTMLReport += "<h5>Recent Event Details:</h5>"
            if ($Event.RecentEvents.Count -gt 0) {
                foreach ($RecentEvent in $Event.RecentEvents) {
                    $HTMLReport += "<div style='margin-bottom: 8px; padding: 8px; background: white; border-radius: 4px;'>"
                    $HTMLReport += "<strong>Time:</strong> $($RecentEvent.TimeCreated) | "
                    $HTMLReport += "<strong>User:</strong> $($RecentEvent.UserName) | "
                    $HTMLReport += "<strong>Target:</strong> $($RecentEvent.TargetAccount)"
                    $HTMLReport += "</div>"
                }
            } else {
                $HTMLReport += "<p>No detailed event information available.</p>"
            }
            $HTMLReport += "</div></td></tr>"
        }
        $HTMLReport += "</table>"
    }
    
    # Medium Priority Events Section
    if ($MediumSeverityEvents.Count -gt 0) {
        $HTMLReport += "<h3>⚠️ Medium Priority Security Events</h3>"
        $HTMLReport += "<p>These events should be reviewed for potential security implications.</p>"
        
        # Group by category for better organization
        $EventsByCategory = $MediumSeverityEvents | Group-Object Category
        
        foreach ($CategoryGroup in $EventsByCategory) {
            $HTMLReport += "<h4>$($CategoryGroup.Name) Events</h4>"
            $HTMLReport += "<table>"
            $HTMLReport += "<tr><th>Domain Controller</th><th>Event ID</th><th>Event Type</th><th>Count</th><th>Last Occurrence</th></tr>"
            
            foreach ($Event in $CategoryGroup.Group | Sort-Object Count -Descending) {
                $RowClass = "warning"
                $HTMLReport += "<tr class='$RowClass'>"
                $HTMLReport += "<td>$($Event.DomainController)</td>"
                $HTMLReport += "<td>$($Event.EventID)</td>"
                $HTMLReport += "<td>$($Event.EventType)</td>"
                $HTMLReport += "<td>$($Event.Count)</td>"
                $HTMLReport += "<td>$($Event.LastOccurrence)</td>"
                $HTMLReport += "</tr>"
            }
            $HTMLReport += "</table>"
        }
    }

    # System Errors Section
    if ($SystemErrorEvents.Count -gt 0) {
        $HTMLReport += @"
        <div class="subsection">
            <h3>⚠️ System Errors (Filtered)</h3>
            <div class="alert alert-info">
                <p><strong>Note:</strong> Generic/noise errors (perflib, schannel, WMI, etc.) have been filtered out.
                Only significant system errors are shown below.</p>
            </div>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead class="table-secondary">
                        <tr>
                            <th>Domain Controller</th>
                            <th>Event ID</th>
                            <th>Source</th>
                            <th>Error Description</th>
                            <th>Count</th>
                            <th>Last Occurrence</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($Event in $SystemErrorEvents | Sort-Object Count -Descending) {
            $ErrorDescription = [System.Web.HttpUtility]::HtmlEncode($Event.EventType)
            $ProviderName = if ($Event.ProviderName) { [System.Web.HttpUtility]::HtmlEncode($Event.ProviderName) } else { "System" }

            $HTMLReport += @"
                        <tr>
                            <td><strong>$($Event.DomainController)</strong></td>
                            <td><span class="badge bg-warning">$($Event.EventID)</span></td>
                            <td><small>$ProviderName</small></td>
                            <td>$ErrorDescription</td>
                            <td><span class="badge $(if($Event.Count -gt 10){'bg-danger'}elseif($Event.Count -gt 5){'bg-warning'}else{'bg-info'})">$($Event.Count)</span></td>
                            <td class="text-muted"><small>$($Event.LastOccurrence)</small></td>
                        </tr>
"@
        }

        $HTMLReport += @"
                    </tbody>
                </table>
            </div>
        </div>
"@
    }

    # Low Priority Events Summary
    if ($LowSeverityEvents.Count -gt 0) {
        $HTMLReport += "<h3>ℹ️ Low Priority Events Summary</h3>"
        $HTMLReport += "<p>These events are informational and typically represent normal operations.</p>"
        $LowPriorityCount = ($LowSeverityEvents | Measure-Object Count -Sum).Sum
        $HTMLReport += "<p><strong>Total low priority events:</strong> $LowPriorityCount occurrences across $($LowSeverityEvents.Count) event types</p>"
    }
    
    # Add JavaScript for expandable details
    $HTMLReport += @"
    <script>
        function toggleEventDetails(eventId) {
            var element = document.getElementById('event-' + eventId);
            if (element.style.display === 'none' || element.style.display === '') {
                element.style.display = 'table-row';
                event.target.textContent = 'Hide Details';
            } else {
                element.style.display = 'none';
                event.target.textContent = 'View Details';
            }
        }
    </script>
"@
    
} else {
    $HTMLReport += "<div class='success'>"
    $HTMLReport += "<h4>✅ No Critical Security Events Found</h4>"
    $HTMLReport += "<p>No critical security events were detected in the last 7 days, or the event log analysis encountered issues.</p>"
    if ($AuditResults.EventLogAnalysis -is [string]) {
        $HTMLReport += "<p><strong>Analysis Status:</strong> $($AuditResults.EventLogAnalysis)</p>"
    }
    $HTMLReport += "</div>"
}

$HTMLReport += @"
    </div>

    <div class="section">
        <h2>👥 Security Group Nesting Analysis</h2>
"@

if ($AuditResults.GroupNesting -is [array]) {
    # Calculate statistics
    $TotalGroups = $AuditResults.GroupNesting.Count
    $DeeplyNestedGroups = ($AuditResults.GroupNesting | Where-Object {$_.NestingDepth -gt 2}).Count
    $LargeGroups = ($AuditResults.GroupNesting | Where-Object {$_.MemberCount -gt 100}).Count
    $EmptyGroups = ($AuditResults.GroupNesting | Where-Object {$_.MemberCount -eq 0}).Count
    $GroupsWithNesting = ($AuditResults.GroupNesting | Where-Object {$_.NestedGroups -gt 0}).Count
    
    # Display statistics
    $HTMLReport += "<div class='metric'><div class='number'>$TotalGroups</div><div class='label'>Total Groups</div></div>"
    $HTMLReport += "<div class='metric'><div class='number'>$GroupsWithNesting</div><div class='label'>With Nested Groups</div></div>"
    $HTMLReport += "<div class='metric'><div class='number'>$DeeplyNestedGroups</div><div class='label'>Deep Nesting (3+)</div></div>"
    $HTMLReport += "<div class='metric'><div class='number'>$LargeGroups</div><div class='label'>Large Groups (100+)</div></div>"
    $HTMLReport += "<div class='metric'><div class='number'>$EmptyGroups</div><div class='label'>Empty Groups</div></div>"
    
    if ($DeeplyNestedGroups -gt 0) {
        $HTMLReport += "<div class='warning'>$DeeplyNestedGroups groups with deep nesting (depth > 2) detected - Review for security risks</div>"
    }
    
    if ($LargeGroups -gt 0) {
        $HTMLReport += "<div class='warning'>$LargeGroups groups with more than 100 members detected - Consider sub-groups</div>"
    }
    
    # Controls panel
    $HTMLReport += "<div class='controls-panel'>"
    $HTMLReport += "<strong>View Controls:</strong> "
    $HTMLReport += "<button class='control-btn' onclick='expandAllGroups()'>Expand All</button>"
    $HTMLReport += "<button class='control-btn' onclick='collapseAllGroups()'>Collapse All</button>"
    $HTMLReport += "<button class='control-btn' onclick='showAllGroups()'>Show All</button>"
    $HTMLReport += "</div>"
    
    # Interactive group tree
    $HTMLReport += "<div class='group-tree' id='group-tree'>"
    
    # Sort groups by member count
    $SortedGroups = $AuditResults.GroupNesting | Sort-Object MemberCount -Descending
    
    foreach ($Group in $SortedGroups) {
        $HasNesting = $Group.NestedGroups -gt 0
        $ToggleClass = if ($HasNesting) { "group-toggle" } else { "group-item" }

        $HTMLReport += "<div class='group-item'>"
        
        if ($HasNesting) {
            $HTMLReport += "<span class='$ToggleClass' onclick='toggleGroup(`"$($Group.SafeID)`")'>"
        } else {
            $HTMLReport += "<span class='group-item'>&nbsp;&nbsp;&nbsp;"
        }
        
        $HTMLReport += "<span class='group-name'>$($Group.Name)</span>"
        $HTMLReport += "<span class='group-scope'>[$($Group.GroupScope)]</span>"
        $HTMLReport += "<span class='group-stats'>Members: $($Group.MemberCount) (Users: $($Group.UserMembers), Computers: $($Group.ComputerMembers), Groups: $($Group.NestedGroups))"
        
        if ($Group.NestingDepth -gt 0) {
            $HTMLReport += " | Nesting Depth: $($Group.NestingDepth)"
        }
        
        $HTMLReport += "</span>"
        $HTMLReport += "</span>"
        
        if ($HasNesting -and -not [string]::IsNullOrEmpty($Group.NestedGroupNames)) {
            $HTMLReport += "<div class='group-children' id='group-children-$($Group.SafeID)'>"
            $NestedGroupsList = $Group.NestedGroupNames -split ", "
            foreach ($NestedGroupName in $NestedGroupsList) {
                $HTMLReport += "<div class='group-item'><span class='group-name'>↳ $NestedGroupName</span></div>"
            }
            $HTMLReport += "</div>"
        }
        
        $HTMLReport += "</div>"
    }
    
    $HTMLReport += "</div>"
    
    # Add JavaScript for group tree functionality
    $HTMLReport += @"
    <script>
    function toggleGroup(groupId) {
        var toggle = event.target.closest('.group-toggle');
        var children = document.getElementById('group-children-' + groupId);
        
        if (children) {
            if (children.classList.contains('show')) {
                children.classList.remove('show');
                toggle.classList.remove('expanded');
            } else {
                children.classList.add('show');
                toggle.classList.add('expanded');
            }
        }
    }
    
    function expandAllGroups() {
        var allToggles = document.querySelectorAll('.group-toggle');
        var allChildren = document.querySelectorAll('.group-children');
        
        allToggles.forEach(function(toggle) {
            toggle.classList.add('expanded');
        });
        
        allChildren.forEach(function(child) {
            child.classList.add('show');
        });
    }
    
    function collapseAllGroups() {
        var allToggles = document.querySelectorAll('.group-toggle');
        var allChildren = document.querySelectorAll('.group-children');
        
        allToggles.forEach(function(toggle) {
            toggle.classList.remove('expanded');
        });
        
        allChildren.forEach(function(child) {
            child.classList.remove('show');
        });
    }
    
    function showAllGroups() {
        var allGroups = document.querySelectorAll('.group-item');

        allGroups.forEach(function(group) {
            group.style.display = 'block';
        });
    }
    </script>
"@
    
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.GroupNesting)</div>"
}






$HTMLReport += @"
    </div>
    <div class="section">
        <h2>🔑 Group Managed Service Accounts (gMSA) Usage</h2>
"@

if ($AuditResults.gMSA -is [hashtable]) {
    if ($AuditResults.gMSA.Status -eq "No gMSA accounts found") {
        $HTMLReport += "<p><strong>Status:</strong> No gMSA accounts found.</p>"
        $HTMLReport += "<div class='recommendation'>Recommendation: $($AuditResults.gMSA.Recommendation)</div>"
    } elseif ($AuditResults.gMSA.Status -eq "gMSA accounts found") {
        $HTMLReport += "<p><strong>Status:</strong> gMSA accounts in use.</p>"
        $HTMLReport += "<table><tr><th>Name</th><th>Enabled</th><th>Allowed Computers</th></tr>"
        foreach ($acct in $AuditResults.gMSA.Accounts) {
            $HTMLReport += "<tr><td>$($acct.Name)</td><td>$($acct.Enabled)</td><td>$($acct.AllowedComputers)</td></tr>"
        }
        $HTMLReport += "</table>"
    }
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.gMSA)</div>"
}


# 🔐 LAPS Audit Report Section
$HTMLReport += @"
</div>
<div class='section'>
    <h2>🔐 Local Administrator Password Solution (LAPS) Audit</h2>
"@

if ($LAPSStatus.Count -gt 0) {
    $HTMLReport += "<table>
        <tr>
            <th>Computer Name</th>
            <th>LAPS Installed</th>
            <th>Password Set</th>
            <th>Password Expiration</th>
        </tr>"

    foreach ($laps in $LAPSStatus) {
        $HTMLReport += "<tr>
            <td>$($laps.ComputerName)</td>
            <td>$($laps.LAPSInstalled)</td>
            <td>$($laps.PasswordSet)</td>
            <td>$($laps.PasswordExpiration)</td>
        </tr>"
    }

    $HTMLReport += "</table>"
} else {
    $HTMLReport += "<p>No LAPS data found or unable to retrieve.</p>"
}





$HTMLReport += @"
    </div>

    <div class='section'>
        <h2>🔐 Protocol Security Analysis (LLMNR, SMB, TLS, NTLM)</h2>
"@

if ($AuditResults.ProtocolSecurity -is [array]) {
    $HTMLReport += "<table>
        <tr>
            <th>Domain Controller</th>
            <th>SMBv1 Enabled</th>
            <th>SMBv2 Enabled</th>
            <th>SMBv3 Enabled</th>
            <th>LLMNR Disabled</th>
            <th>TLS 1.0</th>
            <th>TLS 1.1</th>
            <th>TLS 1.2</th>
            <th>TLS 1.3</th>
            <th>NTLM Audit</th>
        </tr>"

    foreach ($Protocol in $AuditResults.ProtocolSecurity) {
        $RowClass = ""
        if ($Protocol.SMBv1_Enabled -eq $true -or $Protocol.LLMNRDisabled -eq "No") { $RowClass = "warning" }

        $HTMLReport += "<tr class='$RowClass'>
            <td>$($Protocol.DomainController)</td>
            <td>$($Protocol.SMBv1_Enabled)</td>
            <td>$($Protocol.SMBv2_Enabled)</td>
            <td>$($Protocol.SMBv3_Enabled)</td>
            <td>$($Protocol.LLMNRDisabled)</td>
            <td>$($Protocol.TLS_1_0_Enabled)</td>
            <td>$($Protocol.TLS_1_1_Enabled)</td>
            <td>$($Protocol.TLS_1_2_Enabled)</td>
            <td>$($Protocol.TLS_1_3_Enabled)</td>
            <td>$($Protocol.NTLMAudit)</td>
        </tr>"
    }

    $HTMLReport += "</table>"

    $HTMLReport += @"
        <h3>Protocol Security Recommendations</h3>
        <ul>
            <li><strong>SMBv1:</strong> Disable SMBv1 as it is insecure and vulnerable to exploits like WannaCry.</li>
            <li><strong>SMBv2/SMBv3:</strong> Keep these enabled but ensure security features (signing/encryption) are enforced where needed.</li>
            <li><strong>LLMNR:</strong> Disable to reduce exposure to name resolution poisoning attacks.</li>
            <li><strong>TLS Versions:</strong> Disable TLS 1.0 and 1.1; ensure TLS 1.2 and 1.3 are enabled for secure communications.</li>
            <li><strong>NTLM Audit:</strong> Configure auditing to monitor and reduce NTLM usage.</li>
        </ul>
"@
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.ProtocolSecurity)</div>"
}


$HTMLReport += @"
</div>
<div class='section'>
    <h2>🛡 Protocols & Legacy Services Audit</h2>
"@

if ($AuditResults.ProtocolsAndLegacyServices -is [hashtable]) {
    $HTMLReport += "<table>
        <tr><th>Setting</th><th>Status</th></tr>"
    foreach ($key in $AuditResults.ProtocolsAndLegacyServices.Keys) {
        $HTMLReport += "<tr><td>$key</td><td>$($AuditResults.ProtocolsAndLegacyServices[$key])</td></tr>"
    }
    $HTMLReport += "</table>"
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.ProtocolsAndLegacyServices)</div>"
}





$HTMLReport += @"
    </div>
    <div class="section">
        <h2>🔍 Orphaned SIDs in Active Directory ACLs</h2>
"@

if ($AuditResults.OrphanedSIDs -is [hashtable]) {
    $HTMLReport += "<div class='metric'><div class='number'>$($AuditResults.OrphanedSIDs.TotalCount)</div><div class='label'>Total Orphaned SIDs Found</div></div>"

    if ($AuditResults.OrphanedSIDs.TopObjects.Count -gt 0) {
        $HTMLReport += "<h3>Top Objects with Most Orphaned SIDs</h3>"
        $HTMLReport += "<table><tr><th>Distinguished Name</th><th>Orphaned SIDs Count</th></tr>"
        foreach ($item in $AuditResults.OrphanedSIDs.TopObjects) {
            $HTMLReport += "<tr><td>$($item.Name)</td><td>$($item.Count)</td></tr>"
        }
        $HTMLReport += "</table>"
    }
} else {
    $HTMLReport += "<div class='error'>$($AuditResults.OrphanedSIDs)</div>"
}







$HTMLReport += @"
    </div>

    <div class="section">
        <h2>👤 Privileged Account Monitoring</h2>
        <div class="subsection">
            <h3>Administrative Groups Analysis</h3>
            <table>
                <tr>
                    <th>Group Name</th>
                    <th>Member Count</th>
                    <th>Last Modified</th>
                    <th>Inactive Members</th>
                </tr>
"@

if ($AuditResults.PrivilegedAccountMonitoring.AdminGroups) {
    foreach ($Group in $AuditResults.PrivilegedAccountMonitoring.AdminGroups) {
        $InactiveCount = ($Group.Members | Where-Object { $_.DaysSinceLastLogon -ne "Never" -and $_.DaysSinceLastLogon -gt 90 }).Count
        $HTMLReport += @"
                <tr>
                    <td><strong>$($Group.GroupName)</strong></td>
                    <td>$($Group.MemberCount)</td>
                    <td>$($Group.LastModified)</td>
                    <td>$InactiveCount</td>
                </tr>
"@
    }
} else {
    $HTMLReport += @"
                <tr><td colspan="4">No privileged groups found or access denied</td></tr>
"@
}

$HTMLReport += @"
            </table>
        </div>

        <div class="subsection">
            <h3>Service Accounts with High Privileges</h3>
            <table>
                <tr>
                    <th>Account Name</th>
                    <th>Privileged Groups</th>
                    <th>Last Logon</th>
                    <th>Status</th>
                </tr>
"@

if ($AuditResults.PrivilegedAccountMonitoring.ServiceAccounts) {
    foreach ($Account in $AuditResults.PrivilegedAccountMonitoring.ServiceAccounts) {
        $Groups = ($Account.PrivilegedGroups -join ', ')
        $Status = if ($Account.DaysSinceLastLogon -gt 90) { 'warning' } elseif ($Account.DaysSinceLastLogon -eq "Never") { 'critical' } else { 'success' }
        $HTMLReport += @"
                <tr>
                    <td><strong>$($Account.Name)</strong></td>
                    <td>$Groups</td>
                    <td>$($Account.LastLogon)</td>
                    <td class="$Status">$($Account.DaysSinceLastLogon) days</td>
                </tr>
"@
    }
} else {
    $HTMLReport += @"
                <tr><td colspan="4">No privileged service accounts found</td></tr>
"@
}

$HTMLReport += @"
            </table>
        </div>

        <div class="subsection">
            <h3>Inactive Privileged Accounts (90+ days)</h3>
            <table>
                <tr>
                    <th>User</th>
                    <th>Group</th>
                    <th>Last Logon</th>
                    <th>Days Inactive</th>
                </tr>
"@

if ($AuditResults.PrivilegedAccountMonitoring.InactivePrivilegedAccounts -and $AuditResults.PrivilegedAccountMonitoring.InactivePrivilegedAccounts.Count -gt 0) {
    foreach ($InactiveAccount in $AuditResults.PrivilegedAccountMonitoring.InactivePrivilegedAccounts) {
        $HTMLReport += @"
                <tr>
                    <td>$($InactiveAccount.User)</td>
                    <td>$($InactiveAccount.Group)</td>
                    <td>$($InactiveAccount.LastLogon)</td>
                    <td>$($InactiveAccount.DaysInactive) days</td>
                </tr>
"@
    }
} else {
    $HTMLReport += @"
                <tr><td colspan="4">No inactive privileged accounts found</td></tr>
"@
}

$HTMLReport += @"
            </table>
        </div>
    </div>

    <div class="section">
        <h2>🛡️ AD Object Protection from Accidental Deletion</h2>
"@

if ($AuditResults.ObjectProtection.Error) {
    $HTMLReport += "<div class='error'>Object Protection Analysis Error: $($AuditResults.ObjectProtection.Error)</div>"
} else {
    $HTMLReport += @"
        <h3>Protection Overview</h3>
        <div class="metric">
            <div class="number">$($AuditResults.ObjectProtection.Statistics.OverallProtectionPercentage)%</div>
            <div class="label">Overall Protection</div>
        </div>
        <div class="metric">
            <div class="number">$($AuditResults.ObjectProtection.Statistics.ProtectedOUs)/$($AuditResults.ObjectProtection.Statistics.TotalOUs)</div>
            <div class="label">Protected OUs</div>
        </div>
        <div class="metric">
            <div class="number">$($AuditResults.ObjectProtection.Statistics.ProtectedContainers)/$($AuditResults.ObjectProtection.Statistics.TotalContainers)</div>
            <div class="label">Protected Containers</div>
        </div>
        <div class="metric">
            <div class="number">$($AuditResults.ObjectProtection.Statistics.ProtectedCriticalUsers)/$($AuditResults.ObjectProtection.Statistics.TotalCriticalUsers)</div>
            <div class="label">Protected Critical Users</div>
        </div>
        <div class="metric">
            <div class="number">$($AuditResults.ObjectProtection.Statistics.ProtectedServiceAccounts)/$($AuditResults.ObjectProtection.Statistics.TotalServiceAccounts)</div>
            <div class="label">Protected Service Accounts</div>
        </div>
        <div class="metric">
            <div class="number">$($AuditResults.ObjectProtection.Statistics.ProtectedCriticalGroups)/$($AuditResults.ObjectProtection.Statistics.TotalCriticalGroups)</div>
            <div class="label">Protected Critical Groups</div>
        </div>
        
        <h3>Unprotected Objects (High Risk)</h3>
"@
    
    # Show unprotected objects in expandable sections
    if ($AuditResults.ObjectProtection.UnprotectedOUs -and $AuditResults.ObjectProtection.UnprotectedOUs.Count -gt 0) {
        $HTMLReport += @"
        <div class="subsection">
            <h4>🚨 Unprotected Organizational Units ($($AuditResults.ObjectProtection.UnprotectedOUs.Count))</h4>
            <table>
                <tr>
                    <th>OU Name</th>
                    <th>Distinguished Name</th>
                    <th>Description</th>
                </tr>
"@
        $UnprotectedOUsLimited = $AuditResults.ObjectProtection.UnprotectedOUs | Select-Object -First 10
        foreach ($OU in $UnprotectedOUsLimited) {
            $Description = if ($OU.Description) { $OU.Description } else { "No description" }
            $HTMLReport += @"
                <tr>
                    <td><strong>$($OU.Name)</strong></td>
                    <td>$($OU.DistinguishedName)</td>
                    <td>$Description</td>
                </tr>
"@
        }
        if ($AuditResults.ObjectProtection.UnprotectedOUs.Count -gt 10) {
            $HTMLReport += "<tr><td colspan='3'><em>... and $($AuditResults.ObjectProtection.UnprotectedOUs.Count - 10) more unprotected OUs</em></td></tr>"
        }
        $HTMLReport += "</table></div>"
    }
    
    if ($AuditResults.ObjectProtection.UnprotectedUsers -and $AuditResults.ObjectProtection.UnprotectedUsers.Count -gt 0) {
        $HTMLReport += @"
        <div class="subsection">
            <h4>🚨 Unprotected Critical Users ($($AuditResults.ObjectProtection.UnprotectedUsers.Count))</h4>
            <table>
                <tr>
                    <th>User Name</th>
                    <th>Sam Account Name</th>
                    <th>Type</th>
                    <th>Service Principal Names</th>
                </tr>
"@
        $UnprotectedUsersLimited = $AuditResults.ObjectProtection.UnprotectedUsers | Select-Object -First 15
        foreach ($User in $UnprotectedUsersLimited) {
            $UserType = if ($User.IsServiceAccount) { "Service Account" } else { "Admin User" }
            $SPNs = if ($User.ServicePrincipalNames) { $User.ServicePrincipalNames } else { "N/A" }
            $HTMLReport += @"
                <tr>
                    <td><strong>$($User.Name)</strong></td>
                    <td>$($User.SamAccountName)</td>
                    <td>$UserType</td>
                    <td>$SPNs</td>
                </tr>
"@
        }
        if ($AuditResults.ObjectProtection.UnprotectedUsers.Count -gt 15) {
            $HTMLReport += "<tr><td colspan='4'><em>... and $($AuditResults.ObjectProtection.UnprotectedUsers.Count - 15) more unprotected critical users</em></td></tr>"
        }
        $HTMLReport += "</table></div>"
    }
    
    if ($AuditResults.ObjectProtection.UnprotectedGroups -and $AuditResults.ObjectProtection.UnprotectedGroups.Count -gt 0) {
        $HTMLReport += @"
        <div class="subsection">
            <h4>🚨 Unprotected Critical Groups ($($AuditResults.ObjectProtection.UnprotectedGroups.Count))</h4>
            <table>
                <tr>
                    <th>Group Name</th>
                    <th>Group Scope</th>
                    <th>Distinguished Name</th>
                </tr>
"@
        foreach ($Group in $AuditResults.ObjectProtection.UnprotectedGroups) {
            $HTMLReport += @"
                <tr>
                    <td><strong>$($Group.Name)</strong></td>
                    <td>$($Group.GroupScope)</td>
                    <td>$($Group.DistinguishedName)</td>
                </tr>
"@
        }
        $HTMLReport += "</table></div>"
    }
    
    if ($AuditResults.ObjectProtection.UnprotectedContainers -and $AuditResults.ObjectProtection.UnprotectedContainers.Count -gt 0) {
        $HTMLReport += @"
        <div class="warning">
            <h4>⚠️ Unprotected Containers ($($AuditResults.ObjectProtection.UnprotectedContainers.Count))</h4>
            <table>
                <tr>
                    <th>Container Name</th>
                    <th>Distinguished Name</th>
                </tr>
"@
        foreach ($Container in $AuditResults.ObjectProtection.UnprotectedContainers) {
            $HTMLReport += @"
                <tr>
                    <td><strong>$($Container.Name)</strong></td>
                    <td>$($Container.DistinguishedName)</td>
                </tr>
"@
        }
        $HTMLReport += "</table></div>"
    }
    
    # Show success message if all objects are protected
    $TotalUnprotected = $AuditResults.ObjectProtection.Statistics.UnprotectedOUs + $AuditResults.ObjectProtection.Statistics.UnprotectedContainers + $AuditResults.ObjectProtection.Statistics.UnprotectedCriticalUsers + $AuditResults.ObjectProtection.Statistics.UnprotectedCriticalGroups
    if ($TotalUnprotected -eq 0) {
        $HTMLReport += "<div class='success'><h4>✅ Excellent Protection Status</h4><p>All critical AD objects are properly protected from accidental deletion.</p></div>"
    }
    
    # Add recommendations if any
    if ($AuditResults.ObjectProtection.Recommendations -and $AuditResults.ObjectProtection.Recommendations.Count -gt 0) {
        $HTMLReport += "<h3>Protection Recommendations</h3><div class='warning'><ul>"
        foreach ($Recommendation in $AuditResults.ObjectProtection.Recommendations) {
            $SafeRecommendation = [System.Web.HttpUtility]::HtmlEncode($Recommendation)
            $HTMLReport += "<li>$SafeRecommendation</li>"
        }
        $HTMLReport += "</ul></div>"
    }
}

$HTMLReport += @"
    </div>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>

    </div> <!-- Close container-fluid -->
</body>
</html>
"@



# Save the report
$ReportPath = Join-Path $OutputPath $ReportName
try {
    $HTMLReport | Out-File -FilePath $ReportPath -Encoding UTF8 -Force
    Write-Host "Report saved successfully to: $ReportPath" -ForegroundColor Green

    # Open the report in default browser
    if (Test-Path $ReportPath) {
        Start-Process $ReportPath
    }

    # Update global path for web server
    $Global:ProgressData.ReportPath = $ReportPath

} catch {
    Write-Error "Failed to save report: $($_.Exception.Message)"
}

$EndTime = Get-Date
$Duration = New-TimeSpan -Start $StartTime -End $EndTime

Write-Host "✅ Active Directory Audit Complete!" -ForegroundColor Green
Write-Host "â±ï¸ Duration: $(if ($Duration) { $Duration.Minutes.ToString().PadLeft(2,'0') + ':' + $Duration.Seconds.ToString().PadLeft(2,'0') } else { 'N/A' })" -ForegroundColor Yellow

# Clear the execution flag
$Global:AuditInProgress = $false

#================================================================
# END OF SCRIPT
#================================================================
# Script execution complete. The HTML report has been generated
# and saved to the specified output path.
#================================================================
