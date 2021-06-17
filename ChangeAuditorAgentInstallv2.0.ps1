<# .NOTES

        Created By      : Nishant Tyagi
        Created For     : *****
        Created On      : 08/06/2018
        Purpose         : Deploy Change Auditor Agents to Servers Remotely Using Powershell
        Version Control : 2.0

    .SYPNOSIS
            To Automate the Installation of Change Auditor Agents to Remote Windows servers using Change Auditor Powershell Module.
            
    .DESCRIPTION
            1. The Script will do the following:
            
                a. Agreement and Instructions Prompt
                b. Prompt the User to Select the Input Text File through GUI Windows Explorer.
                c. Prompt the user to select the Change Auditor PowerShell Module file.
                d. Prompt the user to select the Change Auditor Log file.
                e. Prompt the user to provide Email Recipients in To & CC fields.

            2. Send an Email to Email recipients indicating Agent Installation Initiation with batch Information.
            3. Run Pre-Requisite Tests
                a. Check If the Server is available i.e. a remote connection is avaiable, if yes - continue
                b. Start a Powershell Remote Session to the server and gather .NET Framework version
                c. Validate if the version is above .NET 4.5.2, if yes - continue
            4. Attempt Installation, Use Do while Loop with Function Deployment-Result to track Installation Status
            5. Export a csv list in script run directory for the following:
                a. List of Unavailable Servers i.e. Unreachable_Servers.csv
                b. List of servers where .Net requirement is not met i.e. Net_Not_Met.csv
                c. List if Servers where Installation Failed i.e. Install_Failure.csv
                d. List of successful Installs with Agent Status i.e. Deployment_Result.csv

    .NOTE
         1. Please Ensure to at-least use Domain Admin credentials for the Domain servers belong to.
            If the List of servers is a mix of diffrent domains inside the same forest, please provide
            Enterprise Admin Credentials at the prompt.
         2. The Script will fail if the agents are to be deployed in different forest servers, Please prepare 
            seperate files on the basis of domains,Forest(In Case Enterprise Admin credential are avaialable.
         3. The Input list should contain the the server names in FQDN Format. The Install will not work if server names are not present in FQDN.

    .Version Control
         1. 08/JUN/2018 : Version 1.0 - Initial Code for Agent Deployment Testing
         2. 21/AUG/2018 : Version 1.1 - .NET Framework Version Section updated
         3. 10/NOV/2018 : Version 1.2 - Installation Logic Adjusted to accomodate Bug Fix's.
         4. 22/APR/2019 : Version 2.0 - Code Rebuilt for v7.0.2 Agent Install + Template Deployment Logic
#>

#Step 1: Provide Code Usage Instructions in the Shell. The Sleep setting has been added to provide enough time to read the Instructions for Admins.

        $Instructions = "
        Running ChangeAuditorAgentInstallv2.0 ...
        Adding PresentationFramework Assembly to Current Powershell Session...
        Loading Usage Instructions..."
        Write-Host $Instructions -ForegroundColor White -BackgroundColor Black
        Add-Type -AssemblyName PresentationFramework
        sleep -Seconds 5
                        
        $Instructions1 = "
                    
1. You must be a part of the following AD groups in the respective AD 
    domain for successful deployment of Agents:
    - DomainAdmins
    - ChangeAuditorAdministrators

2. You will be first asked to choose the Input File. The input file 
    shall meet the following requirements:
   - Must be provided in .txt format.
   - Must NOT contain any headers\Headings.
   - Must contain the server names in FQDN format 
      e.g. server01.domain.com

3. You will then be prompted to choose the following files:

   PROMPT 1: Change Auditor Powershell Module File:   
      - File Name: ChangeAuditor.Powershell.dll
      - File Details: This File is needed to Load Change Auditor PS
                            Module to the current powershell Session. 
      - File Path:InstallDirectory%\ProgramFiles\Quest\ChangeAuditor
                        \Client
     
   PROMPT 2: Change Auditor Log File:                   
      - File Name: ChangeAuditor.Servicelog.nptlog
      - File Details: This file will be used to report the Agent 
                             Install result into export files and emails.
      - File Path: %InstallDirectory%\ProgramFiles\Quest\ChangeAuditor
                        \Service\Logs
    
4. The Email Reciepients must be provided when prompted.
    These recipients will recieve a Initial and Final Report of the
    Agent Deployment Task.
                        
    Please Choose:
                      
     Yes - If all pre-requisites are met.
     No  - To Abort Script.
"

        $PrereqInput = [System.Windows.MessageBox]::Show($Instructions1,'SCRIPT and AGENT Installation PRE-REQUISITES:','YesNo','Error')

        If($PrereqInput -eq "Yes"){ 
        Write-Host "Input 'Yes' Registered. Initializing Code and Loading Variables...." -ForegroundColor Yellow -BackgroundColor Black
        Sleep -Seconds 10
        }
        else{ 
        Write-Host "You Chose 'No'. Terminating Code Execution & Closing Shell. GoodBye!" -ForegroundColor Yellow -BackgroundColor Black
        exit;
        }

#STEP 2: Declare Functions.

    #1. Function to accept Input file(TXT), the Function will open a GUI File explorer window in C: Directory.
        Function Get-FileName($InitialDirectory)
                                {
        [System.Reflection.Assembly]::LoadwithPartialName("System.Windows.Forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.InitialDirectory = $InitialDirectory
        $OpenFileDialog.Filter = "TXT (*.txt) | *.txt"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.FileName
        } #End of Function Get-FileName

    #2. Function to Choose Powershell DLL File, the Function will open a GUI File explorer window in C: Directory.
        Function Get-PSDLL($InitialDirectory)
                                {
        [System.Reflection.Assembly]::LoadwithPartialName("System.Windows.Forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.InitialDirectory = $InitialDirectory
        $OpenFileDialog.Filter = "dll (*.dll) | *.dll"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.FileName
        } #End of Function Get-PSDLL

    #3. Scriptblock to Evaluate .Net Framework Version on target machines
        [scriptblock]$NetEvaluator = {
                        $NetRegKey = Get-Childitem "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue
                        $Release = $NetRegKey.GetValue("Release")
                        Switch($Release){
                            379893 {$NetFrameworkVersion = "4.5.2"}
                            393295 {$NetFrameworkVersion = "4.6"}
                            393297 {$NetFrameworkVersion = "4.6"}
                            394254 {$NetFrameworkVersion = "4.6.1"}
                            394271 {$NetFrameworkVersion = "4.6.1"}
                            394802 {$NetFrameworkVersion = "4.6.2"}
                            394806 {$NetFrameworkVersion = "4.6.2"}
                            460798 {$NetFrameworkVersion = "4.7"}
                            460805 {$NetFrameworkVersion = "4.7"}
                            461308 {$NetFrameworkVersion = "4.7.1"}
                            461310 {$NetFrameworkVersion = "4.7.1"}
                            461808 {$NetFrameworkVersion = "4.7.2"}
                            461814 {$NetFrameworkVersion = "4.7.2"}
                            528040 {$NetFrameworkVersion = "4.8"}
                            528049 {$NetFrameworkVersion = "4.8"}
                            Default {$NetFrameworkVersion = "Net Framework Requirement Not Met"}
            }
        #Create Powershell Object to Return results in Foreach section.
                        $object = [PSCustomObject]@{
                                                        Computername = $env:COMPUTERNAME
                                                        NETFrameworkVersion = $NetFrameworkVersion
                                                   }
        $object
        }

    
    #4. Function to choose Change Auditor Log file, the Function will open a GUI File Explorer Window in C: Directory.
        Function Get-CANptLog($InitialDirectory)
                                {
        [System.Reflection.Assembly]::LoadwithPartialName("System.Windows.Forms") | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.InitialDirectory = $InitialDirectory
        $OpenFileDialog.Filter = "nptlog (*.nptlog) | *.nptlog"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.FileName
        } #End of Function Get-CANptLog
        
    #5. Function to evaluate deployement results; Used in Do-while loop later in this code.
         Function Deployment-Result
        {
            $FailureMessage = "Error" #Event Snip for a Failed Attempt. Extracted from CA Logs
            $SuccessMessage = "Successfully completed deployment to $server"  #Event Snip for a Successful Attempt. Extracted from CA Logs
    
            #Import the Change Auditor Log File to Extract Installation Attempt Events for Status Reporting
            $CAEvents =  gc $nptLogFile

            #Filter CA Events to only return the events for server name in current loop
            $FilteredMachineEvents = $CAEvents | ?{$_ -match $server}
            #Run Status Lookup on $FilteredMachineEvents to evaluate deployment result
                If($FilteredMachineEvents -match $FailureMessage)
                    {
                        return "Installation Failed! Check Change Auditor Logs for More Details"
                    }
                elseIf($FilteredMachineEvents -match $SuccessMessage)
                    {
                        return "Successfully Installed the Agent"
                    }
                else{
                        return 'Install still in Progress'
                    } 

        } #End of Function Deployment-Result
       
  
#STEP 3: SMTP Email Section
        $TodayDate= Get-Date -DisplayHint Date
        $smtphost = "internalmail.r02.xlgs.local"  
        $from = "XLGS-ChangeAuditor@xlgroup.com"  
        $to = Read-Host -Prompt "Please enter Recipients in 'TO' Field:"
        $cc = Read-Host -Prompt "Please enter Recipients in 'CC' Field:"
        $subject = "Started! BULK Change Auditor Agent v7.0.2 Deployment"
        $body = "Hello,

        The Change Auditor Agent Deployment Task has been started by IAM Team at $TodayDate EST.

        Total No. of Servers in Current Batch: $TotalServers

        Regards,
        XLC ADIMS IAM Team"
        Send-MailMessage -From $from -Body $body -To $to -Cc $cc -Subject $subject -SmtpServer $smtphost

#STEP 4: Load Data from Files.

   #1. Load Input Server File
        Write-Host "Loading File Input Window.." -ForegroundColor Yellow -BackgroundColor Black
        $Inputfile =  Get-FileName "C:"
        $Servers = gc $Inputfile
        $TotalServers = $Servers.count
        Write-Host "Input File has been loaded. Total no. of servers in Agent Installation Scope: " $TotalServers -ForegroundColor Green -BackgroundColor Black

   #2. Load Powershell Module File.
        Write-Host "Loading Window to Choose Change Auditor PowerShell DLL File:" -ForegroundColor Yellow -BackgroundColor Black
        $PsdllPath =  Get-PSDLL "C:"
        Import-Module $psdllpath
        Write-Host "Successfully Loaded Change Auditor Powershell Module into this session" -ForegroundColor Green -BackgroundColor Black

   #3.Load Admin Creds & Change Auditor Connection
        Write-Host "Please Enter the Domain Admin or Enterprise Admin Credentials at the Prompt" -ForegroundColor Yellow -BackgroundColor Black
        $Credential = Get-Credential     #Domain\Enterprise Admin Credentials are required
        Write-Host "Connecting to ChangeAuditor2013 Instance. Please wait.."
        $Connection = Connect-CAClient -InstallationName CHANGEAUDITOR2013 #Connect to Local Change Auditor Coordinator Server Instance
        Write-Host "Successfully Connected to CHANGEAUDITOR2013" -ForegroundColor Green -BackgroundColor Black

   #4. Load Agent Information from Chane Auditor Database.
        Write-Host "Importing All Change Auditor Agent status from Change Auditor Database. Please wait.." -ForegroundColor Yellow -BackgroundColor Black
        $AllAgentInfo = Get-CAAgents -IncludeUninstalled -Connection $connection
        Write-Host "All Agent information has been loaded" -ForegroundColor Green -BackgroundColor Black

   #5. Import Change Auditor Logs into Variable
        $nptLogFile = Get-CANptLog "C:"
        Write-Host "Change Auditor logfile path set to :" $nptlogfile -ForegroundColor Green -BackgroundColor Black

   #6. Declare Miscellanoeus Variables
        $Unreachable = @()               #Captures List Of Unavailable Servers, Appends are done by $ConnectionStatus
        $NetOutput = @()                 #Captures the output of .Net Check Invoke command on Remote servers
        $NetNotMet = @()                 #Captures the List of Servers where .Net Framework version is less than .NET 4.5.2
        $FailedInstallation = @()        #Captures the List of Servers where Installation Attempt was failed.
        $SuccessfullyInstalled = @()     #Captures the List of Servers where CA Agent was successfully Installed by the script.


#################################################################################################
###################                       MAIN SECTION                        ###################
#################################################################################################


#STEP 6: Run Pre-requisites test on each server to determine next Action.

          ForEach($server in $Servers)
                {
                    Write-Host $server": Added to Installation Queue" -ForegroundColor Green -BackgroundColor Black
                    Write-Host $server": Initiating Pre-Requisites Checks" -ForegroundColor Yellow -BackgroundColor Black
    #Test Connnectivity to the server. 
                            
                               try
                                {
                                  Write-Host $server": Starting Remote Connectivity Check..." -ForegroundColor Yellow -BackgroundColor Black
                                  if((Test-Connection $server -Quiet) -eq 'True')
                                    {
                                        Write-Host $server": Server is Reachable !" -ForegroundColor Green -BackgroundColor Black
                                        Write-Host "Initiating PowerShell Remote Session to $server ...." -ForegroundColor Yellow -BackgroundColor Black
                                        $Session = New-PSSession $server 
                                        $NetOutput = Invoke-Command -Session $Session -ScriptBlock $NetEvaluator

                                        Write-host "Remote Session Estabilished with $server" -ForegroundColor Green -BackgroundColor Black

    #6.3: Test .NET Framework Version to ensure it is v4.5.2 or greater.

                                        Write-Host $server": Initiating .NET Framework Version Evaluation" -ForegroundColor Yellow -BackgroundColor Black
                                        
                                        If($NetOutput.NetFrameworkversion -eq "Net Framework Requirement Not Met" -or $NetOutput.NETFrameworkVersion -eq $Null)
                                            {
                                                Write-host $Server": .Net Framework Version is Below v4.5.2, Aborting Install" -ForegroundColor Red -BackgroundColor Black
                                                $NetNotMet += $server

                                                #Disconnect Connected PS Session to complete abort
                                                Write-Host $server": Releasing PS Remote Session..." -ForegroundColor Yellow -BackgroundColor Black
                                                Get-PSSession | Remove-PSSession -Confirm:$false
                                                Write-host $server": Remote Session Disconnected." -ForegroundColor Yellow -BackgroundColor Black
                                            }
    #6.4: All Installation Pre-Requisites are met, Attempt Agent Installation .

                                        else
                                            {
                                              try
                                                {
                                                    Write-Host $server ": .Net Framework is Supported! Initiating Agent Installation. Please Wait.." -ForegroundColor Cyan -BackgroundColor Black
                                                    
                                                    Install-CAAgent -Connection $connection -MachineName $server -Credential $Credential -ErrorAction Stop

    #Begin Do While Loop. The Loop Ensures Agent Instalation finished using Function Deployment-Result
                                                    do
                                                        {
    #Allow Installation to Complete in the Background. An ideal installation takes about 30-60 secs to complete.
                                                            sleep -Seconds 30
                                                            $FunctionResult = Deployment-Result #Run the Function Deployment-Result for Status Results

                                                            if( $FunctionResult -eq 'Installation Failed! Check Change Auditor Logs for More Details') {break}
                            
                                                            elseif($FunctionResult -eq 'Successfully Installed the Agent') {break}
                            
                                                            else{ sleep -Seconds 60 } #sleep again for 60 secs to allow more time for agent installation.Loop
                                                        }
                                                    while
                                                        (
                                                            ($FunctionResult -eq "Installation Failed! Check Change Auditor Logs for More Details") -or ($FunctionResult -eq "Successfully Installed the Agent") -or ($FunctionResult -eq "Install still in Progress")
                                                        )
                                                        
                                                    if( $FunctionResult -eq 'Installation Failed! Check Change Auditor Logs for More Details')
                                                        {
                                                            Write-Host $Server ": Agent Installation Attempt Complete. Result:" $FunctionResult -ForegroundColor Red -BackgroundColor Black
                                                            $FailedInstallation += $Server 
                                                        }
                                                    elseif($FunctionResult -eq 'Successfully Installed the Agent') 
                                                        {
                                                            Write-Host $Server": Agent Installation Attempt Complete. Result: Agent Was Installed Successfully" -ForegroundColor Green -BackgroundColor Black
                                                            $SuccessfullyInstalled += $server
                                                        }
                                                    else
                                                        {
                                                            #donothing 
                                                        } 
                                                } #end of Installation Attempt try block

    #Catch Installation Non-Terminating Errors.
                                              Catch
                                                {
                                                   $ErrorMessage = $_.Exception.Message
                                                   $ReError = Write-Host $server ": Installation Failed with Error:" $ErrorMessage -ForegroundColor Red -BackgroundColor Black 
                                                   $FailedInstallation += $ReError
                                                }  
                      
    #Clear ongoing Powershell Remote Session and System Variable $Error for Fresh processing in next loop
                                              Finally
                                                { 
                                                   Get-PSSession | Remove-PSSession -Confirm:$false
                                                   $Error.Clear()
                                                }
                                                 
                                            } #End of Nested Else Condition - Agent Installation Attempt Block

                                    }#End of Main If Statement - Reachable Servers Action Block 
         
                                 else
    #Server is Unavailable
                                    {
                                        Write-host $Server": Connectivity Test Failed.!! Please Check Server Availability" -ForegroundColor Red -BackgroundColor Black
                                        $Unreachable += $server
                                    }   
              }#End of Main Try Block - Prerequisites i.e. Connection, .NET, Event, Status Check
                               Catch
                                {
                                  #DoNothing
                                }
                            #End of CA Staus Check Else Condition

    Write-Host "Attempting Next Server in the List." -ForegroundColor Yellow -BackgroundColor Black `n

    } #End of Main ForEach Loop


    Write-Host "All the Actions have been completed against Current Input. Please Check the Run Directory for Output files" -ForegroundColor Green -BackgroundColor Black

    #Export TXT Files For Status Tracking
    $Unreachable | Out-file Unreachable_Servers.txt
    $FailedInstallation | Out-File Installation_Failures.txt
    $NetNotMet | Out-File Net_Framework_Not_Met.txt
    $SuccessfullyInstalled | Out-File CA_Successfully_Installed.txt

    #Calculate No. of servers in each category for Email report.
    $Success = $SuccessfullyInstalled.Count
    $Unreachable = $Unreachable.Count
    $FrameworkIssue = $NetNotMet.Count
    $ActionNeeded  = $FailedInstallation.Count

    #SMTP Details for Completion Report
    $FinalTime= Get-Date -DisplayHint Date
    $subject1 = "Batch Finished: Change Auditor Agent Deployment Status"
    $Finalbody = " Hi All,

    The Change Auditor Agent Deployment batch has been completed. 
    Batch Start Time : $TodayDate EST
    Batch Completion Time : $FinalTime EST

    Please Check the output files in Script Run Directory.
    Results Summary:

    Total Agents to be Installed                 : $TotalServers

    Total Agents Successfully Installed          : $Success

    Total Servers with Failed Connectivity Test  : $Unreachable

    Total Servers with .Net Framework below 4.5.2: $FrameworkIssue

    Total Servers with Failed Installation Result: $ActionNeeded

    Regards,
    PowerShell Admin"

    Send-MailMessage -From $from -Body $Finalbody -To $to -cc $cc -Subject $subject1 -SmtpServer $smtphost