Add-Type -AssemblyName PresentationFramework

$JobScripts = @{}
$JobScripts.Window = Get-Content -LiteralPath $PSScriptRoot\Window.job.ps1 -Raw
$JobScripts.ImageRetriever = Get-Content -LiteralPath $PSScriptRoot\ImageRetriever.job.ps1 -Raw

function Start-VMConsoleViewer {
  [CmdletBinding(
    PositionalBinding = $false
  )]
  param(
    [Parameter(
      Position = 0,
      Mandatory = $true
    )]
    [Microsoft.HyperV.PowerShell.VirtualMachine[]]
    $VMs,

    [ValidateSet(640, 800)]
    [int]
    $ImageWidth = 800
  )

  function New-UIPowerShell {
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.Open()
  
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
  
    $ps
  }

  try {
    $Threads = @{ # Runspaces hosting WPF *must* be STA; hence the function.
      Window = New-UIPowerShell
      ImageRetriever = New-UIPowerShell
    }

    # Once synchronized, hashtable values are shared between all in-process
    # runspaces where it is used.
    $SynchronizedData = [hashtable]::Synchronized(@{
      ImageRetrieverState = "Init"
      ConfirmWindowExit = $true
    })

    $cmd = [System.Management.Automation.Runspaces.Command]::new(
      $script:JobScripts.Window,
      $true # Script, not command.
    )
    $cmd.Parameters.Add("VMs", $VMs) # VMs[0] will be the initially selected VM.
    $cmd.Parameters.Add("ImageWidth", $ImageWidth)
    $cmd.Parameters.Add("SynchronizedData", $SynchronizedData)
    
    $Threads.Window.Commands.AddCommand($cmd) |
      Out-Null
      
    $cmd = [System.Management.Automation.Runspaces.Command]::new(
      $script:JobScripts.ImageRetriever,
      $true # Script, not command.
    )
    $cmd.Parameters.Add("ImageMode", "FullResolution")
    $cmd.Parameters.Add("RetrievalInterval", 500) # ms between image updates.
    $cmd.Parameters.Add("SynchronizedData", $SynchronizedData)
    
    $Threads.ImageRetriever.Commands.AddCommand($cmd) |
      Out-Null
      
    $Threads.ImageRetriever.BeginInvoke() | Out-Null
    $Threads.Window.BeginInvoke() | Out-Null

    [PSCustomObject]@{
      PSTypeName       = "VMConsoleViewerData"
      Threads          = $Threads
      SynchronizedData = $SynchronizedData
    }
  } catch {
    $PSCmdlet.ThrowTerminatingError($_)
  }
}

function Stop-VMConsoleViewer {
  [CmdletBinding(
    PositionalBinding = $false
  )]
  param(
    [Parameter(
      Position = 0,
      Mandatory = $true
    )]
    [PSTypeName("VMConsoleViewerData")]
    [PSCustomObject]
    $VMConsoleViewerData
  )
  try {
    # This override of Dispatcher.Invoke is very forgiving -- if the action
    # cannot be dispatched within the timeout it simply moves on; it doesn't
    # even throw an exception.
    $VMConsoleViewerData.SynchronizedData.Window.Dispatcher.Invoke(
    [System.Windows.Threading.DispatcherPriority]::Normal,
    [timespan]::FromSeconds(1),
    [action]{
      $VMConsoleViewerData.SynchronizedData.Window.Close()
    })
    
    do { # This is facile -- mature code would loop ~4-5 times, then throw.
      Start-Sleep -Seconds 1
    } until (
      $VMConsoleViewerData.SynchronizedData.ImageRetrieverState -eq "Off" -and
      $VMConsoleViewerData.Threads.Window.InvocationStateInfo.State -eq "Completed" -and
      $VMConsoleViewerData.Threads.ImageRetriever.InvocationStateInfo.State -eq "Completed"
    )
    
    # Runspaces exist in the global scope; those not explicitly disposed of
    # accumulate. In PowerShell 5, you can observe this using 'Get-Runspace'.
    $VMConsoleViewerData.Threads.GetEnumerator() |
      ForEach-Object {
        $_.Value.Runspace.Dispose()
      }
  } catch {
    $PSCmdlet.ThrowTerminatingError($_)
  }
}