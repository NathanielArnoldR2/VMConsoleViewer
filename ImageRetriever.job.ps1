[CmdletBinding()]
param(
  [Parameter(
    Mandatory = $true
  )]
  [ValidateSet("Thumbnail", "FullResolution")]
  [string]
  $ImageMode,

  [Parameter(
    Mandatory = $true
  )]
  [ValidateRange(0, 1000)]
  [int]
  $RetrievalInterval,

  [Parameter(
    Mandatory = $true
  )]
  [hashtable]
  $SynchronizedData
)
try {
  if ($SynchronizedData.ImageRetrieverState -eq "Init") {
    $PSDefaultParameterValues = @{
      "Get-WmiObject:Namespace" = "root\virtualization\v2"
    }
    
    $wmi = @{
      VMMS        = Get-WmiObject -Class Msvm_VirtualSystemManagementService
      VM          = $null
      VMSummary   = $null
      VMVideoHead = $null
    }

    $SynchronizedData.ImageRetrieverState = "Running"
  }

  function Show-NoVideo {
    # Image is collapsed; "No Video" text is made visible. If this dispatch
    # fails, it is silent. No harm, no foul. :-)
    $SynchronizedData.Text.Dispatcher.Invoke(
      [System.Windows.Threading.DispatcherPriority]::Normal,
      [timespan]::FromSeconds(1),
      [action]{
        $SynchronizedData.Image.Visibility = [System.Windows.Visibility]::Collapsed
        $SynchronizedData.Text.Visibility = [System.Windows.Visibility]::Visible
      }
    )
  }
  
  while ($true) {
    if ($SynchronizedData.ImageRetrieverState -eq "Stop") {
      break
    }

    if ($RetrievalInterval -gt 0) {
      Start-Sleep -Milliseconds $RetrievalInterval
    }

    if (
      $null -eq $SynchronizedData.Image -or
      $null -eq $SynchronizedData.Text
    ) {
      continue # Window not shown *or* window closed.
    }

    if ($null -eq $SynchronizedData.VMId) {
      Show-NoVideo # VM not selected.
      continue
    }

    # Initial loop or when signal to change VM feed was received. Either way,
    # everything needs to be retrieved / calculated again.
    if ($null -eq $wmi.VM -or $wmi.VM.Name -ne $SynchronizedData.VMId) {
      $wmi.VM = Get-WmiObject -Class Msvm_ComputerSystem -Filter "Name = `"$($SynchronizedData.VMId)`""
  
      $wmi.VMSummary = $null
      $wmi.VMVideoHead = $null
  
      $i = -1
  
      $imageWidth = 0
      $imageHeight = 0
      $stride = 0
    }
  
    $i++
  
    if ($ImageMode -eq "Thumbnail") {
      if ($null -eq $wmi.VMSummary) {
        $sumObj = $wmi.VM.GetRelated("Msvm_SummaryInformation").Path
  
        $wmi.VMSummary = $sumObj.Path
      } else {
        $sumObj = [wmi]$wmi.VMSummary
      }
    
      if ($null -eq $sumObj.ThumbnailImage) {
        Show-NoVideo
        continue
      }
  
      $imageData   = $sumObj.ThumbnailImage
      $imageWidth  = $sumObj.ThumbnailImageWidth
      $imageHeight = $sumObj.ThumbnailImageHeight
    }
    elseif ($ImageMode -eq "FullResolution") {

      # The call to ascertain dimension of the vm feed is *expensive*, so we
      # cache it for up to 5 iterations. We cannot do so indefinitely, as the
      # dimensions can change if (e.g.) the user changes the screen resolution.
      if ($imageWidth,$imageHeight -eq 0 -or ($i % 5) -eq 0) {
        if ($null -eq $wmi.VMVideoHead) {
          $vhObj = $wmi.VM.GetRelated("Msvm_VideoHead")
        } else {
          $vhObj = [wmi]$wmi.VMVideoHead
        }

        if ($null -eq $vhObj) {
          Show-NoVideo
          continue
        }

        if ($null -eq $wmi.VMVideoHead) {
          $wmi.VMVideoHead = $vhObj.Path
        }

        if (
          [int]$vhObj.CurrentHorizontalResolution -eq 0 -or
          [int]$vhObj.CurrentVerticalResolution -eq 0
        ) {
          Show-NoVideo
          continue
        }
  
        $imageWidth = $vhObj.CurrentHorizontalResolution
        $imageHeight = $vhObj.CurrentVerticalResolution
      }
  
      $imageData = $wmi.VMMS.GetVirtualSystemThumbnailImage(
        $wmi.VM,
        $imageWidth,
        $imageHeight
      ).ImageData

      if ($null -eq $imageData) {
        Show-NoVideo
        continue
      }
    }

    # For reasons unknown, four unnecessary bytes are prepended to the image
    # data; they will corrupt the image if not stripped. System.Drawing
    # methods for constructing/saving a bmp from this data stream are not
    # affected by these four extraneous bytes. Again, reason is unknown.
    $selectedData = [byte[]]::new($imageData.Count - 4)
    [array]::ConstrainedCopy(
      $imageData,
      4,
      $selectedData,
      0,
      $selectedData.Length
    )
  
    $pixelFormat = [System.Windows.Media.PixelFormats]::Bgr565
  
    # Cached on the same terms as image dimensions -- though probably not
    # expensive enough to merit it.
    if ($stride -eq 0 -or ($i % 5) -eq 0) {
      $stride = (($imageWidth * $pixelFormat.BitsPerPixel + 31) -band (-bnot 31)) / 8
    }

    # If this dispatch fails, it is silent. No harm, no foul; we just try again
    # next time. :-)
    $SynchronizedData.Image.Dispatcher.Invoke(
      [System.Windows.Threading.DispatcherPriority]::Normal,
      [timespan]::FromSeconds(1),
      [action]{

        # Cached on same terms as image dimensions/stride. If the image is
        # larger than our viewing dimensions, we'll "shrink" it uniformly
        # to fit the available size; if it's smaller, we're not going to
        # blow it up.
        if (($i % 5) -eq 0) {
          if (
            $imageWidth -lt $SynchronizedData.Image.MinWidth -and
            $imageHeight -lt $SynchronizedData.Image.MinHeight
          ) {
            $SynchronizedData.Image.Stretch = [System.Windows.Media.Stretch]::None
          } else {
            $SynchronizedData.Image.Stretch = [System.Windows.Media.Stretch]::Uniform
          }
        }

        # Since an image source must be constructed in the thread where it is
        # used, we must do so in this dispatch.                
        $SynchronizedData.Image.Source = [System.Windows.Media.Imaging.BitmapSource]::Create(
          $imageWidth,
          $imageHeight,
          96,
          96,
          $pixelFormat,
          $null,
          $selectedData,
          $stride
        )

        # "No Video" text is collapsed; Image made visible.
        $SynchronizedData.Text.Visibility = [System.Windows.Visibility]::Collapsed
        $SynchronizedData.Image.Visibility = [System.Windows.Visibility]::Visible
      }
    )
  }

  $SynchronizedData.ImageRetrieverState = "Off"
} catch {
  $PSCmdlet.ThrowTerminatingError($_)
}