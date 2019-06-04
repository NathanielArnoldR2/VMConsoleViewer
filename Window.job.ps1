[CmdletBinding()]
param(
  [Parameter(
    Mandatory = $true
  )]
  [Microsoft.HyperV.PowerShell.VirtualMachine[]]
  $VMs,

  [Parameter(
    Mandatory = $true
  )]
  [ValidateSet(640, 800)]
  [int]
  $ImageWidth,

  [Parameter(
    Mandatory = $true
  )]
  [hashtable]
  $SynchronizedData
)
try {
  $VMListItems = @(
    $VMs |
      Select-Object Name,Id
  )

  # The window is equipped to handle image output w/ 4:3 and 16:9 aspect ratio;
  # other ratios (5:4, e.g. 1280x1024) may have white "letterbox" bars. The "No
  # Video" text notification is "locked" at the 4:3 ratio.
  $ImageMinHeight = $ImageWidth / 16 * 9
  $ImageMaxHeight = $ImageWidth / 4 * 3

  $windowXml = [xml]@"
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  WindowStyle="None"
  ShowInTaskbar="False"
  SizeToContent="WidthAndHeight"
  ResizeMode="NoResize">
    <StackPanel>
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition />
                <ColumnDefinition Width="Auto" />
            </Grid.ColumnDefinitions>
            <ComboBox
              x:Name="VMList"
              HorizontalContentAlignment="Center"
              VerticalContentAlignment="Center"
              Padding="10"/>
            <Button
              x:Name="ExitButton"
              Grid.Column="1"
              Content="Exit"
              Padding="10"/>
        </Grid>
        <Grid
          x:Name="TextGrid"
          Visibility="Visible"
          Width="$ImageWidth"
          Height="$ImageMaxHeight"
          Background="Black">
            <TextBlock
              x:Name="Text"
              HorizontalAlignment="Center"
              VerticalAlignment="Center"
              Foreground="White"
              Text="No Video"/>
        </Grid>
        <Image
          x:Name="Image"
          Visibility="Collapsed"
          MinWidth="$ImageWidth"
          MaxWidth="$ImageWidth"
          MinHeight="$ImageMinHeight"
          MaxHeight="$ImageMaxHeight"/>
    </StackPanel>
</Window>
"@

  $SynchronizedData.Window = $Window = [System.Windows.Markup.XamlReader]::Load(
    [System.Xml.XmlNodeReader]::new($windowXml)
  )
  $Window.Add_Closing({
    # Signal ImageRetriever to transition to "Off" @ next iteration.
    $SynchronizedData.ImageRetrieverState = "Stop"
  })

  $Window.FindName("VMList") |
    ForEach-Object {
      $_.FontSize = $_.FontSize * 2
      $_.ItemsSource = $VMListItems
      $_.DisplayMemberPath = "Name"
      $_.SelectedValuePath = "Id"
      $_.Add_SelectionChanged({
        param($obj, $evtArgs)
      
        $Window.Title = $obj.SelectedItem.Name

        # Signal ImageRetriever to change VM feed @ next iteration.
        $SynchronizedData.VMId = $obj.SelectedItem.Id
      })
    }

  $Window.FindName("ExitButton") |
    ForEach-Object {
      $_.FontSize = $Window.FindName("VMList").FontSize
      $_.Add_Click({
        if ($SynchronizedData.ConfirmWindowExit) {
          $result = [System.Windows.MessageBox]::Show(
            "Are you sure you want to exit?",
            "Confirm Exit",
            [System.Windows.MessageBoxButton]::OKCancel,
            [System.Windows.MessageBoxImage]::Question,
            [System.Windows.MessageBoxResult]::Cancel,
            [System.Windows.MessageBoxOptions]::None
          )
        }
      
        if (
          $SynchronizedData.ConfirmWindowExit -eq $false -or
          $result -eq "OK"
        ) {
          $Window.Close()
        }
      })
    }

  $SynchronizedData.Text = $TextGrid = $Window.FindName("TextGrid")
  $TextGrid.Add_MouseDown({
    param($obj, $evtArgs)
  
    if ($evtArgs.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) {
      $Window.DragMove()
    }
  })

  $Window.FindName("Text") |
    ForEach-Object {
      $_.FontSize = $_.FontSize * 4
    }

  $SynchronizedData.Image = $Image = $Window.FindName("Image")
  $Image.Add_MouseDown({
    param($obj, $evtArgs)
  
    if ($evtArgs.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) {
      $Window.DragMove()
    }
  })

  $Window.FindName("VMList").SelectedValue = $VMListItems[0].Id

  $Window.ShowDialog()
} catch {
  $PSCmdlet.ThrowTerminatingError($_)
}