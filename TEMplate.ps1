Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.Text = "Select Images"
$file_select = New-Object System.Windows.Forms.OpenFileDialog
$file_select.multiselect = $true
$file_select.showDialog()
$files = $file_select.filenames
write-host $files.length
$images = $files | % {[system.drawing.image]::FromFile($_)}
write-host $images
$form.backgroundimage = $images[0]
$form.size = $form.backgroundimage.size

$form.ShowDialog()
