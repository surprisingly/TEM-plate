#Load forms
Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "Select Images"

$instructs = new-object System.Windows.Forms.Label
$instructs.backcolor = "Transparent"
$instructs.text = "Select the images"
$instructs.left = 10
$instructs.top = 10

$form.controls.add($instructs)

$file_select = New-Object System.Windows.Forms.OpenFileDialog
$file_select.multiselect = $true
$file_select.showDialog()

$files = $file_select.filenames 
$images = $files | % {[system.drawing.image]::FromFile($_)}
$num_files = $files.length
write-host $num_files

#use global variable format to enforce scope
$global:current_image_idx = 0

$form.backgroundimage = $images[$current_image_idx]
$form.size = $form.backgroundimage.size

$okclick = {
    if ($global:current_image_idx -ge ($num_files-1)) {
        $global:current_image_idx = 0
        write-host "reset idx to 0"
    } else {
        write-host "did else"
        $global:current_image_idx = $global:current_image_idx + 1
    }
    $form.backgroundimage = $images[$current_image_idx]
    $form.size = $images[$current_image_idx].size
    write-host $current_image_idx
}

$chgbut = new-object System.Windows.Forms.Button
$form.controls.add($chgbut)
$chgbut.top = 40
$chgbut.left = 40
$chgbut.add_click($okclick)
$form.ShowDialog()
