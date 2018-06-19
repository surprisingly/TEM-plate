#Load forms type
Add-Type -AssemblyName System.Windows.Forms

#instantiate the form
$form = New-Object System.Windows.Forms.Form
#set the title of the form
$form.Text = "Select Images"

#instantiate a label to appear on the image
$instructs = new-object System.Windows.Forms.Label
#set the background to yellow
$instructs.backcolor = "Yellow"
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

$global:accepted_images = @()

$accept_image = {
    $global:accepted_images += $files[$current_image_idx]
    $global:current_image_idx = $global:current_image_idx + 1
    if ($global:current_image_idx -ge ($num_files-1)) {
        #$global:current_image_idx = 0
        #write-host "reset idx to 0"
        $form.close()
        write-host "Accepted the following images:"
        $accepted_images | % {write-host $_}
        } else {
    $form.backgroundimage = $images[$current_image_idx]
    $form.size = $images[$current_image_idx].size
    }
}

$okclick = {
    if ($global:current_image_idx -ge ($num_files-1)) {
        $form.close()
        write-host "Accepted the following images:"
        $accepted_images | % {write-host $_}
    } else {
        $global:current_image_idx = $global:current_image_idx + 1
    }
    $form.backgroundimage = $images[$current_image_idx]
    $form.size = $images[$current_image_idx].size
}

$acceptbut = new-object system.windows.forms.button
$form.controls.add($acceptbut)
$acceptbut.top = 40
$acceptbut.left = $form.width - $acceptbut.width - 40
$acceptbut.text = "Good"
$acceptbut.add_click($accept_image)

$chgbut = new-object System.Windows.Forms.Button
$form.controls.add($chgbut)
$chgbut.top = 40
$chgbut.left = 40
$chgbut.text = "Next"
$chgbut.add_click($okclick)
$form.ShowDialog()
