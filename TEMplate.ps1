#Load forms type
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#instantiate the form
$form = New-Object System.Windows.Forms.Form
#set the title of the form
$form.Text = "Select Images"

#instantiate a label to appear on the image
#$instructs = new-object System.Windows.Forms.Label
#set the instruction background to yellow
#$instructs.backcolor = "Yellow"
#set the instruction text
#$instructs.text = "Select the images"
#set alignment of the instruction text
#$instructs.textalign = "MiddleCenter"
#set x position of the instructions
#$instructs.left = 10
#set y position of the instructions
#$instructs.top = 10

#add the instructions to the form
#$form.controls.add($instructs)

#instantiate a file selection dialog box
$file_select = New-Object System.Windows.Forms.OpenFileDialog
#allow selection of multiple files
$file_select.multiselect = $true
#open the file selection dialog box
$file_select.showDialog()

#collect the selected filenames into a var
$files = $file_select.filenames 
#create an array of bitmap images
$images = $files | % {[system.drawing.image]::FromFile($_)}
#get the number of selected files
$num_files = $files.length
#DEBUG echo the number of selected files
write-host $num_files

#set an index parameter, initially 0
#global is needed for any var that will be modified in and outside of a block scope
$global:current_image_idx = 0

#set the background of the form to the current image
function update-bgimage {
    $i_h = $images[$global:current_image_idx].Height
    $i_w = $images[$global:current_image_idx].Width
    $pic_box = new-object -typename System.Windows.Forms.PictureBox
    $pic_box.Left = 0
    $pic_box.Top = 0
    $pic_box.height = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height
    $pic_box.width = [math]::floor(($i_w)*($pic_box.height)/($i_h))
    $pic_box.image = $images[$global:current_image_idx]
    $pic_box.sizemode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    $form.controls.add($pic_box)
    $pic_box.bringtofront()
    #set the size of the form to match the scaled image (to the monitor height)
    $form.size = $pic_box.size
    try {
        $acceptbut.bringtofront()
        $chgbut.bringtofront()
    } catch {
    }
}
update-bgimage

$global:accepted_idxs = @()

$accept_image = {
    $global:accepted_idxs += $global:current_image_idx
    if ($global:current_image_idx -ge ($num_files-1)) {
        #$global:current_image_idx = 0
        #write-host "reset idx to 0"
        $form.close()
        write-host "Accepted the following images:"
        $accepted_idxs | % {write-host $files[$_]}
    } else {
        $global:current_image_idx += 1
        update-bgimage
        #$form.backgroundimage = $images[$current_image_idx]
        #$form.size = $images[$current_image_idx].size
    }
}

$reject_image = {
    
    if ($global:current_image_idx -ge ($num_files-1)) {
        $form.close()
        write-host "Accepted the following images:"
        $accepted_idxs | % {write-host $files[$_]}
    } else {
        $global:current_image_idx = $global:current_image_idx + 1
    }
    update-bgimage
    #$form.backgroundimage = $images[$current_image_idx]
    #$form.size = $images[$current_image_idx].size
}

$acceptbut = new-object system.windows.forms.button
$form.controls.add($acceptbut)
$acceptbut.top = 40
$acceptbut.left = $form.width - $acceptbut.width - 40
$acceptbut.text = "Accept"
$acceptbut.add_click($accept_image)
$acceptbut.bringtofront()

$chgbut = new-object System.Windows.Forms.Button
$form.controls.add($chgbut)
$chgbut.top = 40
$chgbut.left = 40
$chgbut.text = "Reject"
$chgbut.add_click($reject_image)
$chgbut.bringtofront()

$form.ShowDialog()  
