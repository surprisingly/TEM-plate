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

#code block to get data from csv
function get-temdata {
	$global:csvpath = ''
	$files[0].split("\")[0..($files[0].split("\").length-2)] | % {$global:csvpath += $_ + "\"}
	set-location $global:csvpath
	$global:csvdata = import-csv (get-childitem tj_etch*.csv | ? {(get-content $_)[0] -like "Image*"} | sort -descending lastwritetime | select -first 1)
	#The keys below are going to match the filename in $files
	$global:csv_keys = $csvdata | % {$_.Path.split("\")[-1]}
	#The filename keys are below
	$global:file_keys = $files | % {$_.split("\")[-1]}
	$global:tem_data = @()
	for ($i=0;$i -lt $accepted_idxs.length;$i++) {
		$global:tem_data += $csvdata | ? {$_.path.split("\")[-1] -like $file_keys[$accepted_idxs[$i]]}
		Add-Member -InputObject $tem_data[$i] -NotePropertyName "Wafer" $tem_data[$i].path.split("\")[2].substring(0,3)
	}
}


#code block to fill the template with the data
function fill-xls {
	$global:temxls = new-object -ComObject excel.application
	$temurl = 'http://sharepoint/Etch_LSI/prc_eng_feol/SiteAssets/Lists/TJRG%20TEM%20Archive/AllItems/14nm_TJ_RG_pMTS_Scaling_Template.xlsx'
	$temxls.Workbooks.Open($temurl,$false,$true)

	for($i=0;$i -lt $accepted_idxs.length;$i++) {
		$d = $tem_data[$i]
		#add and scale the image
		$temxls.cells(3,3+2*$i).select()
		$img = $temxls.activesheet.Shapes.AddPicture($files[$accepted_idxs[$i]],$false,$true,1+$temxls.activecell.left,1+$temxls.activecell.top,-1,-1)
		$img.ScaleHeight(.10,0,0)
		
		$temxls.cells(4,3+2*$i).value = $tem_data[$i].PC2PC
		$temxls.cells(5,3+2*$i).value = $tem_data[$i].Tip2Tip
		$temxls.cells(7,3+2*$i).value = $tem_data[$i].Depth
	}
	$temxls.visible = $true
}

$accept_image = {
    $global:accepted_idxs += $global:current_image_idx
    if ($global:current_image_idx -ge ($num_files-1)) {
        #$global:current_image_idx = 0
        #write-host "reset idx to 0"
        $form.close()
        write-host "Accepted the following images:"
        $accepted_idxs | % {write-host $files[$_]}
		get-temdata
		fill-xls
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
		get-temdata
		fill-xls
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
write-host "complete"
