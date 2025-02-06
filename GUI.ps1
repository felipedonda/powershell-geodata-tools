### Global variables

# ADD_DATA_TO_KML - Add the data from an Excel spreadsheet to a KML file.
# ADD_DATA_TO_GEOJSON - Add the data from an Excel spreadsheet to a GeoJSON file.
# CONVERT_KML_TO_GEOJSON - Convert a KML file to a GeoJSON file.
# CONVERT_GEOJSON_TO_KML - Convert a GeoJSON file to a KML file.


$global:Mode = "ADD_DATA_TO_KML"

$global:ExcelFile = New-Object PSObject
$global:ExcelNameField = ""
$global:KmlFile = New-Object PSObject
$global:GeoJsonFile = New-Object PSObject
$global:ExcelData = @()
$global:KmlData = New-Object System.Xml.XmlDocument
$global:MapData = New-Object PSObject


### Controls Defining

$form = New-Object System.Windows.Forms.Form -Property @{
    Text = "Map Functions - made by Felipe Donda"
    ClientSize = "575,615"
    MinimumSize = "575,615"
}

$Modelabel = New-Object System.Windows.Forms.Label  -Property @{
    Text = "Select the function desired:"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    AutoSize = $false
    Width = 400
    Height = 30
    Location = New-Object System.Drawing.Point(20,20)
}

$Modecombobox = New-Object System.Windows.Forms.ComboBox -Property @{
    MinimumSize = "320,50"
    Font = New-Object System.Drawing.Font("Calibri",14,[System.Drawing.FontStyle]::Regular)
    Anchor = "top, left"
    Location = New-Object System.Drawing.Point(20,50)
}

$Modecombobox.Items.AddRange(@("Add data to KML","Add data to GeoJSON","Convert KML to GeoJSON","Convert GeoJSON to KML"))
$Modecombobox.SelectedIndex = 0


## Kml Controls

$LoKml_pos = 110

$LoKml_label = New-Object System.Windows.Forms.Label  -Property @{
    Text = "Select KML file:"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    AutoSize = $false
    Width = 375
    Height = 30
    Location = New-Object System.Drawing.Point(20,$LoKml_pos)
}

$LoKml_textbox = New-Object System.Windows.Forms.Textbox -Property @{
    MinimumSize = "375,30"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    Anchor = "top, left, right"
    ReadOnly = $true
    Location = New-Object System.Drawing.Point(20,($LoKml_pos+40))
}


$LoKml_fileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = "Google Earth KML (*.kml)|*.kml"
}

$LoKml_loadbutton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Open"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    Size = "70,35"
    Anchor = "top, right"
    Location = New-Object System.Drawing.Point(410,($LoKml_pos+ 38))
}

$LoKml_clearbutton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Clear"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    Size = "70,35"
    Anchor = "top, right"
    Location = New-Object System.Drawing.Point(487,($LoKml_pos+ 38))
}

$LoKml_subLabel = New-Object System.Windows.Forms.Label  -Property @{
    Text = ""
    Font = New-Object System.Drawing.Font("Calibri",11,[System.Drawing.FontStyle]::Regular)
    AutoSize = $false
    Width = 375
    Height = 30
    Location = New-Object System.Drawing.Point(20,($LoKml_pos + 75))
}


## Excel Controls

$LoEx_pos = $LoKml_pos + 130

$LoEx_checkbox = New-Object System.Windows.Forms.CheckBox  -Property @{
    Text = "Add data during conversion"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    AutoSize = $false
    Width = 400
    Height = 30
    Anchor = "top, left"
    Visible = $false
}

$LoEx_label = New-Object System.Windows.Forms.Label  -Property @{
    Text = "Select Excel spreadsheet:"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    AutoSize = $false
    Width = 375
    Height = 30
}

$LoEx_textbox = New-Object System.Windows.Forms.Textbox -Property @{
    MinimumSize = "375,30"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    Anchor = "top, left, right"
    ReadOnly = $true
}

$LoEx_fileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = "SpreadSheet (*.xlsx)|*.xlsx"
}

$LoEx_loadbutton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Open"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    Size = "70,35"
    Anchor = "top, right"
}

$LoEx_clearbutton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Clear"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    Size = "70,35"
    Anchor = "top, right"
}

$LoEx_subLabel = New-Object System.Windows.Forms.Label  -Property @{
    Text = ""
    Font = New-Object System.Drawing.Font("Calibri",11,[System.Drawing.FontStyle]::Regular)
    AutoSize = $false
    Width = 375
    Height = 30
    
}

## Out Controls

$Out_pos = $LoEx_pos + 130

$Out_fileDialog = New-Object System.Windows.Forms.SaveFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = "Google Earth KML (*.kml)|*.kml"
}

$LogBox = New-Object System.Windows.Forms.RichTextBox -Property @{
    MinimumSize = "535,150"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    Anchor = "top, left, bottom, right"
    BorderStyle = "FixedSingle"
    ReadOnly = $true
    Location = New-Object System.Drawing.Point(20,$Out_pos)
}

$Out_button = New-Object System.Windows.Forms.Button -Property @{
    Text = "Save file"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    Size = "140,45"
    Anchor = "Left, bottom"
    Location = New-Object System.Drawing.Point(20,($Out_pos + 180))
}


$ReportPanel = New-Object System.Windows.Forms.Panel  -Property @{
    Visible = $false
    AutoSize = $false
    Anchor = "top, left, bottom, right"
    BorderStyle = "FixedSingle"
    Width = $form.Width - 80
    Height = $form.Height - 100
    Location = New-Object System.Drawing.Point(30,30)
}

$ReportBox = New-Object System.Windows.Forms.RichTextBox  -Property @{
    AutoSize = $false
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    Anchor = "top, left, bottom, right"
    ReadOnly = $true
    Width = $ReportPanel.Width - 60
    Height = $ReportPanel.Height - 110
    Location = New-Object System.Drawing.Point(30,30)
}

$ReportBox_CloseButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Close"
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    Size = "70,35"

    Anchor = "right, bottom"
    Location = New-Object System.Drawing.Point(($ReportPanel.Width - 100 ),($ReportPanel.Height - 55 ))
}

function Report_Close {
    $ReportPanel.Visible = $false
}

$ReportBox_CloseButton.Add_Click({Report_Close})

$ReportPanel.Controls.AddRange(@(
    $ReportBox,
    $ReportBox_CloseButton
))


$MessagePanel = New-Object System.Windows.Forms.Panel  -Property @{
    Visible = $false
    AutoSize = $false
    BorderStyle = "FixedSingle"
    Anchor = "top, left, bottom, right"
    Width = $form.Width - 120
    Height = $form.Height - 200
    Location = New-Object System.Drawing.Point(50,100)
}

$MePa_Label = New-Object System.Windows.Forms.Label  -Property @{
    Text = "Loading..."
    Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)
    AutoSize = $false
    Anchor = "top, left, bottom, right"
    TextAlign = "MiddleCenter"
    Width = 200
    Height = 30
    Location = New-Object System.Drawing.Point(($MessagePanel.Width/2 - 100),($MessagePanel.Height/2 - 15))
}

$MessagePanel.Controls.Add($MePa_Label)

### Setting up functions

## Log function

function Write-Log
{
    param(
        [Parameter(Position = 1, ValueFromRemainingArguments)]
        [string[]]$LogText,
        [CmdletBinding(PositionalBinding  = $false)]
        [ValidateSet("Info", "Warning", "Error")]
        $Type = "Info",
        [CmdletBinding(PositionalBinding  = $false)]
        [switch]$DisplayMessageBox
    )
    Switch($Type)
    {
        "Info" {
            Write-Host ("[" + (Get-Date -Format "hh:mm:ss") + "]") $LogText
            $LogBox.AppendText(("[" + (Get-Date -Format "hh:mm:ss") + "] " + $LogText + "`n"))
            if($DisplayMessageBox) {
                [System.Windows.Forms.MessageBox]::Show($LogText,'Info','Ok','Info') | Out-Null
            }
        }
        "Warning" {
            Write-Host ("[" + (Get-Date -Format "hh:mm:ss") + "]") $LogText
            $LogBox.AppendText(("[" + (Get-Date -Format "hh:mm:ss") + "] " + $LogText + "`n"))
            if($DisplayMessageBox) {
                [System.Windows.Forms.MessageBox]::Show($LogText,'Warning','Ok','Warning') | Out-Null
            }
        }
        "Error" {
            Write-Host ("[" + (Get-Date -Format "hh:mm:ss") + "] ERROR:") $LogText -ForegroundColor Yellow
            $LogBox.SelectionStart = $LogBox.TextLength
            $LogBox.SelectionLength = 0
            $LogBox.SelectionColor = [System.Drawing.Color]::Red
            $LogBox.AppendText(("[" + (Get-Date -Format "hh:mm:ss") + "] ERROR: " + $LogText + "`n"))
            $LogBox.SelectionColor = $LogBox.ForeColor

            if($DisplayMessageBox) {
                [System.Windows.Forms.MessageBox]::Show($LogText,'Error','Ok','Error') | Out-Null
            }
        }
    }
    
}


## Place Excel Controls

function Place-ExcelControls {
    param(
        $Offset = 0,
        [Switch]$CheckboxVisible
    )
    if($CheckboxVisible) {
        $Offset += 40
    }
    $LoEx_checkbox.Location = New-Object System.Drawing.Point(20,($Offset-40))
    $LoEx_label.Location = New-Object System.Drawing.Point(20,$Offset)
    $LoEx_textbox.Location = New-Object System.Drawing.Point(20,($Offset + 40))
    $LoEx_loadbutton.Location = New-Object System.Drawing.Point(487,($Offset + 38))
    $LoEx_clearbutton.Location = New-Object System.Drawing.Point(20,($Offset + 38))
    $LoEx_subLabel.Location = New-Object System.Drawing.Point(20,($Offset + 75))
}

## Change Mode

function Change-Mode {
    param(
           [ValidateSet("ADD_DATA_TO_KML","ADD_DATA_TO_GEOJSON","CONVERT_KML_TO_GEOJSON","CONVERT_GEOJSON_TO_KML")]$Mode
    )

    # ADD_DATA_TO_KML - Add the data from an Excel spreadsheet to a KML file.
    # ADD_DATA_TO_GEOJSON - Add the data from an Excel spreadsheet to a GeoJSON file.
    # CONVERT_KML_TO_GEOJSON - Convert a KML file to a GeoJSON file.
    # CONVERT_GEOJSON_TO_KML - Convert a GeoJSON file to a KML file.

    <#
    $MessagePanel,
    $ReportPanel,
    $Modelabel,
    $Modecombobox,
    $LoKml_label,
    $LoKml_textbox,
    $LoKml_loadbutton,
    $LoKml_clearbutton,
    $LoKml_subLabel,
    $LoEx_checkbox,
    $LoEx_label,
    $LoEx_textbox,
    $LoEx_loadbutton,
    $LoEx_clearbutton,
    $LoEx_subLabel,
    $LogBox,
    $Out_button
    #>

    $global:Mode = "ADD_DATA_TO_KML"
    switch($Mode) {
        "ADD_DATA_TO_KML" {
            $LoKml_label.Visible = $true
            $LoKml_textbox.Visible = $true
            $LoKml_loadbutton.Visible = $true
            $LoKml_clearbutton.Visible = $true
            $LoKml_subLabel.Visible = $true
            $LoEx_checkbox.Visible = $false
            $LoEx_label.Visible = $true
            $LoEx_textbox.Visible = $true
            $LoEx_loadbutton.Visible = $true
            $LoEx_clearbutton.Visible = $true
            $LoEx_subLabel.Visible = $true

        }
        "CONVERT_KML_TO_GEOJSON" {
            $LoKml_label.Visible = $true
            $LoKml_textbox.Visible = $true
            $LoKml_loadbutton.Visible = $true
            $LoKml_clearbutton.Visible = $true
            $LoKml_subLabel.Visible = $true
            $LoEx_checkbox.Visible = $true
            $LoEx_label.Visible = $true
            $LoEx_textbox.Visible = $true
            $LoEx_loadbutton.Visible = $true
            $LoEx_clearbutton.Visible = $true
            $LoEx_subLabel.Visible = $true
        }
    }
}

## Excel Clear Button

function Clear_Excel {
    $LoEx_textbox.Text = ""
    $LoEx_subLabel.Text = ""
    $global:ExcelNameField = ""
    $global:ExcelFile = New-Object PSObject
    $global:ExcelData = @()

}

$LoEx_clearbutton.Add_Click({Clear_Excel})

## Excel Load Button

function LoEx_loadbutton_click
{
    if($LoEx_fileDialog.ShowDialog() -eq "OK") {
        Clear_Excel
        
        Write-Log 'Loading Excel file "',$LoEx_fileDialog.FileName,'"'

        $global:ExcelFile = Get-ChildItem $LoEx_FileDialog.FileName
        $LoEx_fileDialog.InitialDirectory = $global:ExcelFile.Directory.FullName
        $LoKml_fileDialog.InitialDirectory = $global:ExcelFile.Directory.FullName
        $Out_fileDialog.InitialDirectory = $global:ExcelFile.Directory.FullName

        $MessagePanel.Visible = $true

        try {
            $global:ExcelData = Read-Excel $global:ExcelFile.FullName -Verbose
            if($global:ExcelData.count -lt 1) {
                Write-Log 'Error reading file "' $global:ExcelFile.Name '"' ". Invalid data format." -Type Warning -DisplayMessageBox
                Clear_Excel
            } else {
                

                if($Data.Name) {
                    $global:ExcelNameField = "Name"
                } else {
                    if($Data.Address) {
                        $global:ExcelNameField = "Address"
                    } else {
                        Write-Log 'Error reading file "' $global:ExcelFile.Name '"' ". Couldn't find a naming field to distinguished the data. The spreadsheet requires a column named 'Name' or 'Address'." -Type Warning -DisplayMessageBox
                    }
                }


                Write-Log 'File "' $global:ExcelFile.Name '" loaded successfully. ' $global:ExcelData.Count ' rows loaded.'
                $LoEx_textbox.Text = $global:ExcelFile.FullName
                $LoEx_subLabel.Text = ("" + $global:ExcelData.Count + " rows loaded.")
            }
            $MessagePanel.Visible = $false
        } catch {
            Write-Log $_ -Type Error -DisplayMessageBox
            Clear_Excel
            $MessagePanel.Visible = $false
        }
    }
}

$LoEx_loadbutton.Add_Click({LoEx_loadbutton_click})

## Kml Clear Button

function Clear_Kml {
    $LoKml_textbox.Text = ""
    $LoKml_subLabel.Text = ""

    $global:KmlFile = New-Object PSObject
    $global:KmlData = New-Object System.Xml.XmlDocument
    $global:MapData = New-Object PSObject
}

$LoKml_clearbutton.Add_Click({Clear_Kml})


## Kml Load Button

function LoKml_loadbutton_click {
    if($LoKml_fileDialog.ShowDialog() -eq "OK") {
        Clear_Kml
        Write-Log 'Loading KML file "',$LoKml_fileDialog.FileName,'"'

        $global:KmlFile = Get-ChildItem $LoKml_FileDialog.FileName
        $LoKml_fileDialog.InitialDirectory = $global:KmlFile.Directory.FullName
        $LoEx_fileDialog.InitialDirectory = $global:KmlFile.Directory.FullName
        $Out_fileDialog.InitialDirectory = $global:KmlFile.Directory.FullName
        $MessagePanel.Visible = $true

        try {
            $global:KmlData = [System.Xml.XmlDocument](Get-Content $global:KmlFile.FullName)

            <#
                if(!$NameField) {
                    if($Data.Name) {
                        $NameField = "Name"
                    } else {
                        if($Data.Address) {
                            $NameField = "Address"
                        } else {
                            throw "Couldn't find a naming field to distinguished the data. Either add a 'Name' or 'Address' column or specify a different column with the '-NameField' argument."
                        }
                    }
                }
            #>

            if($global:KmlData.kml.Document.name -eq $null) {
                Clear_Kml
                Write-Log 'Error reading file "' $global:KmlFile.Name '". Invalid data format.' -Type Warning -DisplayMessageBox
            } else {
                $global:MapData = $global:KmlData | Kml-ToMap
                $LoKml_textbox.Text = $global:KmlFile.FullName
                $LoKml_subLabel.Text = ("" + ($global:MapData | Count-MapPlacemarks -Type "Polygon") + " polygons and " + ($global:MapData | Count-MapPlacemarks -Type "Point") + " points loaded.")
            }
            $MessagePanel.Visible = $false
        } catch {
            Write-Log $_  -Type Error -DisplayMessageBox
            Clear_Kml
            $MessagePanel.Visible = $false
        }
        
    }
}

$LoKml_loadbutton.Add_Click({LoKml_loadbutton_click})

## Out file Button

function Out_button_click {
    try {
        switch($global:Mode)
        {
            # ADD_DATA_TO_KML - Add the data from an Excel spreadsheet to a KML file.
            # ADD_DATA_TO_GEOJSON - Add the data from an Excel spreadsheet to a GeoJSON file.
            # CONVERT_KML_TO_GEOJSON - Convert a KML file to a GeoJSON file.
            # CONVERT_GEOJSON_TO_KML - Convert a GeoJSON file to a KML file.
            "ADD_DATA_TO_KML" {
                $checklist = $true
                try {
                    if($global:ExcelFile.FullName -eq $null) {
                        throw 'Error adding data to KML: No Excel file loaded.'
                    }

                    if($global:ExcelData.Count -lt 1) {
                        throw 'Error adding data to KML: Invalid data from Excel file "' + $global:ExcelFile.Name + '".'
                    }

                    if(-not $global:KmlFile.Name) {
                        throw 'Error adding data to KML: No Kml file loaded.'
                    }
                         
                    if($global:KmlData.kml.Document.name -eq $null){
                        throw 'Error adding data to KML: Error reading file "' + $global:KmlFile.Name + '". Invalid data format.'
                    }


                    if($Out_fileDialog.ShowDialog() -eq "OK") {
                        
                        $Out_fileDialog.FileName
                        $Out_fileDialog.InitialDirectory = $global:KmlFile.Directory.FullName

                        $affectedReportData = @()

                        foreach($row in $global:ExcelData){
                            
                            $affected = $global:MapData.RootFolder | SearchReplace-MapPlacemark -Data $row -NameField $global:ExcelNameField -BulletPointField "Comments"
                            $reportDataRow = New-Object PSObject -Property @{
                                Name = $row."$global:ExcelNameField"
                                Affected = $affected
                            }
                            $affectedReportData += $reportDataRow
                        }

                        Write-Log ("Add data do KML | Placemarks affected: " + ($affectedReportData | Where {$_.Affected -gt 0}).Count +
                            ". Placemarks unaffected: " +
                            (($global:MapData | Count-MapPlacemarks) - ($affectedReportData | Measure -Sum Affected).Sum))

                        # Generating report
                        $placemarks = $global:MapData | Extract-MapPlacemarks
                        $Report = ("Placemarks affected: " + ($affectedReportData | Where {$_.Affected -gt 0}).Count +
                            ". Placemarks unaffected: " +
                            (($global:MapData | Count-MapPlacemarks) - ($affectedReportData | Measure -Sum Affected).Sum))
                        $Report += "`n`nList of data added to placemarks:"
                        $affectedReportData | foreach {$Report += " - " + $_.Name + ": " + $_.Affected + " placemarks affected.`n"}
                        $Report += "`nList of unnafected Placemarks:`n"
                        $placemarks | Where {$_.Name -notin $affectedReportData.Name} | foreach {$Report += " - " + $_.Name + " (" + $_.Type +")`n"}
                        
                        # Report done. Converting and saving Kml
                        $Output = $global:MapData | Map-ToKml
                        $Output.OuterXml | Out-File $Out_fileDialog.FileName -Encoding utf8 -Force

                        #Showing report. (will be added to button in the future)
                        $ReportBox.Text = $Report
                        $ReportPanel.Visible = $true
                        Write-Log 'Add data do KML | Saving file"',$Out_fileDialog.FileName,'"'
                        
                    }

                }
                catch {
                    if($_.Exception.WasThrownFromThrowStatement) {
                        Write-Log $_ -Type Warning -DisplayMessageBox
                    } else {
                        Write-Log $_ -Type Error -DisplayMessageBox
                    }
                    
                    $checklist =  $false
                }

            }
            "ADD_DATA_TO_GEOJSON" {}
            "CONVERT_KML_TO_GEOJSON" {}
            "CONVERT_GEOJSON_TO_KML" {}

        }
    } catch {
        Write-Log $_ -Type Error -DisplayMessageBox
    }
}

$Out_button.Add_Click({Out_button_click})

### Adding everything to the form

$form.Controls.AddRange(@(
    $MessagePanel,
    $ReportPanel,
    $Modelabel,
    $Modecombobox,
    $LoKml_label,
    $LoKml_textbox,
    $LoKml_loadbutton,
    $LoKml_clearbutton,
    $LoKml_subLabel,
    $LoEx_checkbox,
    $LoEx_label,
    $LoEx_textbox,
    $LoEx_loadbutton,
    $LoEx_clearbutton,
    $LoEx_subLabel,
    $LogBox,
    $Out_button
    ))

### Starting form

try{
    Write-Host $form.ShowDialog()
}
catch {
    Write-Log $_ -Type Error -DisplayMessageBox
    $form.Dispose()
}
finally {
    $form.Dispose()
}
