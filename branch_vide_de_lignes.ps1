# Si l'éxecution de script est désactivée
#
# Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
#
#


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Générateur de vide de ligne"
$form.Width = 450
$form.Height = 350
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(400, 350)

# TableLayoutPanel for flexible resizing
$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.RowCount = 6
$tableLayout.ColumnCount = 2
$tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 45)))
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 60)))


# Make the row containing the statusRichTextBox flexible
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # Add AutoSize for all rows initially
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # Row for statusRichTextBox takes the remaining space
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))


$form.Controls.Add($tableLayout)

# Name Label
$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = "Numéro de machine :"
$nameLabel.Anchor = 'Left, Right'
$nameLabel.Margin = 10
$tableLayout.Controls.Add($nameLabel, 0, 0) 

# Name RichTextBox (replacing TextBox with RichTextBox)
$nameRichTextBox = New-Object System.Windows.Forms.RichTextBox
$nameRichTextBox.Anchor = 'Top, Left, Right'
$nameRichTextBox.Width = 200
$nameRichTextBox.Height = 25
$nameRichTextBox.Margin = 10
$nameRichTextBox.BorderStyle = 'FixedSingle'
$tableLayout.Controls.Add($nameRichTextBox, 1, 0)

# Number Label
$numberLabel = New-Object System.Windows.Forms.Label
$numberLabel.Text = "Nombre de code :"
$numberLabel.Anchor = 'Left, Right'
$numberLabel.Margin = 10
$tableLayout.Controls.Add($numberLabel, 0, 1)

# Number RichTextBox
$numberRichTextBox = New-Object System.Windows.Forms.RichTextBox
$numberRichTextBox.Anchor = 'Top, Left, Right'
$numberRichTextBox.Width = 200
$numberRichTextBox.Height = 25
$numberRichTextBox.Margin = 10
$numberRichTextBox.BorderStyle = 'FixedSingle'
$tableLayout.Controls.Add($numberRichTextBox, 1, 1)

# Folder Path Label
$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = "Dossier d'enregistrement :"
$folderLabel.Anchor = 'Left, Right'
$folderLabel.Margin = 10
$tableLayout.Controls.Add($folderLabel, 0, 2)

# Folder Path RichTextBox
$folderRichTextBox = New-Object System.Windows.Forms.RichTextBox
$folderRichTextBox.Anchor = 'Top, Left, Right'
$folderRichTextBox.Width = 200
$folderRichTextBox.Height = 25
$folderRichTextBox.Margin = 10
$folderRichTextBox.BorderStyle = 'FixedSingle'
$tableLayout.Controls.Add($folderRichTextBox, 1, 2)

# Browse Button to select folder
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Parcourir"
$browseButton.Anchor = 'Top, Right'
$browseButton.Width = 80
$tableLayout.Controls.Add($browseButton, 1, 3)

# Status RichTextBox (for showing messages)
$statusRichTextBox = New-Object System.Windows.Forms.RichTextBox
$statusRichTextBox.Anchor = 'Top, Left, Right, Bottom'
$statusRichTextBox.Width = 350
$statusRichTextBox.Height = 100
$statusRichTextBox.Multiline = $true
$statusRichTextBox.ReadOnly = $true
$statusRichTextBox.BorderStyle = 'FixedSingle'
$tableLayout.Controls.Add($statusRichTextBox, 0, 4)
$tableLayout.SetColumnSpan($statusRichTextBox, 2)  # Span across 2 columns


# Generate Button
$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "Générer les codes"
$generateButton.Anchor = 'Left, Right,Bottom'
$generateButton.Width = 150
$tableLayout.Controls.Add($generateButton, 0, 5)
$tableLayout.SetColumnSpan($generateButton, 2)  # Span across 2 columns

# Browse Button Click Event to select folder
$browseButton.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq 'OK') {
        $folderRichTextBox.Text = $folderDialog.SelectedPath
    }
})





# Generate Button Click Event to generate QR codes
$generateButton.Add_Click({
    $name = $nameRichTextBox.Text
    $n = [int]$numberRichTextBox.Text
    $folderPath = $folderRichTextBox.Text
    $prefixe = "FAf@|Q7G8TH7bGxg#c!6Jqj_"


    # Validate inputs
    if ([string]::IsNullOrEmpty($name) -or [string]::IsNullOrEmpty($n) -or [string]::IsNullOrEmpty($folderPath)) {
        $statusRichTextBox.AppendText("Veuillez remplir tous les champs." + [Environment]::NewLine)
        return
    }

    # Validate that the number is a valid positive integer
    if (-not ($n -gt 0)) {
        $statusRichTextBox.AppendText("Entrer un nombre valide ( > 0)" + [Environment]::NewLine)
        return
    }

    # Ensure folder exists
    if (-not (Test-Path $folderPath)) {
        $statusRichTextBox.AppendText("Le dossier spécifier n'existe pas. Création en cours..." + [Environment]::NewLine)
        New-Item -ItemType Directory -Path $folderPath
    }

    # Start Barcode Generation
    $statusRichTextBox.AppendText("`rLancement de la génération...`r`n" + [Environment]::NewLine)

    for ($i = 1; $i -le $n; $i++) {
        if ($i -lt 10){ #Add a zero to print 01 instead of 1
            # Concatenate prefix and barcode content
            $barcodeContent = "$prefixe$name"+"0"+"$i"  # Add the prefix to the barcode content
            $tempName = "$name"+"0"+"$i"
            $barcodeImagePath = Join-Path -Path $folderPath -ChildPath "$tempName.png"
        } else {
            # Concatenate prefix and barcode content
            $barcodeContent = "$prefixe$name$i"  # Add the prefix to the barcode content
            $barcodeImagePath = Join-Path -Path $folderPath -ChildPath "$name$i.png"
        }

        # URL encode the content to ensure it's valid for the barcode API (avoids issues with special characters)
        $encodedBarcodeContent = [uri]::EscapeDataString($barcodeContent)

        # Use TEC-IT barcode API
        $barcodeImageUrl = "https://barcode.tec-it.com/barcode.ashx?data=$encodedBarcodeContent&code=Code128&dpi=96&dataseparator="

        # Download the barcode image from the API


        # Log progress
        $statusRichTextBox.AppendText("Génération du code numéro $i sur $n..." + [Environment]::NewLine)

        try {
            # Use Invoke-WebRequest to fetch the image and save it locally
            Invoke-WebRequest -Uri $barcodeImageUrl -OutFile $barcodeImagePath
            $statusRichTextBox.AppendText("Code barre créé : $barcodeImagePath" + [Environment]::NewLine)
        }
        catch {
            $statusRichTextBox.AppendText("`rEchec de génération du code suivant : $barcodeContent. Erreur : $($_.Exception.Message)" + [Environment]::NewLine)
        }
        

    }

    $statusRichTextBox.AppendText("`rFin de la génération des codes !" + [Environment]::NewLine)

    
    # Create an instance of Excel application
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false  # Optional: set to $false to hide Excel window

    # Add a new workbook
    $workbook = $excelApp.Workbooks.Add()
    $worksheet = $workbook.Sheets.Item(1)

    # Set column widths for better display
    $worksheet.Columns.Item(1).ColumnWidth = 25  # Adjust width for file names
    $worksheet.Columns.Item(2).ColumnWidth = 30  # Adjust width for images

    # Add headers to columns
    $worksheet.Cells.Item(1, 1).Value = "File Name"
    $worksheet.Cells.Item(1, 2).Value = "Image"

    # Get all image files (e.g., .jpg, .png, .jpeg)
    $imageFiles = Get-ChildItem -Path $folderPath -File -Include *.jpg, *.png, *.jpeg

    # Start inserting images
    $row = 2  # Starting row for data (row 1 has headers)
    foreach ($imageFile in $imageFiles) {
        # Insert the file name in column 1
        $worksheet.Cells.Item($row, 1).Value = $imageFile.Name

        # Insert the image into column 2 using AddPicture method
        $imagePath = $imageFile.FullName
        $image = $Worksheet.Pictures().Insert($imagePath)

        $image.Top = $worksheet.Cells.Item($row, 2).Top
        $image.Left = $worksheet.Cells.Item($row, 2).Left

        # Increment row for the next image
        $row++
    }


    # Save the workbook
    $excelFilePath = "$folderPath\vide_de_lignes.xlsx"
    $workbook.SaveAs($excelFilePath)

    # Close the workbook and Excel application
    $workbook.Close()
    $excelApp.Quit()
        
    $statusRichTextBox.AppendText("Document créé : $excelFilePath" + [Environment]::NewLine)



})










# Show the form
$form.ShowDialog()
