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
$null = $tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$null = $tableLayout.RowCount = 6
$null = $tableLayout.ColumnCount = 2
$null = $tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$null = $tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 45)))
$null = $tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 60)))


# Make the row containing the statusRichTextBox flexible
$null = $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # Add AutoSize for all rows initially
$null = $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$null = $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$null = $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$null = $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  # Row for statusRichTextBox takes the remaining space
$null = $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))


$form.Controls.Add($tableLayout)

# Name Label
$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = "Numéro de machine :"
$nameLabel.Anchor = 'Left, Right'
$nameLabel.Margin = 10
$null = $tableLayout.Controls.Add($nameLabel, 0, 0) 

# Name RichTextBox (replacing TextBox with RichTextBox)
$nameRichTextBox = New-Object System.Windows.Forms.RichTextBox
$nameRichTextBox.Anchor = 'Top, Left, Right'
$nameRichTextBox.Width = 200
$nameRichTextBox.Height = 25
$nameRichTextBox.Margin = 10
$nameRichTextBox.BorderStyle = 'FixedSingle'
$null = $tableLayout.Controls.Add($nameRichTextBox, 1, 0)

# Number Label
$numberLabel = New-Object System.Windows.Forms.Label
$numberLabel.Text = "Nombre de code :"
$numberLabel.Anchor = 'Left, Right'
$numberLabel.Margin = 10
$null = $tableLayout.Controls.Add($numberLabel, 0, 1)

# Number RichTextBox
$numberRichTextBox = New-Object System.Windows.Forms.RichTextBox
$numberRichTextBox.Anchor = 'Top, Left, Right'
$numberRichTextBox.Width = 200
$numberRichTextBox.Height = 25
$numberRichTextBox.Margin = 10
$numberRichTextBox.BorderStyle = 'FixedSingle'
$null = $tableLayout.Controls.Add($numberRichTextBox, 1, 1)

# Folder Path Label
$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = "Dossier d'enregistrement :"
$folderLabel.Anchor = 'Left, Right'
$folderLabel.Margin = 10
$null = $tableLayout.Controls.Add($folderLabel, 0, 2)

# Folder Path RichTextBox
$folderRichTextBox = New-Object System.Windows.Forms.RichTextBox
$folderRichTextBox.Anchor = 'Top, Left, Right'
$folderRichTextBox.Width = 200
$folderRichTextBox.Height = 25
$folderRichTextBox.Margin = 10
$folderRichTextBox.BorderStyle = 'FixedSingle'
$null = $tableLayout.Controls.Add($folderRichTextBox, 1, 2)

# Browse Button to select folder
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Parcourir"
$browseButton.Anchor = 'Top, Right'
$browseButton.Width = 80
$null = $tableLayout.Controls.Add($browseButton, 1, 3)

# Status RichTextBox (for showing messages)
$statusRichTextBox = New-Object System.Windows.Forms.RichTextBox
$statusRichTextBox.Anchor = 'Top, Left, Right, Bottom'
$statusRichTextBox.Width = 350
$statusRichTextBox.Height = 100
$statusRichTextBox.Multiline = $true
$statusRichTextBox.ReadOnly = $true
$statusRichTextBox.BorderStyle = 'FixedSingle'
$null = $tableLayout.Controls.Add($statusRichTextBox, 0, 4)
$null = $tableLayout.SetColumnSpan($statusRichTextBox, 2)  # Span across 2 columns


# Generate Button
$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "Générer les codes"
$generateButton.Anchor = 'Left, Right,Bottom'
$generateButton.Width = 150
$null = $tableLayout.Controls.Add($generateButton, 0, 5)
$null = $tableLayout.SetColumnSpan($generateButton, 2)  # Span across 2 columns

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

    # Retry logic configuration
    $maxRetries = 3
    $retryDelaySeconds = 2  # Delay in seconds between retries

    for ($i = 1; $i -le $n; $i++) {
        # Construct barcode content and file name
        if ($i -lt 10) {  # Add a zero to print 01 instead of 1
            $barcodeContent = "$prefixe$name" + "0" + "$i"
            $tempName = "$name" + "0" + "$i"
            $barcodeImagePath = Join-Path -Path $folderPath -ChildPath "$tempName.png"
        } else {
            $barcodeContent = "$prefixe$name$i"
            $barcodeImagePath = Join-Path -Path $folderPath -ChildPath "$name$i.png"
        }

        # URL encode the content to ensure it's valid for the barcode API
        $encodedBarcodeContent = [uri]::EscapeDataString($barcodeContent)

        # Use TEC-IT barcode API
        $barcodeImageUrl = "https://barcode.tec-it.com/barcode.ashx?data=$encodedBarcodeContent&code=Code128&dpi=96&dataseparator="

        # Log progress
        $statusRichTextBox.AppendText("Génération du code numéro $i sur $n..." + [Environment]::NewLine)

        $retryCount = 0
        $success = $false

        while ($retryCount -lt $maxRetries -and !$success) {
            try {
                # Use Invoke-WebRequest to fetch the image and save it locally
                Invoke-WebRequest -Uri $barcodeImageUrl -OutFile $barcodeImagePath
                $statusRichTextBox.AppendText("Code barre créé : $barcodeImagePath" + [Environment]::NewLine)

                # If successful, exit the loop
                $success = $true
            }
            catch {
                $retryCount++
                $statusRichTextBox.AppendText("`rEchec de génération du code suivant : $barcodeContent. Erreur : $($_.Exception.Message). Tentative $retryCount sur $maxRetries..." + [Environment]::NewLine)

                if ($retryCount -lt $maxRetries) {
                    # Wait before retrying
                    Start-Sleep -Seconds $retryDelaySeconds
                }
            }
        }

        # If not successful after retries, log an error
        if (!$success) {
            $statusRichTextBox.AppendText("`rEchec de génération après $maxRetries tentatives pour : $barcodeContent" + [Environment]::NewLine)
        }

        # Add a delay to avoid rate-limiting issues
        Start-Sleep -Seconds 1
    }


    $statusRichTextBox.AppendText("`rFin de la génération des codes !" + [Environment]::NewLine)

    # Create an instance of Word application
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false  # Set to $false if you don't want Word to be visible

    # Add a new document
    $document = $wordApp.Documents.Add()

    # Add a table with two columns: one for images (left and right)
    $table = $document.Tables.Add($document.Range(0, 0), 1, 2)  # Add 1 row and 2 columns initially

    # Get all image files (e.g., .jpg, .png, .jpeg)
    $imageFiles = Get-ChildItem -Path $folderPath -File | Where-Object { $_.Extension -match "\.(jpg|png|jpeg)$" }

    # Start inserting images into Word document
    $row = 1  # Starting row for data
    $leftColumn = $true  # Flag to alternate between left and right columns
    foreach ($imageFile in $imageFiles) {

        # Insert the image in the left or right column based on the flag
        $imagePath = $imageFile.FullName
        if ($leftColumn) {
            # Insert the image in the first column (left)
            $image = $table.Cell($row, 1).Range.InlineShapes.AddPicture($imagePath)
        } else {
            # Insert the image in the second column (right)
            $image = $table.Cell($row, 2).Range.InlineShapes.AddPicture($imagePath)
        }

        # Alternate the flag for the next image

        if ($leftColumn){
            $leftColumn = $false
        } else {
            $leftColumn = $true
            $table.Rows.Add()
            $row++
        }
    }

    # Save the document
    $wordFilePath = "$folderPath\vide_de_lignes.docx"
    $document.SaveAs([ref]$wordFilePath)

    # Close the Word document and Word application
    $document.Close()
    $wordApp.Quit()

    $statusRichTextBox.AppendText("Document créé : $wordFilePath" + [Environment]::NewLine)

})

# Show the form
$form.ShowDialog()
