Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "QR Code Generator"
$form.Width = 450
$form.Height = 350
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(400, 350)

# TableLayoutPanel for flexible resizing
$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayout.RowCount = 6
$tableLayout.ColumnCount = 2
$tableLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 60)))
$form.Controls.Add($tableLayout)

# Name Label
$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = "Enter the name:"
$nameLabel.Anchor = 'Left'
$tableLayout.Controls.Add($nameLabel, 0, 0)

# Name RichTextBox (replacing TextBox with RichTextBox)
$nameRichTextBox = New-Object System.Windows.Forms.RichTextBox
$nameRichTextBox.Anchor = 'Top, Left, Right'
$nameRichTextBox.Width = 200
$nameRichTextBox.Height = 30
$nameRichTextBox.BorderStyle = 'FixedSingle'
$tableLayout.Controls.Add($nameRichTextBox, 1, 0)

# Number Label
$numberLabel = New-Object System.Windows.Forms.Label
$numberLabel.Text = "Enter the number of QR codes:"
$numberLabel.Anchor = 'Left'
$tableLayout.Controls.Add($numberLabel, 0, 1)

# Number RichTextBox
$numberRichTextBox = New-Object System.Windows.Forms.RichTextBox
$numberRichTextBox.Anchor = 'Top, Left, Right'
$numberRichTextBox.Width = 200
$numberRichTextBox.Height = 30
$numberRichTextBox.BorderStyle = 'FixedSingle'
$tableLayout.Controls.Add($numberRichTextBox, 1, 1)

# Folder Path Label
$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = "Select folder to save:"
$folderLabel.Anchor = 'Left'
$tableLayout.Controls.Add($folderLabel, 0, 2)

# Folder Path RichTextBox
$folderRichTextBox = New-Object System.Windows.Forms.RichTextBox
$folderRichTextBox.Anchor = 'Top, Left, Right'
$folderRichTextBox.Width = 200
$folderRichTextBox.Height = 30
$folderRichTextBox.BorderStyle = 'FixedSingle'
$tableLayout.Controls.Add($folderRichTextBox, 1, 2)

# Browse Button to select folder
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse"
$browseButton.Anchor = 'Top, Right'
$browseButton.Width = 80
$tableLayout.Controls.Add($browseButton, 1, 3)

# Status RichTextBox (for showing messages)
$statusRichTextBox = New-Object System.Windows.Forms.RichTextBox
$statusRichTextBox.Anchor = 'Top, Left, Right'
$statusRichTextBox.Width = 350
$statusRichTextBox.Height = 100
$statusRichTextBox.Multiline = $true
$statusRichTextBox.ReadOnly = $true
$statusRichTextBox.BorderStyle = 'FixedSingle'
$tableLayout.Controls.Add($statusRichTextBox, 0, 4)
$tableLayout.SetColumnSpan($statusRichTextBox, 2)  # Span across 2 columns

# Generate Button
$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = "Generate QR Codes"
$generateButton.Anchor = 'Top, Left, Right'
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
    $prefixe = "Faf@|Q7G8TH7bGxg#e!6Jqj_"

    # Validate inputs
    if ([string]::IsNullOrEmpty($name) -or [string]::IsNullOrEmpty($n) -or [string]::IsNullOrEmpty($folderPath)) {
        $statusRichTextBox.AppendText("Please fill in all fields." + [Environment]::NewLine)
        return
    }

    # Validate that the number is a valid positive integer
    if (-not ($n -gt 0)) {
        $statusRichTextBox.AppendText("Please enter a valid positive number for the number of barcodes." + [Environment]::NewLine)
        return
    }

    # Ensure folder exists
    if (-not (Test-Path $folderPath)) {
        $statusRichTextBox.AppendText("The specified folder does not exist. Creating the folder..." + [Environment]::NewLine)
        New-Item -ItemType Directory -Path $folderPath
    }

    # Start Barcode Generation
    $statusRichTextBox.AppendText("Generating Barcodes..." + [Environment]::NewLine)

    for ($i = 1; $i -le $n; $i++) {
        $barcodeContent = "$name$i"  # Content for the barcode (name + number)
        
        # Using Barcode API to generate a Code 128 barcode (PNG format)
        $barcodeImageUrl = "https://www.barcodesinc.com/generator/image.php?code=$barcodeContent&type=C128B&width=300&height=150&xres=1&font=3"

        # Download the barcode image from the API
        $barcodeImagePath = Join-Path -Path $folderPath -ChildPath "$name$i.png"

        # Log progress
        $statusRichTextBox.AppendText("Generating Barcode $i of $n..." + [Environment]::NewLine)

        try {
            # Use Invoke-WebRequest to fetch the image and save it locally
            Invoke-WebRequest -Uri $barcodeImageUrl -OutFile $barcodeImagePath
            $statusRichTextBox.AppendText("Barcode saved: $barcodeImagePath" + [Environment]::NewLine)
        }
        catch {
            $statusRichTextBox.AppendText("Failed to generate or save barcode for: $barcodeContent. Error: $($_.Exception.Message)" + [Environment]::NewLine)
        }
    }

    $statusRichTextBox.AppendText("Barcode generation complete!" + [Environment]::NewLine)
})





# Show the form
$form.ShowDialog()
