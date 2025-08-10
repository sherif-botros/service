param(
    [Parameter(Mandatory=$true)]
    [string]$InputFilePath,
    
    [string]$OutputFilePath = "",
    
    [int]$MaxWordsPerSlide = 50
)

# Validate input file path
if (-not (Test-Path $InputFilePath)) {
    Write-Error "Input file path '$InputFilePath' does not exist."
    exit 1
}

try {
    # Load PowerPoint COM object
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = [Microsoft.Office.Interop.PowerPoint.PpWindowState]::ppWindowMinimized
    
    # Add a new presentation
    $presentation = $powerPoint.Presentations.Add()
    
    # Read text from input file with UTF-8 encoding
    $text = Get-Content -Path $InputFilePath -Encoding UTF8 -Raw
    
    # Split text into lines
    $lines = $text -split "\n"
    
    # Identify and group verses
    $verses = @()
    $currentVerse = ""
    
    foreach ($line in $lines) {
        # Check if line starts with a number (indicating a new verse)
        if ($line -match "^\d+") {
            # If we have a current verse, add it to the array
            if ($currentVerse -ne "") {
                $verses += $currentVerse.Trim()
            }
            # Start a new verse
            $currentVerse = $line
        } else {
            # Continue building the current verse
            $currentVerse += "`n" + $line
        }
    }
    
    # Add the last verse if it exists
    if ($currentVerse -ne "") {
        $verses += $currentVerse.Trim()
    }
    
    # Create slides and add verses
    $slideIndex = 1
    $verseIndex = 0
    
    while ($verseIndex -lt $verses.Length) {
        # Create a blank slide
        $slide = $presentation.Slides.Add($slideIndex, [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutBlank)
        
        # Add verses to the slide until we reach the word limit
        $slideContent = ""
        $currentWordCount = 0
        
        while ($verseIndex -lt $verses.Length) {
            $verse = $verses[$verseIndex]
            $verseWordCount = ($verse -split "\s+").Length
            
            # Check if adding this verse would exceed the word limit
            if ($slideContent -eq "" -or ($currentWordCount + $verseWordCount) -le $MaxWordsPerSlide) {
                if ($slideContent -ne "") {
                    $slideContent += "`n`n"
                }
                $slideContent += $verse
                $currentWordCount += $verseWordCount
                $verseIndex++
            } else {
                # If adding this verse would exceed the limit, break and create a new slide
                break
            }
        }
        
        # Add a text box that fills the slide (with small margins)
        # Standard slide dimensions are approximately 960x540 points
        $textBox = $slide.Shapes.AddTextbox(1, 20, 20, 920, 500)
        $textBox.TextFrame.TextRange.Text = $slideContent
        
        # Format text with font size 36
        $textBox.TextFrame.TextRange.Font.Size = 36
        $textBox.TextFrame.TextRange.Font.Name = "Calibri"
        
        # Enable auto-sizing to fit text
        $textBox.TextFrame.AutoSize = 1  # ppAutoSizeShapeToFitText
        
        $slideIndex++
    }
    
    # Set output path
    if ($OutputFilePath -eq "") {
        # Get file name without extension from input file
        $inputFileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFilePath)
        $OutputFilePath = [Environment]::GetFolderPath("Desktop") + "\$inputFileName.pptx"
    }
    
    # Save the presentation
    $presentation.SaveAs($OutputFilePath)
    $presentation.Close()
    $powerPoint.Quit()
    
    Write-Host "Presentation saved to $OutputFilePath"
} catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    
    # Clean up
    if ($presentation -ne $null) {
        $presentation.Close()
    }
    if ($powerPoint -ne $null) {
        $powerPoint.Quit()
    }
    
    exit 1
}
