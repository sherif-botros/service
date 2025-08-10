# Load PowerPoint COM object
$powerPoint = New-Object -ComObject PowerPoint.Application
$powerPoint.Visible = [Microsoft.Office.Interop.PowerPoint.PpWindowState]::ppWindowMinimized

# Add a new presentation
$presentation = $powerPoint.Presentations.Add()

# Set the text (Acts Chapter 8, for example)
$text = @"
[Your full Acts chapter 8 verses here...]
"@

# Split text into chunks for each slide
$verses = $text -split "\n"
$maxLinesPerSlide = 5

# Create slides and add verses
$slideIndex = 1
for ($i = 0; $i -lt $verses.Length; $i += $maxLinesPerSlide) {
    $slide = $presentation.Slides.Add($slideIndex, [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutText)
    $shape = $slide.Shapes.Item(1)
    $shape.TextFrame.TextRange.Text = ($verses[$i..($i + $maxLinesPerSlide - 1)] -join "`n")
    $shape.TextFrame.TextRange.Font.Size = 36
    $shape.TextFrame.TextRange.Font.Name = "Calibri"
    $slideIndex++
}

# Save the presentation
$presentation.SaveAs("C:\Path\To\Save\Acts_Chapter8.pptx")
$presentation.Close()
$powerPoint.Quit()
