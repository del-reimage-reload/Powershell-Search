$shcount = $null# Search Powerpoint for a text string
Add-type -AssemblyName office
Add-type -AssemblyName microsoft.office.interop.powerpoint

$path = "\\192.168.1.1\Share" #Place holder will need to add the ability to have user input
$dest_path = "\\192.168.1.1\Results" #Place holder will need to add the ability to have user input
$MatchString = "cisco" #Place holder will need to add the ability to have user input
$pptSlides = Get-Childitem -Path $path -Recurse -Include *.pptx,*.ppt #Need to test .ppt capabilities also $pptSlides returns a list of files the .Name give the file name the .FullName give the FQFN and the .Directory give the direcotry

# Create instance of the PowerPoint.Application COM object
$ppt = New-Object -ComObject PowerPoint.Application

# iterate over each PPT in $pptSlides
foreach($pptSlide in $pptSlides)
{

#write-host $pptSlide
$presentation = $ppt.Presentations.open($pptSlide)

# Max slides is determined for slide one
$slcount = $ppt.ActivePresentation.Slides.Count
#Write-Host "Total Slides = $slcount"
$e = $slcount
$i = 0
while($i -ne $e)
{$i++

$test = $ppt.ActivePresentation.Slides($i).Shapes.count
$testo = $ppt.ActivePresentation.Slides($i).Shapes
$x = 0
while($x -ne $test){
$x++
foreach($d in $testo){
        if($d.HasTextFrame)
            {
            #$name = $ppt.ActivePresentation.Slides($i).Shapes($x).Name
            $string = $ppt.ActivePresentation.Slides($i).Shapes($x).TextFrame.TextRange | select Text #pulls the text out of each shape in each slide of a PPT
            if($string -Match $MatchString)
            {
            #$string  #prints matched string if you want to see it
            Copy-Item $pptSlide.FullName -Destination $dest_path 
            
            }
            }
}}}}