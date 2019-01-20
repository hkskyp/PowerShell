$input = "D:\하이퍼링크테스트.pptx"
$output = "D:\하이퍼링크테스트-링킹.pptx"

#MsoHyperlinkType enumeration
#https://docs.microsoft.com/en-us/office/vba/api/office.msohyperlinktype
$msoHyperlinkInlineShape	= 2	#Hyperlink applies to an inline shape. Used only with Microsoft Word.
$msoHyperlinkRange	= 0	#Hyperlink applies to a Range object.
$msoHyperlinkShape	= 1	#Hyperlink applies to a Shape object.

#PpMouseActivation Enumeration
#https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmouseactivation
$ppMouseClick	= 1	#Mouse click
$ppMouseOver	= 2	#Mouse over

#PpActionType Enumeration
#https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
$ppActionEndShow	= 6	#Slide show ends.
$ppActionFirstSlide	= 3	#Returns to the first slide.
$ppActionHyperlink	= 7	#Hyperlink.
$ppActionLastSlide	= 4	#Moves to the last slide.
$ppActionLastSlideViewed	= 5	#Moves to the last slide viewed.
$ppActionMixed	= -2	#Performs a mixed action.
$ppActionNamedSlideShow	= 10	#Runs the slideshow.
$ppActionNextSlide	= 1	#Moves to the next slide.
$ppActionNone	= 0	#No action is performed.
$ppActionOLEVerb	= 11	#OLE Verb.
$ppActionPlay	= 12	#Begins the slideshow.
$ppActionPreviousSlide	= 2	#Moves to the previous slide.
$ppActionRunMacro	= 8	#Runs a macro.
$ppActionRunProgram	= 9	#Runs a program.


Add-Type -AssemblyName Office
$objPPT = New-Object -ComObject powerpoint.Application
#open but hide
$Doc = $objPPT.Presentations.Open($input, $Null,$Null,[Microsoft.Office.Core.MsoTriState]::msoFalse)
$Slides = $Doc.Slides

$moveTo = @{}   #내용이 있는 슬라이드
$moveFrom = @{} #초기 슬라이드

#슬라이드 인덱스 조회
Write-Output "슬라이드 인덱스 조회"
foreach ($Slide in $Slides)
{
    $out = "SlideID: " + $Slide.SlideID.ToString() + ", SlideNumber: " + $Slide.SlideNumber + ","
    Write-Output $out

    #되돌아 올 슬라이드 주소를 등록한다.
    foreach ($shape in $Slide.Shapes)
    {
        #To: 슬라이드로 이동 후 $moveFrom 슬라이드로 돌아온다.
        if ($shape.AlternativeText -match "To:")
        {
            $moveFrom[$shape.AlternativeText.Split(":")[1]] =  $Slide.SlideID.ToString() + "," + $Slide.SlideNumber + ","
            continue;
        }
        #From: 슬라이드로 이동 후 $moveTo 슬라이드로 돌아온다.
        if ($shape.AlternativeText -match "From:")
        {
            $moveTo[$shape.AlternativeText.Split(":")[1]] = $Slide.SlideID.ToString() + "," + $Slide.SlideNumber + ","
            continue;
        }
    }
}

#하이퍼링크 세팅
Write-Output "하이퍼링크 세팅"
foreach ($Slide in $Slides)
{
    $out = "SlideID: " + $Slide.SlideID.ToString() + ", SlideNumber: " + $Slide.SlideNumber + ","
    Write-Output $out
    foreach ($shape in $Slide.Shapes)
    {
        if ($shape.AlternativeText -match "To:")
        {
            $shape.ActionSettings($ppMouseClick).Action = $ppActionHyperlink
            $shape.ActionSettings($ppMouseClick).Hyperlink.SubAddress = $moveTo[$shape.AlternativeText.Split(":")[1]]
            $out = "`t" + $shape.AlternativeText
            Write-Output $out
            continue;
        }
        if ($shape.AlternativeText -match "From:")
        {
            $shape.ActionSettings($ppMouseClick).Action = $ppActionHyperlink
            $shape.ActionSettings($ppMouseClick).Hyperlink.SubAddress = $moveFrom[$shape.AlternativeText.Split(":")[1]]
            $out = "`t" + $shape.AlternativeText
            Write-Output $out
            continue;
        }
    }
}


$Doc.SaveAs($output)
$Doc.Close()
$objPPT.Quit()
$objPPT = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
