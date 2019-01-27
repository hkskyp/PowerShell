
function week2day([int]$year, [int]$month, [int]$week)
{
    $a = [DateTime]("$month/1/$year")
    $offset = @(3, 2, 1, 0, 6, 5, 4, 3)

    $b = $a.AddDays($offset[$a.DayOfWeek])

    $c = $b.AddDays(($week-1) * 7)

    # @(월요일, 금요일)
    return @($c.AddDays($offset[3]-2), $c.AddDays($offset[3]+2))
}

#2019년 2월 1주
$result = week2day 2019 2 1
Write-Output $result

function OpenWord()
{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $True
    $doc = $word.documents.Add()
    #$doc = $word.documents.Open($filename)

    $wdBulletGallery	= 1	#Bulleted list.
    $wdNumberGallery	= 2	#Numbered list.
    $wdOutlineNumberGallery	= 3	#Outline numbered list.
    $gallery = $word.ListGalleries[$wdOutlineNumberGallery]
    $listTemplate = $gallery.ListTemplates[1]
    #$style = $word.Styles["Heading 1"]
    #$style.LinkToListTemplate($listTemplate, 1);

    $selection = $word.selection


    #$selection.Range.ListFormat.ApplyNumberDefault()
    $selection.Range.ListFormat.ApplyListTemplate($listTemplate)
    $selection.Range.ListFormat.ListLevelNumber = 1
    $selection.TypeText("aa")
    $selection.TypeParagraph()
    $selection.TypeText("bb")
    $selection.TypeParagraph()

    $selection.Range.ListFormat.ListLevelNumber = 2
    $selection.TypeText("aaa")
    $selection.TypeParagraph()
    $selection.TypeText("bbb")
    $selection.TypeParagraph()
    $selection.TypeText("ccc")
    $selection.TypeParagraph()

    $selection.Range.ListFormat.ListLevelNumber = 1
    $selection.TypeText("dd")
    $selection.TypeParagraph()


    $selection.Range.ListFormat.ApplyListTemplate($listTemplate, $False)
    $selection.Range.ListFormat.ListLevelNumber = 1
    $selection.TypeText("aa")
    $selection.TypeParagraph()
    $selection.TypeText("bb")
    $selection.TypeParagraph()

    $selection.Range.ListFormat.RemoveNumbers()
    $selection.TypeText("aa")
    $selection.TypeParagraph()
    $selection.TypeText("bb")
    $selection.TypeParagraph()

    $wdColumnBreak	= 8	#Column break at the insertion point.
    $wdLineBreak	= 6	#Line break.
    $wdLineBreakClearLeft	= 9	#Line break.
    $wdLineBreakClearRight	= 10	#Line break.
    $wdPageBreak	= 7	#Page break at the insertion point.
    $wdSectionBreakContinuous	= 3	#New section without a corresponding page break.
    $wdSectionBreakEvenPage	= 4	#Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.
    $wdSectionBreakNextPage	= 2	#Section break on next page.
    $wdSectionBreakOddPage	= 5	#Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.
    $wdTextWrappingBreak	= 11	#Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.
    $selection.InsertBreak( $wdPageBreak)

    $selection.Range.ListFormat.RemoveNumbers()
    $selection.TypeText("a1")
    $selection.TypeParagraph()
    $selection.TypeText("b2")
    $selection.TypeParagraph()

    return
    $word.Quit()
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable word
}

function OpenExcel([string]$filename)
{
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $False
    $wb = $xl.Workbooks.Open($filename)
    $ws = $wb.Sheets.Item(1)

    $cell = $ws.Cells.Item(1, 2)
    $cell = $ws.Range("A2")
    if ($cell.MergeCells)
    {
        Write-Output "MergeCells"
        Write-Output $cell.MergeArea.Cells(1, 1).Text
    }
    else
    {
        Write-Output $cell.Text
    }

    Write-Output ([Convert]::ToString($ws.Cells.Item(3, 1).Interior.Color, 16))
    Write-Output ($ws.Cells.Item(3, 2).Interior.Color -eq 0xC47244)
    
    $xl.Quit()
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$xl)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable xl
}

#OpenExcel "D:\주간작성.xlsx"
OpenWord