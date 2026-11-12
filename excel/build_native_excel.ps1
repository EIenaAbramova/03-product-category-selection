$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$file = Join-Path $root 'outputs\03_решение_Excel.xlsx'
$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$app.ScreenUpdating = $false
$book = $app.Workbooks.Add()
while ($book.Worksheets.Count -gt 1) { $book.Worksheets.Item($book.Worksheets.Count).Delete() }
$main=$book.Worksheets.Item(1); $main.Name='Результат'
$settings=$book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count));$settings.Name='Настройки'
$cat=$book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count));$cat.Name='Категории'
$examples=$book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count));$examples.Name='Товары'
$audit=$book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count));$audit.Name='Проверки'
function Put($s,$addr,$value) {
    if($value -is [ValueType]){$s.Range($addr).Formula='='+$value.ToString([System.Globalization.CultureInfo]::InvariantCulture)}
    else{$s.Range($addr).Value2=[string]$value}
}
function Header($range) { $range.Interior.Color=0x4D3320;$range.Font.Color=0xFFFFFF;$range.Font.Bold=$true;$range.WrapText=$true;$range.RowHeight=35;$range.VerticalAlignment=-4108 }
function Title($s,$text) { Put $s 'A2' $text; $s.Range('A2').Font.Size=19;$s.Range('A2').Font.Bold=$true; $s.Rows.Item(2).RowHeight=32 }
foreach($s in @($main,$settings,$cat,$examples,$audit)) {
    $s.Cells.Font.Name='Aptos';$s.Cells.Font.Size=11;$s.Cells.Font.Color=0x4D3320
    $s.Rows.RowHeight=22;$s.Activate();$app.ActiveWindow.DisplayGridlines=$false
}
Title $main 'Выбор товарных категорий — решение в Excel'
Put $main 'A3' 'Excel + Power Query | данные за 03–30 июля 2024 года'
Title $settings 'Настройки расчёта в Excel'
Put $settings 'A4' 'Папка с четырьмя исходными CSV.GZ'
Put $settings 'B5' (Join-Path $root 'data')
$null=$book.Names.Add('DataFolder',"='Настройки'!`$B`$5")
Put $settings 'A7' 'ДРР';Put $settings 'B7' 'Доля';Put $settings 'C7' 'Товаров'
Put $settings 'A8' 0.03;Put $settings 'B8' 0.10;Put $settings 'C8' 30
$settingTable=$settings.ListObjects.Add(1,$settings.Range('A7:C8'),$null,1);$settingTable.Name='Settings';$settingTable.TableStyle='TableStyleMedium2'
$settings.Range('A8:B8').NumberFormatLocal='0,0%';$settings.Range('A8:C8').Interior.Color=0xCCF2FF
$settings.Range('B5').Interior.Color=0xCCF2FF
Put $settings 'A10' 'ДРР — максимум; доля миллионников и число товаров — минимум.'
Put $settings 'A12' 'После изменения порогов таблица категорий и итоговый список пересчитываются автоматически.'
Put $settings 'A13' 'Для обновления исходных данных и примеров товаров: Данные → Обновить всё.'
Put $settings 'A15' 'Порог миллионника: сумма заказов строго больше 1 000 000 ₽ за 28 дней.'
Put $settings 'A16' 'Пороги 3% / 10% / 30 — аналитические допущения, а не условия работодателя.'
Put $settings 'A18' 'Путь укажите к папке data из скачанного репозитория. Распаковывать CSV.GZ не требуется.'
Put $settings 'A20' 'Показатель выручки: сумма заказов. Возвраты и себестоимость в источнике отсутствуют.'
Put $settings 'A21' 'Расходы восстановлены как сумма заказов × ДРР карточки.'
$settings.Columns.Item('A').ColumnWidth=30;$settings.Columns.Item('B').ColumnWidth=60;$settings.Columns.Item('C').ColumnWidth=22
foreach($queryName in @('Products','Categories','Quality','Examples')) {
    $m=Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot ($queryName+'.pq'))
    $null=$book.Queries.Add($queryName,$m)
}
function LoadQuery($name,$sheet,$tableName) {
    Write-Host ('REFRESH '+$name)
    $conn='OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location='+$name+';Extended Properties=""'
    $lo=$sheet.ListObjects.Add(0,$conn,$null,1,$sheet.Range('A5'))
    $lo.Name=$tableName
    $qt=$lo.QueryTable;$qt.CommandType=2;$qt.CommandText='SELECT * FROM ['+$name+']';$qt.BackgroundQuery=$false
    $qt.AdjustColumnWidth=$false;$qt.PreserveFormatting=$true
    $null=$qt.Refresh($false)
    $lo.TableStyle='TableStyleMedium2'
    Write-Host ($name+' rows: '+$lo.ListRows.Count)
    return $lo
}
Title $cat 'Все категории — расчёт Power Query'
Put $cat 'A3' 'Ключ категории — полный путь четырёх уровней. Все строки участвуют в расчёте.'
$categoryTable=LoadQuery 'Categories' $cat 'tblCategories'
if($categoryTable.ListRows.Count -ne 4588){throw 'Не совпало число категорий'}
$selection=$categoryTable.ListColumns.Add();$selection.Name='Отбор'
$selection.DataBodyRange.Formula='=IF(AND([@Товаров]>=Настройки!$C$8,[@[ДРР взвешенная]]<=Настройки!$A$8,[@[Доля миллионников]]>=Настройки!$B$8),"Да","Нет")'
$position=$categoryTable.ListColumns.Add();$position.Name='Позиция'
$position.DataBodyRange.FormulaR1C1='=IF(RC[-1]="Да",COUNTIFS(R6C17:RC17,"Да"),"")'
Header $categoryTable.HeaderRowRange
$cat.Columns.Item('A:C').ColumnWidth=28;$cat.Columns.Item('D').ColumnWidth=47;$cat.Columns.Item('E:R').ColumnWidth=18
$categoryTable.DataBodyRange.RowHeight=45;$cat.Range('A6:D4593').WrapText=$true
foreach($col in @('Доля миллионников','ДРР взвешенная','Доля лидера','ДРР средняя')){$categoryTable.ListColumns.Item($col).DataBodyRange.NumberFormatLocal='0,00%'}
foreach($col in @('Сумма заказов','Расходы расчетные','Медианная цена','Лидер заказы')){$categoryTable.ListColumns.Item($col).DataBodyRange.NumberFormatLocal='# ##0'}
Title $audit 'Проверка исходных данных'
Put $audit 'A3' 'Независимые контроли по всем карточкам; ДРР больше 100% сохранена.'
$qualityTable=LoadQuery 'Quality' $audit 'tblQuality'
Header $qualityTable.HeaderRowRange
$audit.Columns.Item('A').ColumnWidth=44;$audit.Columns.Item('B:D').ColumnWidth=26
$qualityTable.ListColumns.Item(2).DataBodyRange.NumberFormatLocal='# ##0,00'
Put $audit 'D5' 'Проверка';Put $audit 'E5' 'Результат'
Put $audit 'D6' 'Количество карточек';$audit.Range('E6').Formula='=IF(SUM(tblCategories[Товаров])=B6,"Совпало","Расхождение")'
Put $audit 'D7' 'Количество миллионников';$audit.Range('E7').Formula='=IF(SUM(tblCategories[Более миллиона])=B9,"Совпало","Расхождение")'
Put $audit 'D8' 'Сумма заказов';$audit.Range('E8').Formula='=IF(ABS(SUM(tblCategories[Сумма заказов])-B10)<0.01,"Совпало","Расхождение")'
Put $audit 'D10' 'Проверено в Microsoft Excel';Put $audit 'E10' '12.11.2026'
$audit.Columns.Item('D').ColumnWidth=32;$audit.Columns.Item('E').ColumnWidth=23;Header $audit.Range('D5:E5')
Title $examples 'Примеры товаров — решение в Excel'
Put $examples 'A3' 'До трёх карточек на отобранную категорию. Сумма заказов >1 млн ₽; ДРР ≤порога.'
$exampleTable=LoadQuery 'Examples' $examples 'tblExamples'
Header $exampleTable.HeaderRowRange
$examples.Columns.Item('A:C').ColumnWidth=25;$examples.Columns.Item('D').ColumnWidth=40;$examples.Columns.Item('E').ColumnWidth=65;$examples.Columns.Item('F').ColumnWidth=22;$examples.Columns.Item('G:L').ColumnWidth=20
$exampleTable.DataBodyRange.RowHeight=60;$exampleTable.DataBodyRange.WrapText=$true
$exampleTable.ListColumns.Item('ДРР').DataBodyRange.NumberFormatLocal='0,00%'
foreach($col in @('Заказы, ₽','Цена, ₽')){$exampleTable.ListColumns.Item($col).DataBodyRange.NumberFormatLocal='# ##0'}
# Main sheet contains native formulas that rebuild the shortlist when thresholds change.
Put $main 'A5' 'Товаров изучено';$main.Range('B5').Formula='=SUM(tblCategories[Товаров])'
Put $main 'D5' 'Категорий всего';$main.Range('E5').Formula='=ROWS(tblCategories[Категория])'
Put $main 'G5' 'В отборе';$main.Range('H5').Formula='=COUNTIFS(tblCategories[Отбор],"Да")'
$main.Range('B5').NumberFormatLocal='# ##0';$main.Range('B5:H5').Font.Bold=$true
Put $main 'A7' 'ДРР не выше';$main.Range('B7').Formula='=Настройки!A8';$main.Range('B7').NumberFormatLocal='0,0%'
Put $main 'D7' 'Миллионников от';$main.Range('E7').Formula='=Настройки!B8';$main.Range('E7').NumberFormatLocal='0,0%'
Put $main 'G7' 'Товаров от';$main.Range('H7').Formula='=Настройки!C8'
$headers=@('Категория','Товаров','>1 млн, шт.','Доля >1 млн','ДРР взвеш.','Доля лидера','Продавцов','Концентрация')
for($i=0;$i -lt 8;$i++){$main.Cells.Item(10,$i+1).Value2=$headers[$i]};Header $main.Range('A10:H10')
for($r=11;$r -le 40;$r++) {
    $rank=$r-10
    $main.Range('J'+$r).Formula='=IF('+ $rank +'<=COUNTIFS(tblCategories[Отбор],"Да"),MATCH('+ $rank +',tblCategories[Позиция],0),"")'
    $refs=@('Категория','Товаров','Более миллиона','Доля миллионников','ДРР взвешенная','Доля лидера','Продавцов')
    for($i=0;$i -lt 7;$i++){$col=[char](65+$i);$main.Range("$col$r").Formula='=IF($J'+$r+'="","",INDEX(tblCategories['+$refs[$i]+'],$J'+$r+'))'}
    $main.Range('H'+$r).Formula='=IF($J'+$r+'="","",IF(F'+$r+'>=0.9,"Лидер ≥90%","Лидер <90%"))'
}
$main.Range('D11:F40').NumberFormatLocal='0,00%';$main.Range('A11:A40').WrapText=$true;$main.Range('A11:H40').RowHeight=30
$main.Columns.Item('A').ColumnWidth=46;$main.Columns.Item('B:C').ColumnWidth=12;$main.Columns.Item('D:G').ColumnWidth=15;$main.Columns.Item('H').ColumnWidth=21;$main.Columns.Item('J').Hidden=$true
Put $main 'A43' 'Показаны первые 30 категорий текущего отбора. Полный перечень — на листе «Категории».'
Put $main 'A45' 'Выводы для закупочной проверки'
$main.Range('A45').Font.Bold=$true
Put $main 'A46' 'Смартфоны Samsung и iPhone 15: сочетание низкой ДРР и высокой доли миллионников.'
Put $main 'A47' 'Фитнес-браслеты: более низкая медианная цена; проверка поставки небольшими партиями.'
Put $main 'A48' 'Детские смеси и встраиваемые стиральные машины: высокая концентрация у лидера.'
Put $main 'A50' 'Данные за июль 2024 года. Перед закупкой нужны свежие данные и расчёт маржи.'
Put $main 'A51' 'Низкая ДРР не подтверждает прибыльность. Сумма заказов не учитывает возвраты.'
# Hide unused shortlist rows for the default view, retaining formulas for 30 positions.
# Rows remain visible: they are needed when settings expand the shortlist.
$main.Range('A11:H20').Interior.Color=0xFFF5EA
$main.Tab.Color=0xFF5B00
foreach($s in @($main,$settings,$cat,$examples,$audit)){
    $s.Activate();$app.ActiveWindow.Zoom=85;$app.ActiveWindow.DisplayGridlines=$false
    $s.PageSetup.Orientation=2;$s.PageSetup.Zoom=$false;$s.PageSetup.FitToPagesWide=1;$s.PageSetup.FitToPagesTall=$false
    $s.Range('A1').Select()
}
$cat.Activate();$cat.Range('E6').Select();$app.ActiveWindow.FreezePanes=$true
$examples.Activate();$examples.Range('F6').Select();$app.ActiveWindow.FreezePanes=$true
$app.CalculateFullRebuild()
if($main.Range('H5').Value2 -ne 10){throw 'Базовый отбор должен дать 10 категорий'}
Put $settings 'A8' 0.01;$app.CalculateFullRebuild()
if($main.Range('H5').Value2 -ne 4){throw 'Строгий порог должен дать 4 категории'}
Put $settings 'A8' 0.03;$app.CalculateFullRebuild()
if($main.Range('H5').Value2 -ne 10){throw 'Не восстановился базовый отбор'}
foreach($r in 6..8){if($audit.Range('E'+$r).Value2 -ne 'Совпало'){throw 'Контроль не сошёлся'}}
$cat.Range('D2').Value2='Все категории — расчёт Power Query';$cat.Range('D2').Font.Size=19;$cat.Range('D2').Font.Bold=$true
$cat.Range('D3').Value2='Полный путь категории сохранён в столбцах A–D. Кнопка «+» раскрывает уровни.'
$cat.Columns.Item('A:C').Group()|Out-Null;$cat.Columns.Item('A:C').Hidden=$true
$cat.Columns.Item('I:J').Group()|Out-Null;$cat.Columns.Item('I:J').Hidden=$true
$cat.Columns.Item('M').Group()|Out-Null;$cat.Columns.Item('M').Hidden=$true;$cat.Columns.Item('R').Hidden=$true
$cat.Columns.Item('D').ColumnWidth=46;$cat.Columns.Item('E:F').ColumnWidth=12;$cat.Columns.Item('G:H').ColumnWidth=20
$cat.Columns.Item('K').ColumnWidth=12;$cat.Columns.Item('L').ColumnWidth=16;$cat.Columns.Item('N:P').ColumnWidth=16;$cat.Columns.Item('Q').ColumnWidth=10
$cat.Range('A6:R4593').RowHeight=40;$cat.Range('A5:R4593').VerticalAlignment=-4108;$cat.Rows.Item(5).RowHeight=42
$audit.Range('B6:B9').NumberFormatLocal='# ##0';$audit.Range('B11:B13').NumberFormatLocal='# ##0'
$examples.Columns.Item('A:C').Group()|Out-Null;$examples.Columns.Item('A:C').Hidden=$true
$examples.Range('D2').Value2='Примеры товаров — решение в Excel';$examples.Range('D2').Font.Size=19;$examples.Range('D2').Font.Bold=$true
$examples.Range('D3').Value2='До трёх карточек на выбранную категорию. Полный путь раскрывается кнопкой +.'
$examples.UsedRange.VerticalAlignment=-4108
$audit.Activate();$app.ActiveWindow.Zoom=120
$main.Activate();$main.Range('A1').Select();$app.ActiveWindow.Zoom=100
$book.SaveAs($file,51)
$settings.Range('B5').Formula='=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"..\data"'
$app.CalculateFullRebuild()
$book.Save()
Write-Output ('SAVED '+$file)
Write-Output ('QUERIES '+$book.Queries.Count+'; CATEGORY_ROWS '+$categoryTable.ListRows.Count+'; EXAMPLES '+$exampleTable.ListRows.Count)
$app.ScreenUpdating=$true;$app.Visible=$true
# Leave only this newly authored workbook open for screenshots and manual review.
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
