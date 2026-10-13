param(
    [Parameter(Position=0, ValueFromRemainingArguments=$true)]
    [string[]]$Paths
)

$ErrorActionPreference = 'Stop'
$ToolRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ToolRoot 'mapping.json'
$InputDir = Join-Path $ToolRoot 'input'
$OutputDir = Join-Path $ToolRoot 'output'
New-Item -ItemType Directory -Force -Path $InputDir, $OutputDir | Out-Null

if (-not (Test-Path -LiteralPath $ConfigPath)) { throw "Missing mapping.json beside the converter." }
$Config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$InventoryLocation = [string]$Config.inventory.location
if ([string]::IsNullOrWhiteSpace($InventoryLocation) -or $InventoryLocation -eq 'SET THIS BEFORE FIRST USE') {
    throw "Inventory location is not configured. Double-click 'Set Inventory Location.cmd' first."
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName Microsoft.VisualBasic

function Get-ZipXml([string]$Path, [string]$EntryName) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry($EntryName)
        if (-not $entry) { return $null }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $zip.Dispose() }
}

function Get-XlsxRows([string]$Path) {
    $shared = @()
    $sharedXml = Get-ZipXml $Path 'xl/sharedStrings.xml'
    if ($sharedXml) {
        $ns = New-Object System.Xml.XmlNamespaceManager($sharedXml.NameTable)
        $ns.AddNamespace('s','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        foreach ($si in $sharedXml.SelectNodes('//s:si', $ns)) {
            $parts = @($si.SelectNodes('.//s:t', $ns) | ForEach-Object { $_.'#text' })
            $shared += ($parts -join '')
        }
    }
    $sheet = Get-ZipXml $Path 'xl/worksheets/sheet1.xml'
    if (-not $sheet) { throw 'The XLSX does not contain a first worksheet.' }
    $ns = New-Object System.Xml.XmlNamespaceManager($sheet.NameTable)
    $ns.AddNamespace('s','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $result = @()
    foreach ($rowNode in $sheet.SelectNodes('//s:sheetData/s:row', $ns)) {
        while ($result.Count -lt ([int]$rowNode.r - 1)) { $result += ,@() }
        $values = @{}
        foreach ($cell in $rowNode.SelectNodes('s:c', $ns)) {
            $ref = [string]$cell.r
            $letters = ([regex]::Match($ref, '^[A-Z]+')).Value
            $col = 0
            foreach ($ch in $letters.ToCharArray()) { $col = ($col * 26) + ([int][char]$ch - [int][char]'A' + 1) }
            $type = [string]$cell.t
            $value = ''
            if ($type -eq 'inlineStr') {
                $value = (@($cell.SelectNodes('.//s:t', $ns) | ForEach-Object { $_.'#text' }) -join '')
            } elseif ($cell.v) {
                $raw = [string]$cell.v
                if ($type -eq 's') { $value = $shared[[int]$raw] }
                elseif ($type -eq 'b') { $value = if ($raw -eq '1') { 'true' } else { 'false' } }
                else { $value = $raw }
            }
            $values[$col] = $value
        }
        $max = if ($values.Count) { ($values.Keys | Measure-Object -Maximum).Maximum } else { 0 }
        $arr = @()
        for ($i=1; $i -le $max; $i++) { $arr += [string]$values[$i] }
        $result += ,$arr
    }
    return ,$result
}

function Get-OdsRows([string]$Path) {
    $doc = Get-ZipXml $Path 'content.xml'
    if (-not $doc) { throw 'The ODS file has no content.xml.' }
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('table','urn:oasis:names:tc:opendocument:xmlns:table:1.0')
    $ns.AddNamespace('text','urn:oasis:names:tc:opendocument:xmlns:text:1.0')
    $table = $doc.SelectSingleNode('//table:table', $ns)
    $result = @()
    foreach ($rowNode in $table.SelectNodes('table:table-row', $ns)) {
        $rowRepeat = [string]$rowNode.GetAttribute('number-rows-repeated','urn:oasis:names:tc:opendocument:xmlns:table:1.0')
        if (-not $rowRepeat) { $rowRepeat = 1 }
        $arr = @()
        foreach ($cell in $rowNode.SelectNodes('table:table-cell', $ns)) {
            $colRepeat = [string]$cell.GetAttribute('number-columns-repeated','urn:oasis:names:tc:opendocument:xmlns:table:1.0')
            if (-not $colRepeat) { $colRepeat = 1 }
            $texts = @($cell.SelectNodes('.//text:p', $ns) | ForEach-Object { $_.InnerText })
            $value = $texts -join "`n"
            for ($i=0; $i -lt [Math]::Min([int]$colRepeat, 100); $i++) { $arr += $value }
        }
        for ($r=0; $r -lt [Math]::Min([int]$rowRepeat, 500); $r++) { $result += ,$arr }
    }
    return ,$result
}

function Get-CsvRows([string]$Path) {
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, [Text.Encoding]::UTF8, $true)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        $result = @()
        while (-not $parser.EndOfData) {
            try { $result += ,@($parser.ReadFields()) }
            catch [Microsoft.VisualBasic.FileIO.MalformedLineException] {
                throw "Invalid CSV formatting near source line $($parser.ErrorLineNumber): $($parser.ErrorLine)"
            }
        }
        return ,$result
    } finally { $parser.Dispose() }
}

function Normalize-Header([string]$Text) { return (($Text -replace '[^A-Za-z0-9]','').ToLowerInvariant()) }

function Convert-Decimal([string]$Value) {
    $clean = ($Value -replace '[^0-9.\-]','')
    $number = 0.0
    if ([double]::TryParse($clean, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return $null
}

function Normalize-Barcode([string]$Value) {
    $text = $Value.Trim()
    if ($text -match '^[0-9]+\.0$') { return $text.Substring(0, $text.Length - 2) }
    if ($text -match '^[0-9.]+[Ee][+\-]?[0-9]+$') {
        $number = [decimal]0
        if ([decimal]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
            return $number.ToString('0', [Globalization.CultureInfo]::InvariantCulture)
        }
    }
    return $text
}

function Get-Cell($Row, $Index) {
    if ($Index -ge 0 -and $Index -lt $Row.Count) { return ([string]$Row[$Index]).Trim() }
    return ''
}

function New-Handle([string]$Title, [string]$Sku) {
    $text = ($Title + '-' + $Sku).ToLowerInvariant()
    $text = [Text.NormalizationForm]::FormD | ForEach-Object { $text.Normalize($_) }
    $text = $text -replace '[^a-z0-9]+','-'
    return $text.Trim('-')
}

function Escape-Csv([object]$Value) {
    $s = if ($null -eq $Value) { '' } else { [string]$Value }
    if ($s -match '[,"\r\n]') { return '"' + ($s -replace '"','""') + '"' }
    return $s
}

function Write-CsvUtf8Lf([string]$Path, [string[]]$Headers, $Rows) {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add((@($Headers | ForEach-Object { Escape-Csv $_ }) -join ','))
    foreach ($row in $Rows) {
        $lines.Add((@($Headers | ForEach-Object { Escape-Csv $row[$_] }) -join ','))
    }
    $text = ($lines -join "`n") + "`n"
    [IO.File]::WriteAllText($Path, $text, (New-Object Text.UTF8Encoding($false)))
}

function ConvertTo-XmlText([object]$Value) {
    $s = if ($null -eq $Value) { '' } else { [string]$Value }
    return [Security.SecurityElement]::Escape($s)
}

function Get-ExcelColumnName([int]$Number) {
    $name = ''
    while ($Number -gt 0) {
        $Number--
        $name = [char](65 + ($Number % 26)) + $name
        $Number = [Math]::Floor($Number / 26)
    }
    return $name
}

function Add-ZipTextEntry($Zip, [string]$Name, [string]$Text, [bool]$NoCompression = $false) {
    $level = if ($NoCompression) { [IO.Compression.CompressionLevel]::NoCompression } else { [IO.Compression.CompressionLevel]::Optimal }
    $entry = $Zip.CreateEntry($Name, $level)
    $stream = $entry.Open()
    $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Write-XlsxTable([string]$Path, [string[]]$Headers, $Rows) {
    $allRows = New-Object System.Collections.Generic.List[object]
    $allRows.Add(@($Headers))
    foreach ($row in $Rows) { $allRows.Add(@($Headers | ForEach-Object { $row[$_] })) }
    $sheetRows = New-Object Text.StringBuilder
    for ($r=0; $r -lt $allRows.Count; $r++) {
        $rowNumber = $r + 1
        [void]$sheetRows.Append("<row r=`"$rowNumber`">")
        for ($c=0; $c -lt $Headers.Count; $c++) {
            $reference = "$(Get-ExcelColumnName ($c + 1))$rowNumber"
            $value = $allRows[$r][$c]
            if ($r -gt 0 -and $value -is [ValueType] -and $value -isnot [bool]) {
                $number = [Convert]::ToString($value, [Globalization.CultureInfo]::InvariantCulture)
                [void]$sheetRows.Append("<c r=`"$reference`" t=`"n`"><v>$number</v></c>")
            } else {
                $text = ConvertTo-XmlText $value
                [void]$sheetRows.Append("<c r=`"$reference`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$text</t></is></c>")
            }
        }
        [void]$sheetRows.Append('</row>')
    }
    $lastColumn = Get-ExcelColumnName $Headers.Count
    $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
        "<sheetData>$sheetRows</sheetData><autoFilter ref=`"A1:$lastColumn$($allRows.Count)`"/></worksheet>"
    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    $rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Inventory" sheetId="1" r:id="rId1"/></sheets></workbook>'
    $workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew)
    $zip = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Create)
    try {
        Add-ZipTextEntry $zip '[Content_Types].xml' $contentTypes
        Add-ZipTextEntry $zip '_rels/.rels' $rootRels
        Add-ZipTextEntry $zip 'xl/workbook.xml' $workbook
        Add-ZipTextEntry $zip 'xl/_rels/workbook.xml.rels' $workbookRels
        Add-ZipTextEntry $zip 'xl/worksheets/sheet1.xml' $sheetXml
    } finally { $zip.Dispose(); $stream.Dispose() }
}

function Write-OdsTable([string]$Path, [string[]]$Headers, $Rows) {
    $allRows = New-Object System.Collections.Generic.List[object]
    $allRows.Add(@($Headers))
    foreach ($row in $Rows) { $allRows.Add(@($Headers | ForEach-Object { $row[$_] })) }
    $tableRows = New-Object Text.StringBuilder
    foreach ($values in $allRows) {
        [void]$tableRows.Append('<table:table-row>')
        foreach ($value in $values) {
            if ($value -is [ValueType] -and $value -isnot [bool]) {
                $number = [Convert]::ToString($value, [Globalization.CultureInfo]::InvariantCulture)
                [void]$tableRows.Append("<table:table-cell office:value-type=`"float`" office:value=`"$number`"><text:p>$number</text:p></table:table-cell>")
            } else {
                $text = ConvertTo-XmlText $value
                [void]$tableRows.Append("<table:table-cell office:value-type=`"string`"><text:p>$text</text:p></table:table-cell>")
            }
        }
        [void]$tableRows.Append('</table:table-row>')
    }
    $content = '<?xml version="1.0" encoding="UTF-8"?>' +
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="Inventory">' +
        $tableRows.ToString() + '</table:table></office:spreadsheet></office:body></office:document-content>'
    $manifest = '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>'
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew)
    $zip = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Create)
    try {
        Add-ZipTextEntry $zip 'mimetype' 'application/vnd.oasis.opendocument.spreadsheet' $true
        Add-ZipTextEntry $zip 'content.xml' $content
        Add-ZipTextEntry $zip 'META-INF/manifest.xml' $manifest
    } finally { $zip.Dispose(); $stream.Dispose() }
}

function Convert-One([string]$Path) {
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -eq '.xlsx') { $rows = Get-XlsxRows $Path }
    elseif ($ext -eq '.ods') { $rows = Get-OdsRows $Path }
    elseif ($ext -eq '.csv') { $rows = Get-CsvRows $Path }
    else { throw "Unsupported file type: $ext" }

    $wanted = @{}
    foreach ($p in $Config.source_columns.PSObject.Properties) { $wanted[$p.Name] = Normalize-Header ([string]$p.Value) }
    $headerIndex = -1; $indexes = @{}
    for ($r=0; $r -lt $rows.Count; $r++) {
        $norm = @($rows[$r] | ForEach-Object { Normalize-Header ([string]$_) })
        if (($norm -contains $wanted.sku) -and ($norm -contains $wanted.title)) {
            $headerIndex = $r
            foreach ($name in $wanted.Keys) { $indexes[$name] = [array]::IndexOf($norm, $wanted[$name]) }
            break
        }
    }
    if ($headerIndex -lt 0) { throw 'Could not find the configured product header row.' }

    $items = @{}
    $review = New-Object System.Collections.Generic.List[object]
    $sawProduct = $false
    for ($r=$headerIndex+1; $r -lt $rows.Count; $r++) {
        $row = $rows[$r]
        $sku = Get-Cell $row $indexes.sku
        $title = Get-Cell $row $indexes.title
        if (-not $sku -and -not $title) {
            if ($sawProduct) { break }
            continue
        }
        $sawProduct = $true
        $qty = Convert-Decimal (Get-Cell $row $indexes.quantity)
        $retail = Convert-Decimal (Get-Cell $row $indexes.retail)
        $cost = Convert-Decimal (Get-Cell $row $indexes.cost)
        $barcode = Normalize-Barcode (Get-Cell $row $indexes.barcode)
        $dateText = Get-Cell $row $indexes.in_store_date
        $categoryCode = (Get-Cell $row $indexes.item_category) -replace '\.0$',''
        $country = Get-Cell $row $indexes.country
        $reasons = New-Object System.Collections.Generic.List[string]
        $blocking = $false
        if (-not $sku) { $reasons.Add('Missing SKU/code'); $blocking = $true }
        if (-not $title) { $reasons.Add('Missing title'); $blocking = $true }
        if ($null -eq $qty -or $qty -le 0 -or $qty -ne [Math]::Floor($qty)) { $reasons.Add('Quantity must be a positive whole number'); $blocking = $true }
        if ($null -eq $retail) { $reasons.Add('Missing or invalid retail price'); $blocking = $true }
        elseif ($retail -eq 0 -and $Config.validation.exclude_zero_price) { $reasons.Add('Zero retail price: confirm whether this promotional item belongs in Shopify'); $blocking = $true }
        if ($null -eq $cost) { $reasons.Add('Missing or invalid cost'); }
        if (-not $barcode) { $reasons.Add('Missing barcode'); }
        elseif ($barcode -notmatch '^\d+$') { $reasons.Add('Barcode contains non-digits'); }
        $dateNumber = Convert-Decimal $dateText
        if ($dateText -match '(?:^|[/\-])1901(?:$|[/\-])' -or $dateText -match '^1901' -or ($null -ne $dateNumber -and $dateNumber -gt 0 -and $dateNumber -lt 3653)) { $reasons.Add('Suspicious in-store date: 1901'); }

        if ($sku -and $items.ContainsKey($sku)) {
            $existing = $items[$sku]
            if ($existing.Title -eq $title -and $existing.Retail -eq $retail -and $existing.Barcode -eq $barcode) {
                $existing.Quantity += $qty
                $review.Add([pscustomobject]@{ SourceRow=$r+1; SKU=$sku; Title=$title; Severity='INFO'; ReviewReason='Duplicate SKU merged; quantity added to earlier row'; IncludedInShopifyCsv='Yes' })
                continue
            } else {
                $reasons.Add('Duplicate SKU has conflicting product details'); $blocking = $true
            }
        }

        $item = [pscustomobject]@{ SourceRow=$r+1; SKU=$sku; Title=$title; Quantity=$qty; Retail=$retail; Cost=$cost; Barcode=$barcode; Date=$dateText; CategoryCode=$categoryCode; Country=$country; Blocking=$blocking }
        if ($sku -and -not $items.ContainsKey($sku)) { $items[$sku] = $item }
        if ($reasons.Count) {
            $review.Add([pscustomobject]@{ SourceRow=$r+1; SKU=$sku; Title=$title; Severity=if($blocking){'REVIEW'}else{'WARNING'}; ReviewReason=$reasons -join '; '; IncludedInShopifyCsv=if($blocking){'No'}else{'Yes'} })
        }
    }

    $shopHeaders = @(
        'Title','URL handle','Description','Vendor','Product category','Type','Tags','Published on online store','Status','SKU','Barcode',
        'Option1 name','Option1 value','Option1 Linked To','Option2 name','Option2 value','Option2 Linked To','Option3 name','Option3 value','Option3 Linked To',
        'Price','Compare-at price','Cost per item','Charge tax','Tax code','Unit price total measure','Unit price total measure unit','Unit price base measure','Unit price base measure unit',
        'Inventory tracker','Inventory quantity','Continue selling when out of stock','Weight value (grams)','Weight unit for display','Requires shipping','Fulfillment service',
        'Product image URL','Image position','Image alt text','Variant image URL','Gift card','SEO title','SEO description','Color (product.metafields.shopify.color-pattern)',
        'Google Shopping / Google product category','Google Shopping / Gender','Google Shopping / Age group','Google Shopping / Manufacturer part number (MPN)',
        'Google Shopping / Ad group name','Google Shopping / Ads labels','Google Shopping / Condition','Google Shopping / Custom product',
        'Google Shopping / Custom label 0','Google Shopping / Custom label 1','Google Shopping / Custom label 2','Google Shopping / Custom label 3','Google Shopping / Custom label 4'
    )
    $shopRows = New-Object System.Collections.Generic.List[hashtable]
    $inventoryRows = New-Object System.Collections.Generic.List[hashtable]
    foreach ($item in ($items.Values | Sort-Object SourceRow)) {
        if ($item.Blocking) { continue }
        $type = [string]$Config.defaults.type
        $mapProp = $Config.category_map.PSObject.Properties | Where-Object Name -eq $item.CategoryCode | Select-Object -First 1
        if ($mapProp) { $type = [string]$mapProp.Value }
        $tags = @([string]$Config.defaults.tags)
        if ($Config.defaults.add_country_tag -and $item.Country) { $tags += ('Country-' + $item.Country) }
        $shopRows.Add(@{
            'Title'=$item.Title; 'URL handle'=(New-Handle $item.Title $item.SKU); 'Description'=''; 'Vendor'=[string]$Config.defaults.vendor;
            'Product category'=[string]$Config.defaults.product_category; 'Type'=$type; 'Tags'=($tags | Where-Object { $_ } | Select-Object -Unique) -join ', ';
            'Published on online store'=([string]$Config.defaults.published_on_online_store).ToUpperInvariant(); 'Status'=[string]$Config.defaults.status; 'SKU'=$item.SKU;
            'Barcode'=$item.Barcode; 'Option1 name'='Title'; 'Option1 value'='Default Title'; 'Price'=$item.Retail.ToString('0.00',[Globalization.CultureInfo]::InvariantCulture);
            'Cost per item'=if($null -eq $item.Cost){''}else{$item.Cost.ToString('0.0000',[Globalization.CultureInfo]::InvariantCulture)};
            'Charge tax'=([string]$Config.defaults.charge_tax).ToUpperInvariant(); 'Inventory tracker'=[string]$Config.defaults.inventory_tracker; 'Inventory quantity'=[int]$item.Quantity;
            'Continue selling when out of stock'=[string]$Config.defaults.continue_selling_when_out_of_stock; 'Weight value (grams)'=[string]$Config.defaults.weight_grams;
            'Weight unit for display'=[string]$Config.defaults.weight_unit; 'Requires shipping'=([string]$Config.defaults.requires_shipping).ToUpperInvariant(); 'Fulfillment service'=[string]$Config.defaults.fulfillment_service;
            'Gift card'='FALSE'
        })
        $inventoryRows.Add(@{
            'Handle'=(New-Handle $item.Title $item.SKU); 'Title'=$item.Title; 'Option1 Name'='Title'; 'Option1 Value'='Default Title';
            'Option2 Name'=''; 'Option2 Value'=''; 'Option3 Name'=''; 'Option3 Value'=''; 'SKU'=$item.SKU; 'HS Code'=''; 'COO'=$item.Country;
            'Location'=$InventoryLocation; 'Bin name'=''; 'Incoming (not editable)'=''; 'Unavailable (not editable)'='';
            'Committed (not editable)'=''; 'Available (not editable)'=''; 'On hand (current)'=''; 'On hand (new)'=[int]$item.Quantity
        })
    }

    $base = [IO.Path]::GetFileNameWithoutExtension($Path)
    $sourceType = [IO.Path]::GetExtension($Path).TrimStart('.').ToLowerInvariant()
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $shopPath = Join-Path $OutputDir ("Shopify_{0}_{1}_{2}.csv" -f $base,$sourceType,$stamp)
    $inventoryPath = Join-Path $OutputDir ("Inventory_{0}_{1}_{2}.csv" -f $base,$sourceType,$stamp)
    $inventoryXlsxPath = Join-Path $OutputDir ("Inventory_{0}_{1}_{2}.xlsx" -f $base,$sourceType,$stamp)
    $inventoryOdsPath = Join-Path $OutputDir ("Inventory_{0}_{1}_{2}.ods" -f $base,$sourceType,$stamp)
    $reviewPath = Join-Path $OutputDir ("Review_{0}_{1}_{2}.csv" -f $base,$sourceType,$stamp)
    Write-CsvUtf8Lf $shopPath $shopHeaders $shopRows
    $inventoryHeaders = @('Handle','Title','Option1 Name','Option1 Value','Option2 Name','Option2 Value','Option3 Name','Option3 Value','SKU','HS Code','COO','Location','Bin name','Incoming (not editable)','Unavailable (not editable)','Committed (not editable)','Available (not editable)','On hand (current)','On hand (new)')
    Write-CsvUtf8Lf $inventoryPath $inventoryHeaders $inventoryRows
    Write-XlsxTable $inventoryXlsxPath $inventoryHeaders $inventoryRows
    Write-OdsTable $inventoryOdsPath $inventoryHeaders $inventoryRows
    $reviewHeaders = @('SourceRow','SKU','Title','Severity','ReviewReason','IncludedInShopifyCsv')
    $reviewRows = @($review | ForEach-Object { @{SourceRow=$_.SourceRow;SKU=$_.SKU;Title=$_.Title;Severity=$_.Severity;ReviewReason=$_.ReviewReason;IncludedInShopifyCsv=$_.IncludedInShopifyCsv} })
    Write-CsvUtf8Lf $reviewPath $reviewHeaders $reviewRows
    return [pscustomobject]@{ Input=$Path; Products=$shopRows.Count; Units=(($inventoryRows | ForEach-Object {[int]$_['On hand (new)']} | Measure-Object -Sum).Sum); ReviewItems=$review.Count; ShopifyCsv=$shopPath; InventoryCsv=$inventoryPath; InventoryXlsx=$inventoryXlsxPath; InventoryOds=$inventoryOdsPath; ReviewCsv=$reviewPath }
}

$files = New-Object System.Collections.Generic.List[string]
if ($Paths -and $Paths.Count) {
    foreach ($p in $Paths) {
        if (Test-Path -LiteralPath $p -PathType Container) {
            Get-ChildItem -LiteralPath $p -File | Where-Object Extension -in '.xlsx','.ods','.csv' | ForEach-Object { $files.Add($_.FullName) }
        } elseif (Test-Path -LiteralPath $p -PathType Leaf) { $files.Add((Resolve-Path -LiteralPath $p).Path) }
    }
} else {
    Get-ChildItem -LiteralPath $InputDir -File | Where-Object Extension -in '.xlsx','.ods','.csv' | ForEach-Object { $files.Add($_.FullName) }
}

if (-not $files.Count) { throw "No .csv, .xlsx, or .ods files found. Drop a file onto 'Convert Invoices.cmd' or place it in the input folder." }
$results = @($files | ForEach-Object { Convert-One $_ })
foreach ($result in $results) {
    Write-Host "Converted: $([IO.Path]::GetFileName($result.Input))"
    Write-Host "  Shopify products: $($result.Products)"
    Write-Host "  Inventory units: $($result.Units)"
    Write-Host "  Review notes: $($result.ReviewItems)"
    Write-Host "  Shopify CSV: $($result.ShopifyCsv)"
    Write-Host "  Inventory CSV: $($result.InventoryCsv)"
    Write-Host "  Inventory Excel: $($result.InventoryXlsx)"
    Write-Host "  Inventory ODS: $($result.InventoryOds)"
    Write-Host "  Review CSV:  $($result.ReviewCsv)"
}
