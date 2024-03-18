<?php

$titleDocument = 'Deliverables List';
$shipmentNumber = 'XXXXX';
$directionLine1 = 'Genmab B.V.';

$directionLine1 = 'Uppsalalaan 15';
$directionLine2 = '3584 CT Utrecht';
$directionLine3 = 'P.O Box 85199';
$directionLine4 = '3508 AD Utrecht';
$directionLine5 = 'The Netherlands';
$directionLine6 = 'Tel. +31 (0) 30 2 123 123';
$directionLine7 = 'www.genmab.com';
$directionLine8 = 'KvK 30169902';

$contact1 = 'Name: John Doe';
$contact2 = 'john@mail.com';
$origin = 'Genmab B.V. Utrecht, The Netherlands';

$dataTable1 = [
    [
        'description' => 'Ab/Ag code',
        'cellId' => 'Cell ID',
        'shortCode' => 'Short code',
        'batchSampleCode' => 'Batch/ sample code',
        'concMgMl' => 'Conc. mg/mL',
        'kindOfContainer' => 'Kind of Container',
        'numberOfContainers' => '# of containers',
        'volumePerContainer' => 'Volume (µL) per container',
        'totalShippedAmountUl' => 'Total shipped amount',
        'totalShippedAmountMg' => 'Total shipped amount',
        'shipmentTemperature' => 'Shipment temperature',
        'storageTemperature' => 'Storage temperature',
        'formulationBuffer' => 'Formulation buffer',
        'expiryDate' => 'Expiry date',
        'extinctionCoefficient' => '12'
    ],
    [
        'description' => 'Ab/Ag code',
        'cellId' => 'Cell ID',
        'shortCode' => 'Short code',
        'batchSampleCode' => 'Batch/ sample code',
        'concMgMl' => 'Conc. mg/mL',
        'kindOfContainer' => 'Kind of Container',
        'numberOfContainers' => '# of containers',
        'volumePerContainer' => 'Volume (µL) per container',
        'totalShippedAmountUl' => 'Total shipped amount',
        'totalShippedAmountMg' => 'Total shipped amount',
        'shipmentTemperature' => 'Shipment temperature',
        'storageTemperature' => 'Storage temperature',
        'formulationBuffer' => 'Formulation buffer',
        'expiryDate' => 'Expiry date',
        'extinctionCoefficient' => '12'
    ],
    [
        'description' => 'Ab/Ag code',
        'cellId' => 'Cell ID',
        'shortCode' => 'Short code',
        'batchSampleCode' => 'Batch/ sample code',
        'concMgMl' => 'Conc. mg/mL',
        'kindOfContainer' => 'Kind of Container',
        'numberOfContainers' => '# of containers',
        'volumePerContainer' => 'Volume (µL) per container',
        'totalShippedAmountUl' => 'Total shipped amount',
        'totalShippedAmountMg' => 'Total shipped amount',
        'shipmentTemperature' => 'Shipment temperature',
        'storageTemperature' => 'Storage temperature',
        'formulationBuffer' => 'Formulation buffer',
        'expiryDate' => 'Expiry date',
        'extinctionCoefficient' => '12'
    ],
    [
        'description' => 'Ab/Ag code',
        'cellId' => 'Cell ID',
        'shortCode' => 'Short code',
        'batchSampleCode' => 'Batch/ sample code',
        'concMgMl' => 'Conc. mg/mL',
        'kindOfContainer' => 'Kind of Container',
        'numberOfContainers' => '# of containers',
        'volumePerContainer' => 'Volume (µL) per container',
        'totalShippedAmountUl' => 'Total shipped amount',
        'totalShippedAmountMg' => 'Total shipped amount',
        'shipmentTemperature' => 'Shipment temperature',
        'storageTemperature' => 'Storage temperature',
        'formulationBuffer' => 'Formulation buffer',
        'expiryDate' => 'Expiry date',
        'extinctionCoefficient' => '12'
    ],
];

$dataTable2 = [
    [
        'description' => 'Ab/Ag code',
        'supplier' => 'Supplier',
        'shortCode' => 'Short code',
        'catNumber' => 'Cat. number',
        'lotNumber' => 'Lot number',
        'kindOfContainer' => 'Kind of Container',
        'numberOfContainers' => '# of containers',
        'amountPerContainer' => 'Amount per container',
        'unit' => 'Unit',
        'species' => 'Species',
        'genus' => 'Genus',
        'shipmentTemperature' => 'Shipment temperature',
        'storageTemperature' => 'Storage temperature',
        'expiryDate' => 'Expiry date',
    ],
    [
        'description' => 'Ab/Ag code',
        'supplier' => 'Supplier',
        'shortCode' => 'Short code',
        'catNumber' => 'Cat. number',
        'lotNumber' => 'Lot number',
        'kindOfContainer' => 'Kind of Container',
        'numberOfContainers' => '# of containers',
        'amountPerContainer' => 'Amount per container',
        'unit' => 'Unit',
        'species' => 'Species',
        'genus' => 'Genus',
        'shipmentTemperature' => 'Shipment temperature',
        'storageTemperature' => 'Storage temperature',
        'expiryDate' => 'Expiry date',
    ],
    [
        'description' => 'Ab/Ag code',
        'supplier' => 'Supplier',
        'shortCode' => 'Short code',
        'catNumber' => 'Cat. number',
        'lotNumber' => 'Lot number',
        'kindOfContainer' => 'Kind of Container',
        'numberOfContainers' => '# of containers',
        'amountPerContainer' => 'Amount per container',
        'unit' => 'Unit',
        'species' => 'Species',
        'genus' => 'Genus',
        'shipmentTemperature' => 'Shipment temperature',
        'storageTemperature' => 'Storage temperature',
        'expiryDate' => 'Expiry date',
    ],
];
//Set the name of the document
$filename = 'DeliverablesList.xlsx';
//load autoloader
require_once 'vendor/autoload.php';

//Create a document excel with the class PhpSpreadsheet
$spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

//Set the name of the document
$spreadsheet->getProperties()
    ->setCreator('ProcessMaker')
    ->setLastModifiedBy('ProcessMaker')
    ->setTitle('Deliverables List')
    ->setSubject('Deliverables List')
    ->setDescription('Deliverables List')
    ->setKeywords('Deliverables List')
    ->setCategory('Deliverables List');

//Set the all document font family verdana
$spreadsheet->getDefaultStyle()->getFont()->setName('Verdana');
//Set size default font 10
$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
//Set Fit to page
$spreadsheet->getActiveSheet()->getPageSetup()->setFitToPage(true);
//Hide gridlines
$spreadsheet->getActiveSheet()->setShowGridlines(false);
//Set Size of the paper
$spreadsheet->getActiveSheet()->getPageSetup()->setPaperSize(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);
//Set Orientation of the paper
$spreadsheet->getActiveSheet()->getPageSetup()->setOrientation(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
//Set the margins of the paper
$spreadsheet->getActiveSheet()->getPageMargins()->setTop(0.8);
$spreadsheet->getActiveSheet()->getPageMargins()->setRight(0.1);
$spreadsheet->getActiveSheet()->getPageMargins()->setLeft(0.1);
$spreadsheet->getActiveSheet()->getPageMargins()->setBottom(0.8);


//Set the width of the columns
$sheet->getColumnDimension('A')->setWidth(7);
$sheet->getColumnDimension('B')->setWidth(28);
$sheet->getColumnDimension('C')->setWidth(34);
$sheet->getColumnDimension('D')->setWidth(12);
$sheet->getColumnDimension('E')->setWidth(25);
$sheet->getColumnDimension('F')->setWidth(15);
$sheet->getColumnDimension('G')->setWidth(15);
$sheet->getColumnDimension('H')->setWidth(10);
$sheet->getColumnDimension('I')->setWidth(18);
$sheet->getColumnDimension('J')->setWidth(7);
$sheet->getColumnDimension('K')->setWidth(10);
$sheet->getColumnDimension('L')->setWidth(25);
$sheet->getColumnDimension('M')->setWidth(15);
$sheet->getColumnDimension('N')->setWidth(22);
$sheet->getColumnDimension('O')->setWidth(13);
$sheet->getColumnDimension('P')->setWidth(15);
$sheet->getColumnDimension('Q')->setWidth(12);

$sheet->getRowDimension('19')->setRowHeight(30);


//Set the title of the document
$sheet->setTitle('Deliverables List');

$sheet->setCellValue('A8', $shipmentNumber);
$sheet->setCellValue('A9', $titleDocument);
// Set the style of the title
$sheet->getStyle('A9')->getFont()->setBold(true);
$sheet->getStyle('A9')->getFont()->setSize(18);
$sheet->getStyle('A9')->getFont()->getColor()->setARGB('FF008080');
$sheet->getStyle('A9')->getFont()->setName('Verdana');

$sheet->setCellValue('A14', 'If more information is needed, please contact :');
$sheet->getStyle('A14')->getFont()->setBold(true);
$sheet->setCellValue('E14', $contact1);
$sheet->setCellValue('E15', $contact2);

//Insert the image in the sheet position M1 from $logoGenmabBase64
$logoGenmabBase64 = 'logo_genmab.jpg';
$logoGenmab = new PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
$logoGenmab->setPath($logoGenmabBase64);
$logoGenmab->setCoordinates('M1');  //set image to cell

// set image with and height 6cm x 1.5cm
$logoGenmab->setWidth(180);
$logoGenmab->setHeight(55);
$logoGenmab->setWorksheet($sheet);
// Cell combine

$sheet->setCellValue('M5', $directionLine1);
$sheet->setCellValue('M6', $directionLine2);
$sheet->setCellValue('M7', $directionLine3);
$sheet->setCellValue('M8', $directionLine4);
$sheet->setCellValue('M9', $directionLine5);
$sheet->setCellValue('M10', $directionLine6);
$sheet->setCellValue('M11', $directionLine7);
$sheet->setCellValue('M12', $directionLine8);


//Set the headers of the table 1
$sheet->setCellValue('B17', 'Origin:');
$sheet->setCellValue('C17', $origin);
$sheet->mergeCells('A17:A18');
$sheet->mergeCells('B17:B18');
$sheet->mergeCells('C17:E18');
$sheet->mergeCells('A19:A20');
$sheet->mergeCells('F17:Q18');

$sheet->setCellValue('B19', 'Description (Ab/Ag code)');
$sheet->setCellValue('C19', 'Cell ID');
$sheet->setCellValue('D19', 'Short code');
$sheet->setCellValue('E19', 'Batch/ sample code');
$sheet->setCellValue('F19', 'Conc. mg/mL');
$sheet->setCellValue('G19', 'Kind of Container');
$sheet->setCellValue('H19', '# of containers');
$sheet->setCellValue('I19', 'Volume (µL) per container');
$sheet->setCellValue('J19', 'Total shipped amount');
$sheet->setCellValue('J20', 'µL');
$sheet->setCellValue('K20', 'mg');
$sheet->setCellValue('L19', 'Shipment temperature');
$sheet->setCellValue('M19', 'Storage temperature');
$sheet->setCellValue('N19', 'Formulation buffer');
$sheet->setCellValue('O19', 'Expiry date');
$sheet->setCellValue('Q19', 'Extinction coefficient Lg-1cm-1');

$sheet->mergeCells('B19:B20');
$sheet->mergeCells('C19:C20');
$sheet->mergeCells('D19:D20');
$sheet->mergeCells('E19:E20');
$sheet->mergeCells('F19:F20');
$sheet->mergeCells('G19:G20');
$sheet->mergeCells('H19:H20');
$sheet->mergeCells('I19:I20');
$sheet->mergeCells('J19:K19');
$sheet->mergeCells('L19:L20');
$sheet->mergeCells('M19:M20');
$sheet->mergeCells('N19:N20');
$sheet->mergeCells('O19:O20');
$sheet->mergeCells('Q19:Q20');

//Adjust the text content
$sheet->getStyle('B17:Q20')->getAlignment()->setWrapText(true);
$sheet->getStyle('B17:Q20')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

$rowNumberStartTableBody1 = 21;
$rowNumber = $rowNumberStartTableBody1;
$i = 0;
foreach ($dataTable1 as $row) {
    //Set the data of the table 1
    $i++;
    $sheet->setCellValue('A' . $rowNumber, $i);
    $sheet->setCellValue('B' . $rowNumber, $row['description']);
    $sheet->setCellValue('C' . $rowNumber, $row['cellId']);
    $sheet->setCellValue('D' . $rowNumber, $row['shortCode']);
    $sheet->setCellValue('E' . $rowNumber, $row['batchSampleCode']);
    $sheet->setCellValue('F' . $rowNumber, $row['concMgMl']);
    $sheet->setCellValue('G' . $rowNumber, $row['kindOfContainer']);
    $sheet->setCellValue('H' . $rowNumber, $row['numberOfContainers']);
    $sheet->setCellValue('I' . $rowNumber, $row['volumePerContainer']);
    $sheet->setCellValue('J' . $rowNumber, $row['totalShippedAmountUl']);
    $sheet->setCellValue('K' . $rowNumber, $row['totalShippedAmountMg']);
    $sheet->setCellValue('L' . $rowNumber, $row['shipmentTemperature']);
    $sheet->setCellValue('M' . $rowNumber, $row['storageTemperature']);
    $sheet->setCellValue('N' . $rowNumber, $row['formulationBuffer']);
    $sheet->setCellValue('O' . $rowNumber, $row['expiryDate']);
    $sheet->setCellValue('Q' . $rowNumber, $row['extinctionCoefficient']);
    $rowNumber++;
}
$rowNumberEndTableBody1 = $rowNumber - 1;

//Set the style of the document
$styleArrayHeader = [
    'font' => [
        'bold' => true,
        'color' => ['argb' => 'FFFFFFFF'],
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
        'vertical' => 'center',
    ],
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            'color' => ['argb' => 'FFa6a6a6'],
        ],
    ],
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
        'startColor' => [
            'argb' => 'FF008080',
        ],
    ],
];
$styleArrayBody = [
    'font' => [
        'bold' => false,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            'color' => ['argb' => 'FFa6a6a6'],
        ],
    ],
];

$styleFooter = [
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
        'startColor' => [
            'argb' => 'FFb8cce4',
        ],
    ],
];



$sheet->getStyle('A17:Q20')->applyFromArray($styleArrayHeader);
$sheet->getStyle('A' . $rowNumberStartTableBody1 . ':A' . $rowNumberEndTableBody1)->applyFromArray($styleArrayHeader);
$sheet->getStyle('A' . $rowNumberStartTableBody1 . ':Q' . $rowNumberEndTableBody1)->applyFromArray($styleArrayBody);
$sheet->getStyle('A' . $rowNumberEndTableBody1 + 1 . ':Q' . $rowNumberEndTableBody1 + 1)->applyFromArray($styleFooter);

$rowNumberEndTableBody1 = $rowNumber - 1;
//End Table 1


$rowNumberStartTableHeader2 = $rowNumberEndTableBody1 + 2;
$rowNumber = $rowNumberStartTableHeader2;

//Set the headers of the table 2
$sheet->setCellValue('B' . $rowNumber, 'Origin:');
$sheet->setCellValue('C' . $rowNumber, $origin);
$sheet->mergeCells('A' . $rowNumber . ':A' . ($rowNumber + 1));
$sheet->mergeCells('B' . $rowNumber . ':B' . ($rowNumber + 1));
$sheet->mergeCells('C' . $rowNumber . ':E' . ($rowNumber + 1));
$sheet->mergeCells('F' . $rowNumber . ':J' . ($rowNumber + 1));
$sheet->setCellValue('K' . $rowNumber, 'Only for animal products (ADP)/cells');
$sheet->mergeCells('K' . $rowNumber . ':L' . $rowNumber + 1);
$sheet->mergeCells('M' . $rowNumber . ':O' . $rowNumber + 1);

$rowNumber += 2;

$sheet->setCellValue('B' . $rowNumber, 'Description');
$sheet->setCellValue('C' . $rowNumber, 'Supplier');
$sheet->setCellValue('D' . $rowNumber, 'Short code');
$sheet->setCellValue('E' . $rowNumber, 'Cat #');
$sheet->setCellValue('F' . $rowNumber, 'Lot #');
$sheet->setCellValue('G' . $rowNumber, 'Kind of Container');
$sheet->setCellValue('H' . $rowNumber, '# containers');
$sheet->setCellValue('I' . $rowNumber, 'Amount per container');
$sheet->setCellValue('J' . $rowNumber, 'Unit');
$sheet->setCellValue('K' . $rowNumber, 'Species');
$sheet->setCellValue('L' . $rowNumber, 'Genus');
$sheet->setCellValue('M' . $rowNumber, 'Shipment temperature');
$sheet->setCellValue('N' . $rowNumber, 'Storage temperature');
$sheet->setCellValue('O' . $rowNumber, 'Expiry date');

$sheet->mergeCells('A' . $rowNumber . ':A' . $rowNumber + 1);
$sheet->mergeCells('B' . $rowNumber . ':B' . $rowNumber + 1);
$sheet->mergeCells('C' . $rowNumber . ':C' . $rowNumber + 1);
$sheet->mergeCells('D' . $rowNumber . ':D' . $rowNumber + 1);
$sheet->mergeCells('E' . $rowNumber . ':E' . $rowNumber + 1);
$sheet->mergeCells('F' . $rowNumber . ':F' . $rowNumber + 1);
$sheet->mergeCells('G' . $rowNumber . ':G' . $rowNumber + 1);
$sheet->mergeCells('H' . $rowNumber . ':H' . $rowNumber + 1);
$sheet->mergeCells('I' . $rowNumber . ':I' . $rowNumber + 1);
$sheet->mergeCells('J' . $rowNumber . ':J' . $rowNumber + 1);
$sheet->mergeCells('K' . $rowNumber . ':K' . $rowNumber + 1);
$sheet->mergeCells('L' . $rowNumber . ':L' . $rowNumber + 1);
$sheet->mergeCells('M' . $rowNumber . ':M' . $rowNumber + 1);
$sheet->mergeCells('N' . $rowNumber . ':N' . $rowNumber + 1);
$sheet->mergeCells('O' . $rowNumber . ':O' . $rowNumber + 1);
$sheet->getRowDimension($rowNumber)->setRowHeight(30);

$rowNumberEndTableHeader2 = $rowNumber;

//Adjust the text content
$sheet->getStyle('B' . $rowNumberStartTableHeader2 . ':O' . $rowNumberEndTableHeader2)->getAlignment()->setWrapText(true);
$sheet->getStyle('A' . ($rowNumberEndTableHeader2 - 2) . ':O' . ($rowNumberEndTableHeader2 + 1))->applyFromArray($styleArrayHeader);

$rowNumberStartTableBody2 = $rowNumber + 2;
$rowNumber = $rowNumberStartTableBody2;
$i = 0;
foreach ($dataTable2 as $row) {
    $i++;
    $sheet->setCellValue('A' . $rowNumber, $i);
    $sheet->setCellValue('B' . $rowNumber, $row['description']);
    $sheet->setCellValue('C' . $rowNumber, $row['supplier']);
    $sheet->setCellValue('D' . $rowNumber, $row['shortCode']);
    $sheet->setCellValue('E' . $rowNumber, $row['catNumber']);
    $sheet->setCellValue('F' . $rowNumber, $row['lotNumber']);
    $sheet->setCellValue('G' . $rowNumber, $row['kindOfContainer']);
    $sheet->setCellValue('H' . $rowNumber, $row['numberOfContainers']);
    $sheet->setCellValue('I' . $rowNumber, $row['amountPerContainer']);
    $sheet->setCellValue('J' . $rowNumber, $row['unit']);
    $sheet->setCellValue('K' . $rowNumber, $row['species']);
    $sheet->setCellValue('L' . $rowNumber, $row['genus']);
    $sheet->setCellValue('M' . $rowNumber, $row['shipmentTemperature']);
    $sheet->setCellValue('N' . $rowNumber, $row['storageTemperature']);
    $sheet->setCellValue('O' . $rowNumber, $row['expiryDate']);
    $rowNumber++;
}
$rowNumberEndTableBody2 = $rowNumber - 1;

$sheet->getStyle('A' . $rowNumberStartTableBody2 . ':A' . $rowNumberEndTableBody2)->applyFromArray($styleArrayHeader);
$sheet->getStyle('A' . $rowNumberStartTableBody2 . ':O' . $rowNumberEndTableBody2)->applyFromArray($styleArrayBody);
$sheet->getStyle('A' . $rowNumberEndTableBody2 + 1 . ':O' . $rowNumberEndTableBody2 + 1)->applyFromArray($styleFooter);

$sheet->setCellValue('A' . $rowNumberEndTableBody2 + 2, 'Remarks');
$sheet->mergeCells('A' . $rowNumberEndTableBody2 + 2 . ':O' . $rowNumberEndTableBody2 + 3);
$sheet->getStyle('A' . $rowNumberEndTableBody2 + 2 . ':O' . $rowNumberEndTableBody2 + 3)->applyFromArray($styleArrayHeader);

$rowNumberEndPage = $rowNumberEndTableBody2 + 3;

//Set the print area
$spreadsheet->getActiveSheet()->getPageSetup()->setPrintArea('A1:Q' . $rowNumberEndPage);

//Set the header of the document
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="' . $filename . '"');
header('Cache-Control: max-age=0');


//Save on the path
// $writer = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
// $writer->save('DeliverablesList.xlsx');

//Save the document
$writer = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');
exit;
