<?php
//load autoloader
require_once 'vendor/autoload.php';


//Create a document excel with the class PhpSpreadsheet
$spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

//Set the name of the document
$spreadsheet->getProperties()
    ->setCreator('Maarten Balliauw')
    ->setLastModifiedBy('Maarten Balliauw')
    ->setTitle('Office 2007 XLSX Test Document')
    ->setSubject('Office 2007 XLSX Test Document')
    ->setDescription('Test document for Office 2007 XLSX, generated using PHP classes.')
    ->setKeywords('office 2007 openxml php')
    ->setCategory('Test result file');

//Set the title of the document
$sheet->setTitle('Deliverables List');

//Set the headers of the document
$sheet->setCellValue('A1', 'ID');
$sheet->setCellValue('B1', 'Name');
$sheet->setCellValue('C1', 'Description');
$sheet->setCellValue('D1', 'Status');
$sheet->setCellValue('E1', 'Start Date');
$sheet->setCellValue('F1', 'End Date');
$sheet->setCellValue('G1', 'Project ID');

// //Set the data of the document
// $deliverables = Deliverable::getAll();
// $row = 2;

// foreach ($deliverables as $deliverable) {
//     $sheet->setCellValue('A' . $row, $deliverable->getId());
//     $sheet->setCellValue('B' . $row, $deliverable->getName());
//     $sheet->setCellValue('C' . $row, $deliverable->getDescription());
//     $sheet->setCellValue('D' . $row, $deliverable->getStatus());
//     $sheet->setCellValue('E' . $row, $deliverable->getStartDate());
//     $sheet->setCellValue('F' . $row, $deliverable->getEndDate());
//     $sheet->setCellValue('G' . $row, $deliverable->getProjectId());
//     $row++;
// }

// //Set the headers of the document
// $sheet->setCellValue('A' . $row, 'ID');
// $sheet->setCellValue('B' . $row, 'Name');
// $sheet->setCellValue('C' . $row, 'Description');
// $sheet->setCellValue('D' . $row, 'Status');
// $sheet->setCellValue('E' . $row, 'Start Date');
// $sheet->setCellValue('F' . $row, 'End Date');
// $sheet->setCellValue('G' . $row, 'Project ID');

//Set the style of the document
$styleArray = [
    'font' => [
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
    'borders' => [
        'top' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
        'rotation' => 90,
        'startColor' => [
            'argb' => 'FFA0A0A0',
        ],
        'endColor' => [
            'argb' => 'FFFFFFFF',
        ],
    ],
];

$sheet->getStyle('A1:G1')->applyFromArray($styleArray);

//Set the width of the columns
$sheet->getColumnDimension('A')->setWidth(5);
$sheet->getColumnDimension('B')->setWidth(20);
$sheet->getColumnDimension('C')->setWidth(30);
$sheet->getColumnDimension('D')->setWidth(10);
$sheet->getColumnDimension('E')->setWidth(15);
$sheet->getColumnDimension('F')->setWidth(15);
$sheet->getColumnDimension('G')->setWidth(10);

//Set the name of the document
$filename = 'DeliverablesList.xlsx';

// //Set the header of the document
// header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
// header('Content-Disposition: attachment;filename="' . $filename . '"');
// header('Cache-Control: max-age=0');


//Save on the path
$writer = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('DeliverablesList.xlsx');

//Save the document
$writer = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');
exit;
