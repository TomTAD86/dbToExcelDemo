<?php

require '../../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Get data from JSON
$result = file_get_contents("../util/employeesData.json");
$decode = json_decode("$result");
$employeeData = $decode->employees;

$spreadsheet = new Spreadsheet();

$columnHeaders = array('First Name', 'Last Name', 'Position', 'Department', 'Salary', 'Sales This Year');
$columnArray = array('B', 'C', 'D', 'E', 'F', 'G');

// Parse data to create 3d array from employeeData
$parsedEmployeeData = array();
foreach ($employeeData as $employee) {
  array_push($parsedEmployeeData, array($employee->firstName, $employee->lastName, $employee->position, $employee->department, $employee->salary, $employee->salesThisYear));
};

// Get cell references for TOTAL headers and cells for formulae and highest row
$totalHeaderCellReference = "B" . count($employeeData) + 5;
$totalSalaryCellReference = "F" . count($employeeData) + 5;
$totalSalesCellReference = "G" . count($employeeData) + 5;
$highestRow = count($employeeData) + 4;


// Fill in basic table from data
$spreadsheet->getActiveSheet()
  // Create Column Titles
  ->fromArray(
    $columnHeaders,
    NULL,
    'B4'
  )
  // Fill in data from 3d array
  ->fromArray(
    $parsedEmployeeData,
    NULL,
    'B5'
  )
  // Insert table title
  ->setCellValue('B2', 'Sales and Retentions')
  // Insert header and formulae for TOTALs
  ->setCellValue($totalHeaderCellReference, 'TOTAL')
  ->setCellValue($totalSalaryCellReference, "=SUM(F5:F${highestRow})")
  ->setCellValue($totalSalesCellReference, "=SUM(G5:G${highestRow})");

// Formatting

// White background
$whiteBackgroundArray = [
  'fill' => [
    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
    'color' => [
      'argb' => 'FFFFFFFF'
    ]
  ]
];
$spreadsheet->getActiveSheet()->getStyle('A1:Z100')->applyFromArray($whiteBackgroundArray);


// Format Table Title
// Create Array For Title Styles
$titleStyleArray = [
  'font' => [
    'bold' => true,
    'size' => 16,
    'color' => [
      'argb' => 'FFFFFFFF'
    ],
  ],
  'fill' => [
    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
    'color' => [
      'argb' => 'FF0e5373'
    ]
  ],
  'borders' => [
    'bottom' => [
      'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
      'color' => [
        'argb' => 'FFFFFFFF'
      ],
    ],
  ],
  'alignment' => [
    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
  ],
];

// Merge Title Cells
$spreadsheet->getActiveSheet()->mergeCells('B2:G3');

// Apply Style Array To Title
$spreadsheet->getActiveSheet()->getStyle('B2:G3')->applyFromArray($titleStyleArray);
//

// Format Column Headers
// Create Array For Column Headers
$columnHeadersStyleArray = [
  'font' => [
    'bold' => true,
    'color' => [
      'argb' => 'FFFFFFFF'
    ]
  ],
  'alignment' => [
    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
  ],
  'fill' => [
    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
    'color' => [
      'argb' => 'FF0e5373'
    ]
  ],
];

// Set column widths to autosize based on data
foreach ($columnArray as $column) {
  $spreadsheet->getActiveSheet()->getColumnDimension("${column}")->setWidth(24);
};

$writer = new Xlsx($spreadsheet);
$writer->save('output.xlsx');
