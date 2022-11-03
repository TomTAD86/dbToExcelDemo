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
$totalRowReference = count($employeeData) + 5;
$totalHeaderCellReference = "B" . $totalRowReference;
$totalSalaryCellReference = "F" . $totalRowReference;
$totalSalesCellReference = "G" . $totalRowReference;
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
  ->setCellValue($totalHeaderCellReference, 'TOTALS')
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
// Create Style Array For Title Styles
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
// Create Style Array For Column Headers
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

// Apply Style Array To Column Headers
$spreadsheet->getActiveSheet()->getStyle('B4:G4')->applyFromArray($columnHeadersStyleArray);

// Format Table Data
// Create Style Array For Table Data
$tableDataStyleArray = [
  'font' => [
    'color' => [
      'argb' => 'FF0e5373'
    ],
  ],
  'borders' => [
    'inside' => [
      'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
      'color' => [
        'argb' => 'FF0e5373'
      ],
    ],
  ],
  'fill' => [
    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
    'color' => [
      'argb' => 'FFdce6f1'
    ]
  ]
];

// Apply Style Array To Table Data
$spreadsheet->getActiveSheet()->getStyle("B5:G${highestRow}")->applyFromArray($tableDataStyleArray);

// Format TOTALS
// Create Style Array for TOTALS header
$totalHeaderStyleArray = [
  'font' => [
    'bold' => true,
    'size' => 12,
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
  'alignment' => [
    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
  ],
];

// Merge TOTALS Header
$endOfTable = $highestRow + 2;
$totalMergedCellEnd = 'E' . $endOfTable;
$spreadsheet->getActiveSheet()->mergeCells("${totalHeaderCellReference}:${totalMergedCellEnd}");

// Apply Style Array to TOTALS header
$spreadsheet->getActiveSheet()->getStyle("${totalHeaderCellReference}:${totalMergedCellEnd}")->applyFromArray($totalHeaderStyleArray);
//

// Format Salary And Sales Totals
// Create Style Array For Salary And Sales Totals
$salaryAndSalesTotalStyleArray = [
  'font' => [
    'bold' => true,
    'size' => 12,
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
  'alignment' => [
    'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
  ],
  'borders' => [
    'left' => [
      'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
      'color' => [
        'argb' => 'FFFFFFFF'
      ],
    ],
    'vertical' => [
      'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
      'color' => [
        'argb' => 'FFFFFFFF'
      ],
    ],
  ],
];

// Merge Salary total
$salaryTotalMergedCellEnd = 'F' . $endOfTable;
$spreadsheet->getActiveSheet()->mergeCells("${totalSalaryCellReference}:${salaryTotalMergedCellEnd}");

// Merge Sales total
$salesTotalMergedCellEnd = 'G' . $endOfTable;
$spreadsheet->getActiveSheet()->mergeCells("${totalSalesCellReference}:${salesTotalMergedCellEnd}");

// Apply Style Array to Salary And Sales Total
$spreadsheet->getActiveSheet()->getStyle("${totalSalaryCellReference}:${salesTotalMergedCellEnd}")->applyFromArray($salaryAndSalesTotalStyleArray);

// Set Sales and Salary As Currency
$spreadsheet->getActiveSheet()->getStyle("F5:{$salesTotalMergedCellEnd}")
  ->getNumberFormat()->setFormatCode('"£"#,##0.00_-');

// Conditional Formatting For Sales
// Sales Greater Than 25000 Conditional
$conditional = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
$conditional->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS);
$conditional->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_GREATERTHAN);
$conditional->addCondition(25000);
$conditional->getStyle()->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_DARKGREEN);
$conditional->getStyle()->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
$conditional->getStyle()->getFill()->getEndColor()->setARGB('FFc4d79b');

$conditionalStyles = $spreadsheet->getActiveSheet()->getStyle("G5:G{$highestRow}")->getConditionalStyles();
$conditionalStyles[] = $conditional;

$spreadsheet->getActiveSheet()->getStyle("G5:G{$highestRow}")->setConditionalStyles($conditionalStyles);

// Sales Less Than 25000 Conditional
$conditional2 = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
$conditional2->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS);
$conditional2->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_LESSTHAN);
$conditional2->addCondition(25000);
$conditional2->getStyle()->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_DARKRED);
$conditional2->getStyle()->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
$conditional2->getStyle()->getFill()->getEndColor()->setARGB('FFe6b8b7');

$conditionalStyles = $spreadsheet->getActiveSheet()->getStyle("G5:G{$highestRow}")->getConditionalStyles();
$conditionalStyles[] = $conditional2;

$spreadsheet->getActiveSheet()->getStyle("G5:G{$highestRow}")->setConditionalStyles($conditionalStyles);

// Set column widths
foreach ($columnArray as $column) {
  $spreadsheet->getActiveSheet()->getColumnDimension("${column}")->setWidth(24);
};

// Create Spreadseet
$writer = new Xlsx($spreadsheet);
$writer->save('output.xlsx');
