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
$columnsArray = array('B', 'C', 'D', 'E', 'F', 'G');

// Parse data to create 3d array from employeeData
$parsedEmployeeData = array();
foreach ($employeeData as $employee) {
  array_push($parsedEmployeeData, array($employee->firstName, $employee->lastName, $employee->position, $employee->department, $employee->salary, $employee->salesThisYear));
};

// Get cell references for TOTAL headers and cells for formulae and highest row
$totalHeaderCellReference = "E" . count($employeeData) + 5;
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
// Set column widths

$writer = new Xlsx($spreadsheet);
$writer->save('output.xlsx');
