<?php

namespace ExcelFormats\Formats;

use ExcelFormats\Interfaces\iFileExcel;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class PhpSpreadsheet implements iFileExcel
{
    private $src_file;
    private $objExcel;
    private array $addedRows = array(0 => 0);

    function __construct(string $src_file)
    {
        $this->src_file = $src_file;
    }

    function selectTemplateExcel()
    {
        if ($this->src_file === "")
            $this->objExcel = new Spreadsheet();
        else
            $this->objExcel = IOFactory::load($this->src_file);

        $this->objExcel->setActiveSheetIndex(0);

        return $this->objExcel;
    }

    function duplicateSheet($fromSheet, $toSheet): void
    {
        $clonedWorksheet = clone $this->objExcel->getSheetByName($fromSheet);
        $clonedWorksheet->setTitle($toSheet);
        $this->objExcel->addSheet($clonedWorksheet);
        $index = $this->objExcel->getIndex($clonedWorksheet);
        $this->addedRows[$index] = 0;
    }

    function activeSheet($sheet): void
    {
        if (is_numeric($sheet))
            $this->objExcel->setActiveSheetIndex($sheet);
        else
            $this->objExcel->setActiveSheetIndexByName($sheet);
    }

    // Fills
    function fillCells($data)
    {
        $worksheet = $this->objExcel->getActiveSheet();
        $sheetIndex = $this->objExcel->getIndex($worksheet);
        foreach ($data as $row => $content) {
            $row = intval($row) + $this->addedRows[$sheetIndex];
            if (isset($content["data"]))
                $this->fillCellsByArray($content, $row);
            else
                $this->fillCellByValue($content, $row);
        }
    }

    function fillCellByValue($content, $row): void
    {
        $worksheet = $this->objExcel->getActiveSheet();
        $mergedCells = $worksheet->getMergeCells();
        foreach ($content as $col => $value) {
            $worksheet->setCellValue([$col, $row], $value);
            $this->setHeightRowByCell($col, $row, $mergedCells);
        }
    }

    function fillCellsByArray($confg_array, $row): void
    {

        $data = $confg_array["data"];

        if (count($data) == 0)
            return $this->objExcel;

        unset($confg_array["data"]);
        $confg = $confg_array;

        $cols_using = array_keys($confg);
        $min_col = min($cols_using);
        $max_col = max($cols_using);

        $this->addRowWithStyle($row, count($data), $min_col, $max_col);

        $worksheet = $this->objExcel->getActiveSheet();
        $mergedCells = $worksheet->getMergeCells();
        foreach ($data as $i => $row_data) {
            $row_data = $data[$i];
            foreach ($confg as $col => $id) {
                if ($id == 'finishRow')
                    continue;
                $value = $row_data[$id];
                $worksheet->setCellValue([$col, $row + $i], $value);
                $this->setHeightRowByCell($col, $row + $i, $mergedCells);
            }
        }
    }


    // styles
    function applyFormatByRowCol($min_col, $max_col, $row, $cant): void
    {
        $from_row = $row + $cant;
        $worksheet = $this->objExcel->getActiveSheet();
        $sourceRange = Coordinate::stringFromColumnIndex(strval($min_col)) . "$from_row:" .
            Coordinate::stringFromColumnIndex(strval($max_col)) . $from_row;
        $targetRange = Coordinate::stringFromColumnIndex(strval($min_col)) . "$row:" .
            Coordinate::stringFromColumnIndex(strval($max_col)) . ($from_row - 1);

        $sourceStyle = $worksheet->getStyle($sourceRange);
        $worksheet->duplicateStyle($sourceStyle, $targetRange);
        $this->mergeCell($worksheet, $sourceRange, $from_row, $row);
    }

    function mergeCell($worksheet, $sourceRange, $from_row, $to_row): void
    {
        $mergeCells = $worksheet->getMergeCells($sourceRange);
        foreach ($mergeCells as $mergeCell) {
            $mergedRange = Coordinate::splitRange($mergeCell);
            $isWithinSourceRange = true;
            foreach ($mergedRange as $range) {
                [$startCell, $endCell] = $range;
                if (
                    !($worksheet->getCell($startCell)->isInRange($sourceRange) &&
                        $worksheet->getCell($endCell)->isInRange($sourceRange))
                ) {
                    $isWithinSourceRange = false;
                    break;
                }
            }

            if ($isWithinSourceRange) {
                list($columnStart, $rowStart) = Coordinate::coordinateFromString($mergedRange[0][0]);
                list($columnEnd, $rowEnd) = Coordinate::coordinateFromString($mergedRange[0][1]);

                for ($i = 0; $i < $from_row - $to_row; $i++) {
                    $rowMergeStart = $to_row + $i;
                    $rowMergeEnd = $to_row + $rowEnd - $rowStart + $i;
                    $worksheet->mergeCells("$columnStart$rowMergeStart:$columnEnd$rowMergeEnd");
                }
            }
        }
    }

    function setHeightRowByCell($col, $row, $mergedCells): void
    {
        $worksheet = $this->objExcel->getActiveSheet();
        $currentHeight = $worksheet->getRowDimension($row)->getRowHeight();
        $targetCell = Coordinate::stringFromColumnIndex($col) . "$row";

        $keys = array_filter(
            $mergedCells,
            function ($key) use ($targetCell) {
                return strpos($key, $targetCell) === 0;
            },
            ARRAY_FILTER_USE_KEY
        );

        $keys = array_keys($keys);
        $totalWidth = 12.89;
        if (count($keys) > 0) {
            [$startCell, $endCell] = explode(":", $keys[0]);
            $startCol = Coordinate::columnIndexFromString($worksheet->getCell($startCell)->getColumn());
            $endCol = Coordinate::columnIndexFromString($worksheet->getCell($endCell)->getColumn());
            $totalWidth *= ($endCol - $startCol);
        }

        $newHeight = ceil((strlen($worksheet->getCell([$col, $row])->getValue()) + 1) / $totalWidth) * 18;

        if ($currentHeight < $newHeight)
            $worksheet->getRowDimension($row)->setRowHeight($newHeight, 'px');
    }

    function duplicateByRange($fromRange, $toCell): void
    {
        $worksheet = $this->objExcel->getActiveSheet();

        $fromValueRange = $worksheet->getCell($fromRange)->getValue();
        $fromStyleRange = $worksheet->getStyle($fromRange);

        $worksheet->setCellValue($toCell, $fromValueRange);
        $worksheet->duplicateStyle($fromStyleRange, $toCell);
    }

    // adds
    function addRowWithStyle($row, $cant, $min_col, $max_col): void
    {
        $worksheet = $this->objExcel->getActiveSheet();
        $sheetIndex = $this->objExcel->getIndex($worksheet);
        $this->addedRows[$sheetIndex] += $cant;
        $this->addRow($row, $cant);
        $this->applyFormatByRowCol(
            $min_col,
            $max_col,
            $row,
            $cant
        );
    }

    function addRow($row, $cant): void
    {
        $this->objExcel->getActiveSheet()->insertNewRowBefore($row, $cant);
    }

    function addPaginator($cell): void
    {
        $this->objExcel->getActiveSheet()->setCellValue($cell, 'Página &P de &N');
    }

    function addHeader($range): void
    {
        $this->objExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(...$range);
    }

    function addFooter($footer): void
    {
        $this->objExcel
            ->getActiveSheet()
            ->getHeaderFooter()
            ->setOddFooter($footer);
    }

    function saveExcel($path): void
    {
        $objWriter = IOFactory::createWriter($this->objExcel, "Xlsx");
        $objWriter->save($path);

        $ruta_carpeta = dirname($path);

        return [$path, $ruta_carpeta];
    }
}