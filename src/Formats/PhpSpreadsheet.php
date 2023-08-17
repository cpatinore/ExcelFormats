<?php

namespace ExcelFormats\Formats;

use ExcelFormats\Interfaces\iFileExcel;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style;

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
    function fillCells($data): void
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

    function fillCellsByArray($confg_array, $row)
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
        $this->mergeCell($sourceRange, $targetRange);
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

    function copyCells($sourceRange, $destinationRange)
    {
        $worksheet = $this->objExcel->getActiveSheet();

        $sourceRange = explode(":", $sourceRange);
        $destinationRange = explode(":", $destinationRange);

        $sourceCellIterator = $worksheet->getHighestRowAndColumn();
        $destinationColumnOffset = ord($destinationRange[0][0]) - ord($sourceRange[0][0]);
        $destinationRowOffset = intval(substr($destinationRange[0], 1)) - intval(substr($sourceRange[0], 1));

        for ($row = 1; $row <= $sourceCellIterator['row']; $row++) {
            for ($col = 'A'; $col <= $sourceCellIterator['column']; $col++) {
                $cellCoordinate = $col . $row;
                $destinationCellCoordinate = chr(ord($col) + $destinationColumnOffset) . ($row + $destinationRowOffset);

                $cellValue = $worksheet->getCell($cellCoordinate)->getValue();
                $worksheet->setCellValue($destinationCellCoordinate, $cellValue);

                $cellStyle = $worksheet->getStyle($cellCoordinate);
                $worksheet->duplicateStyle($cellStyle, $destinationCellCoordinate);

                $worksheet->getColumnDimension($col)->setWidth($worksheet->getColumnDimension($col)->getWidth());
                $worksheet->getRowDimension($row)->setRowHeight($worksheet->getRowDimension($row)->getRowHeight());
            }
        }

        $this->mergeCell(implode(":", $sourceRange), implode(":", $destinationRange), "range");
    }

    function mergeCells($range)
    {
        $this->objExcel->getActiveSheet()->mergeCells($range);
    }

    function mergeCellBySource($mergedRange, $sourceRange, $targetRange)
    {
        $sourceRange = explode(":", $sourceRange);
        $targetRange = explode(":", $targetRange);
        $destinationRowOffset = intval(substr($targetRange[0], 1)) - intval(substr($sourceRange[0], 1));
        [$cellStart, $cellEnd] = explodeRange($mergedRange);
        
        $cellStart[1] = intval($cellStart[1]) + $destinationRowOffset;
        $cellEnd[1] = intval($cellEnd[1]) + $destinationRowOffset;
        
        $this->objExcel->getActiveSheet()->mergeCells(implode("", $cellStart) . ":" . implode("", $cellEnd));
    }

    function mergeTableBySource($mergedRange, $sourceRange, $targetRange)
    {
        list($columnStart, $rowStart) = Coordinate::coordinateFromString($mergedRange[0][0]);
        list($columnEnd, $rowEnd) = Coordinate::coordinateFromString($mergedRange[0][1]);

        [$toStart, $a] = explodeRange($sourceRange);
        [$fromStart, $b] = explodeRange($targetRange);

        $toRow = intval($toStart[1]);
        $fromRow = intval($fromStart[1]);
        for ($i = 0; $i < $toRow - $fromRow; $i++) {
            $rowMergeStart = $fromRow + $i;
            $rowMergeEnd = $fromRow + $rowEnd - $rowStart + $i;
            $this->mergeCells("$columnStart$rowMergeStart:$columnEnd$rowMergeEnd");
        }
    }

    function mergeCell($sourceRange, $targetRange, $type = "table"): void
    {
        $worksheet = $this->objExcel->getActiveSheet();
        $mergeCells = $worksheet->getMergeCells();
        foreach ($mergeCells as $mergeCell) {
            $mergedRange = Coordinate::splitRange($mergeCell);
            foreach ($mergedRange as $range) {
                [$startCell, $endCell] = $range;
                if (($worksheet->getCell($startCell)->isInRange($sourceRange) && $worksheet->getCell($endCell)->isInRange($sourceRange))) {
                    if ($type == "table") {
                        $this->mergeTableBySource($mergedRange, $sourceRange, $targetRange);
                    } else {
                        $this->mergeCellBySource($mergeCell, $sourceRange, $targetRange);
                    }
                    break;
                }
            }


        }
    }

    function saveExcel($path)
    {
        $objWriter = IOFactory::createWriter($this->objExcel, "Xlsx");
        $objWriter->save($path);

        $ruta_carpeta = dirname($path);

        return [$path, $ruta_carpeta];
    }
}