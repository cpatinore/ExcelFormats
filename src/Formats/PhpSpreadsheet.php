<?php

namespace ExcelFormats\Formats;

use ExcelFormats\Interfaces\iFileExcel;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Style;

class PhpSpreadsheet implements iFileExcel
{
    private $src_file;
    private $objExcel;
    private array $addedRows = array(0 => 0);

    function __construct(string $src_file)
    {
        $this->src_file = $src_file;
    }

    /*-----------------START OVERIDE-----------------*/
    function selectTemplateExcel(): Spreadsheet
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

    function copyCells(string $fromRange, string $toRange, array $mergeCells = null): void
    {
        $worksheet = $this->objExcel->getActiveSheet();

        $fromStartIt = explodeRange($fromRange)[0];
        $fromEndIt = explodeRange($fromRange)[1];

        $fromRange = explode(":", $fromRange);
        $toRange = explode(":", $toRange);

        $ColumnOffset = ord($toRange[0][0]) - ord($fromRange[0][0]);
        $RowOffset = intval(substr($toRange[0], 1)) - intval(substr($fromRange[0], 1));

        for ($fromRow = intval($fromStartIt[1]); $fromRow <= intval($fromEndIt[1]); $fromRow++) {
            
            for ($fromCol = $fromStartIt[0]; $fromCol <= $fromEndIt[0]; $fromCol++) {
                
                $fromCell = $fromCol . $fromRow;
                $toCell = chr(ord($fromCol) + $ColumnOffset) . ($fromRow + $RowOffset);

                $cellValue = $worksheet->getCell($fromCell)->getValue();
                $worksheet->setCellValue($toCell, $cellValue);

                $cellStyle = $worksheet->getStyle($fromCell);
                $worksheet->duplicateStyle($cellStyle, $toCell);
                
            
                $this->numberFormat($fromCell, $toCell);
                $worksheet->getRowDimension(($fromRow + $RowOffset))->setRowHeight($worksheet->getRowDimension($fromRow)->getRowHeight());
            }
        }

        $this->mergeCell(implode(":", $fromRange), implode(":", $toRange), "range", $mergeCells);
    }

    function activeSheet($sheet): void
    {
        if (is_numeric($sheet))
            $this->objExcel->setActiveSheetIndex($sheet);
        else
            $this->objExcel->setActiveSheetIndexByName($sheet);
    }

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

    function addRow(int $row, int $cant): void
    {
        $worksheet = $this->objExcel->getActiveSheet();
        $sheetIndex = $this->objExcel->getIndex($worksheet);

        $worksheet->insertNewRowBefore($row, $cant);
        $this->addedRows[$sheetIndex] += $cant;
    }

    function removeRow(int $row, int $cant): void
    {
        $worksheet = $this->objExcel->getActiveSheet();
        $sheetIndex = $this->objExcel->getIndex($worksheet);

        $worksheet->removeRow($row, $cant);
        $this->addedRows[$sheetIndex] -= $cant;
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

    function mergeCells(string $range): void
    {
        $this->objExcel->getActiveSheet()->mergeCells($range);
    }

    function unmergeCells(string $range): void
    {
        $this->objExcel->getActiveSheet()->unmergeCells($range);
    }

    function getMergeCells(): array
    {
        return $this->objExcel->getActiveSheet()->getMergeCells();
    }

    function getAddedRows(): int
    {
        $worksheet = $this->objExcel->getActiveSheet();
        $sheetIndex = $this->objExcel->getIndex($worksheet);
        return $this->addedRows[$sheetIndex];
    }

    function saveExcel($path): array
    {
        $objWriter = IOFactory::createWriter($this->objExcel, "Xlsx");
        $objWriter->save($path);

        $ruta_carpeta = dirname($path);

        return [$path, $ruta_carpeta];
    }

    /*-----------------END OVERIDE-----------------*/

    /**
     * Fill row by single value
     * @param array $content: Column values
     * @param int $row: Row number to be filled
     * 
     * @return void
     */
    function fillCellByValue(array $content, int $row): void
    {
        $worksheet = $this->objExcel->getActiveSheet();
        $mergedCells = $worksheet->getMergeCells();
        foreach ($content as $col => $value) {
            $worksheet->setCellValue([$col, $row], $value);
            $this->setHeightRowByCell($col, $row, $mergedCells);
        }
    }

    /**
     * Fill rows by array of values (table format)
     * @param array $confgTable: A key column array, where 
     * finishRow indicates the column to finish and 
     * data the information to fill per row. 
     * @param int $row: Row number to start filling in the values
     * 
     * @return void
     */
    function fillCellsByArray(array $confgTable, int $row): void
    {

        $data = $confgTable["data"];

        if (count($data) > 0) {

            unset($confgTable["data"]);
            $confg = $confgTable;

            $cols_using = array_keys($confg);
            $minCol = min($cols_using);
            $maxCol = max($cols_using);
            $cant = count($data);

            $this->addRow($row, $cant);
            $this->applyFormatByRowCol($minCol, $maxCol, $row, $cant);

            $worksheet = $this->objExcel->getActiveSheet();
            $mergedCells = $worksheet->getMergeCells();
            foreach ($data as $i => $row_data) {
                $row_data = $data[$i];
                foreach ($confg as $col => $id) {
                    if ($id == 'finishRow' or !array_key_exists($id, $row_data))
                        continue;
                    $value = $row_data[$id];
                    $worksheet->setCellValue([$col, $row + $i], $value);
                    $this->setHeightRowByCell($col, $row + $i, $mergedCells);
                }
            }
        }
    }

    /**
     * 
     * @param int $minCol: Initial column of the range to be copied
     * @param int $maxCol: End column of the range to be copied
     * @param int $row: Row number to copy format
     * @param int $cant: Number of rows to copy the format
     * 
     * @return void
     */
    function applyFormatByRowCol(int $minCol, int $maxCol, int $row, int $cant): void
    {
        $fromRow = $row + $cant;
        $worksheet = $this->objExcel->getActiveSheet();
        $sourceRange = Coordinate::stringFromColumnIndex(strval($minCol)) . "$fromRow:" .
            Coordinate::stringFromColumnIndex(strval($maxCol)) . $fromRow;
        $targetRange = Coordinate::stringFromColumnIndex(strval($minCol)) . "$row:" .
            Coordinate::stringFromColumnIndex(strval($maxCol)) . ($fromRow - 1);

        $sourceStyle = $worksheet->getStyle($sourceRange);
        $worksheet->duplicateStyle($sourceStyle, $targetRange);
        $this->mergeCell($sourceRange, $targetRange);
    }

    /**
     * @param int $col
     * @param int $row
     * @param array $mergedCells
     * 
     * @return void
     */
    function setHeightRowByCell(int $col, int $row, array $mergedCells): void
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
        $cellStyle = $worksheet->getStyle($targetCell);
        $fontSize = $cellStyle->getFont()->getSize();
        $totalWidth = $worksheet->getColumnDimension(Coordinate::stringFromColumnIndex($col))->getWidth();
        
        if (count($keys) > 0) {
            [$startCell, $endCell] = explode(":", $keys[0]);
            $startCol = Coordinate::columnIndexFromString($worksheet->getCell($startCell)->getColumn());
            $endCol = Coordinate::columnIndexFromString($worksheet->getCell($endCell)->getColumn());
            $totalWidth *= abs($endCol - $startCol);
        }
        
        $newHeight = ceil((strlen($worksheet->getCell([$col, $row])->getValue()) + 1) / $totalWidth) * $fontSize * 1.8;

        
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

    /**
     * @param string $mergedRange
     * @param string $sourceRange
     * @param string $targetRange
     * 
     * @return void
     */
    function mergeCellBySource(string $mergedRange, string $sourceRange, string $targetRange): void
    {
        $sourceRange = explode(":", $sourceRange);
        $targetRange = explode(":", $targetRange);
        $destinationRowOffset = intval(substr($targetRange[0], 1)) - intval(substr($sourceRange[0], 1));
        [$cellStart, $cellEnd] = explodeRange($mergedRange);
        $formCell = implode("", $cellStart);

        $cellStart[1] = intval($cellStart[1]) + $destinationRowOffset;
        $cellEnd[1] = intval($cellEnd[1]) + $destinationRowOffset;
        $this->numberFormat($formCell, implode("", $cellStart));
        $this->objExcel->getActiveSheet()->mergeCells(implode("", $cellStart) . ":" . implode("", $cellEnd));
    }

    /**
     * @param string $mergedRange
     * @param string $sourceRange
     * @param string $targetRange
     * 
     * @return void
     */
    function mergeTableBySource(array $mergedRange, string $sourceRange, string $targetRange): void
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
            $this->numberFormat($mergedRange[0][0], "$columnStart$rowMergeStart");
            $this->mergeCells("$columnStart$rowMergeStart:$columnEnd$rowMergeEnd");
        }
    }

    /**
     * @param mixed $sourceRange
     * @param mixed $targetRange
     * @param string $type
     * @param array|null $mergeCells
     * 
     * @return void
     */
    function mergeCell($sourceRange, $targetRange, $type = "table", array $mergeCells = null): void
    {
        $worksheet = $this->objExcel->getActiveSheet();
        if (!isset($mergeCells))
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

    function numberFormat($fromCell, $toCell)
    {
        $worksheet = $this->objExcel->getActiveSheet();
        $fromCellStyle = $worksheet->getStyle($fromCell);
        $formCellStyleArray = $fromCellStyle->getNumberFormat()->getFormatCode();
        $worksheet->getStyle($toCell)->getNumberFormat()->setFormatCode($formCellStyleArray);
    }


}