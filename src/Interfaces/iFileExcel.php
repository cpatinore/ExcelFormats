<?php

namespace ExcelFormats\Interfaces;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

interface iFileExcel
{
    /**
     * Select template or create blank format
     * 
     * @return Spreadsheet
     */
    function selectTemplateExcel(): Spreadsheet;

    /**
     * Duplicates a sheet based on an existing sheet in the workbook
     * @param string | int $fromSheet : Unique name or index of the sheet to be duplicated
     * @param string $toSheet : Unique name of the sheet duplicated
     * 
     * @return void
     */
    function duplicateSheet(string|int $fromSheet, string $toSheet): void;

    /**
     * Copy cells with value, formatting (styles, borders, width and height) and merged cells
     * @param string $sourceRange: Range of information source
     * @param string $destinationRange: Destination range of information
     * @param array|null $mergeCells: Combined cells to be evaluated if they are within 
     * the range (default takes all the combined cells of the active sheet)
     * 
     * @return void
     */
    function copyCells(string $sourceRange, string $destinationRange, array $mergeCells = null): void;

    /**
     * Active a worksheet 
     * @param int $sheet: Index unique of the worksheet to actived
     * 
     * @return void
     */
    function activeSheet(int $sheet): void;

    /**
     * Fill cells in worksheet by an array data (multiple lines) or uniqe value
     * @param array $data: Multidimensional array of rows by columns
     * 
     * @return void
     */
    function fillCells(array $data): void;

    /**
     * Inserts $cant new rows, right before row $row
     * @param int $row: Row to insert before
     * @param int $cant: Number of rows to insert
     * 
     * @return void
     */
    function addRow(int $row, int $cant): void;

    /**
     * Remove $cant rows starting at row number $row
     * @param int $row: Start row
     * @param int $cant: Number of rows to remove
     * 
     * @return void
     */
    function removeRow(int $row, int $cant): void;

    /**
     * Duplicate a range cell in all pages
     * @param array $range: Cell of the range to duplicate
     * 
     * @return void
     */
    function addHeader(array $range): void;

    /**
     * Add footer in the active worksheet
     * @param string $footer: worksheet footer
     * 
     * @return void
     */
    function addFooter(string $footer): void;

    /**
     * Merge two or more cells
     * @param string $range: Range of cells to merge
     * 
     * @return void
     */
    function mergeCells(string $range): void;

    /**
     * Removing a merge
     * @param string $range: Range of merged cells
     * 
     * @return void
     */
    function unmergeCells(string $range): void;

    /**
     * Get all merged cells of the active worksheet
     * @return array
     */
    function getMergeCells(): array;

    /**
     * Get number of rows added to the active spreadsheet
     * @return int
     */
    function getAddedRows(): int;

    /**
     * Save the created document in the specified path
     * @param string $path: Path to save the excel
     * 
     * @return array
     */
    function saveExcel(string $path): array;
}