<?php

namespace ExcelFormats\Interfaces;

interface iFileExcel
{
    function selectTemplateExcel();

    /**
     * Duplicates a sheet based on an existing sheet in the workbook
     * @param string | int $fromSheet : Unique name or index of the sheet to be duplicated
     * @param string $toSheet : Unique name of the sheet duplicated
     * 
     * @return void
     */
    function duplicateSheet(string|int $fromSheet, string $toSheet): void;

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

    function addRow($row, $cant):void;

    function removeRow($row, $cant):void;

    function saveExcel(string $path);
    function addPaginator($cell): void;
    
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

    function copyCells($sourceRange, $destinationRange, $mergeCells = null);
    function getMergeCells(): array;
    function getAddedRows(): int;
}