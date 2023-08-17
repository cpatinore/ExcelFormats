<?php

namespace ExcelFormats;

use ExcelFormats\Interfaces\iFileExcel;

class ExcelFormat
{
    private iFileExcel $iFileExcel;

    function __construct(iFileExcel $iFileExcel)
    {
        $this->iFileExcel = $iFileExcel;
    }

    function selectTemplateExcel()
    {
        return $this->iFileExcel->selectTemplateExcel();
    }
    function duplicateSheet($fromSheet, $toSheet)
    {
        return $this->iFileExcel->duplicateSheet($fromSheet, $toSheet);
    }
    function activeSheet($sheet)
    {
        return $this->iFileExcel->activeSheet($sheet);
    }

    function fillCells($data)
    {
        $this->iFileExcel->fillCells($data);
    }

    function addPaginator($cell)
    {
        $this->iFileExcel->addPaginator($cell);
    }

    function addHeader($range)
    {
        $this->iFileExcel->addHeader($range);
    }
    function addFooter($footer)
    {
        $this->iFileExcel->addFooter($footer);
    }
    function copyCells($sourceRange, $destinationRange){
        $this->iFileExcel->copyCells($sourceRange, $destinationRange);
    }
    function getMergeCells() {
        $this->iFileExcel->getMergeCells();
    }

    function setUpFormat($confgFormat)
    {

        if (!isset($confgFormat))
            return;

        $paginator = isset($confgFormat["paginator"]) ? $confgFormat["paginator"] : null;

        if (isset($paginator))
            $this->addPaginator($paginator);

        if (isset($confgFormat["header"]) and count($confgFormat["header"]) == 2)
            $this->addHeader($confgFormat["header"]);

        if (isset($confgFormat["footer"]))
            $this->addFooter($confgFormat["footer"]);

    }

    public function createFormat($confgFormat = null)
    {
        $this->selectTemplateExcel();
        $this->setUpFormat($confgFormat);
    }

    public function saveExcel(string $path)
    {
        return $this->iFileExcel->saveExcel($path);
    }
}