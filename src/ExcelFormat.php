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

    function setUpFormat($confgFormat)
    {

        if (!isset($confgFormat))
            return;

        $paginator = isset($confgFormat["paginator"]) ? $confgFormat["paginator"] : null;

        if (isset($paginator))
            $this->addPaginator($paginator);

        if (isset($confgFormat["header"]) and count($confgFormat["header"]) == 2)
            $this->addHeader($confgFormat["header"]);

    }

    public function createFormat($data, $confgFormat = null)
    {
        $this->selectTemplateExcel();
        $this->setUpFormat($confgFormat);
        $this->fillCells($data);
    }

    public function saveExcel(string $path)
    {
        return $this->iFileExcel->saveExcel($path);
    }
}