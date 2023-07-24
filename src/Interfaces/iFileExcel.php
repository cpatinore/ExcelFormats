<?php

namespace ExcelFormats\Interfaces;
interface iFileExcel
{
    function selectTemplateExcel();
    function fillCells($data);
    function saveExcel($path);
    function addPaginator($cell);
    function addHeader($range);
}
