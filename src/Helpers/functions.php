<?php
function explodeRange($range)
{

    $cells = explode(':', $range);
    $explodeRange = [];

    if (count($cells) == 2) {
        [$cellStart, $cellEnd ] = $cells;

        $explodeRange["start"] = preg_split('/(?<=[A-Z])(?=\d)/', $cellStart);
        $explodeRange["end"] = preg_split('/(?<=[A-Z])(?=\d)/', $cellEnd);
    }
    
    return $explodeRange;
}