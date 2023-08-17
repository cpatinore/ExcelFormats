<?php
function explodeRange($range)
{

    $cells = explode(':', $range);
    $explodeRange = [];

    if (count($cells) == 2) {
        [$cellStart, $cellEnd ] = $cells;

        $explodeRange[0] = preg_split('/(?<=[A-Z])(?=\d)/', $cellStart);
        $explodeRange[1] = preg_split('/(?<=[A-Z])(?=\d)/', $cellEnd);
    }

    return $explodeRange;
}