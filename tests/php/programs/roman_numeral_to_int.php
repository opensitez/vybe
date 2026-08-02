<?php
// vybe-test: php/programs/roman_numeral_to_int
// origin: languages/php/tests/php/test_programs.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

function romanToInt(string $s): int {
    $map = ['I'=>1,'V'=>5,'X'=>10,'L'=>50,'C'=>100,'D'=>500,'M'=>1000];
    $result = 0;
    $prev = 0;
    foreach (array_reverse(str_split($s)) as $c) {
        $v = $map[$c];
        if ($v < $prev) $result -= $v;
        else $result += $v;
        $prev = $v;
    }
    return $result;
}
echo romanToInt('XIV') . "\n";
echo romanToInt('IX') . "\n";
echo romanToInt('XLII') . "\n";

__vybe_check(ob_get_clean(), "14\n9\n42");
