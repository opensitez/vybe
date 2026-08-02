<?php
// vybe-test: php/programs/roman_clock_hours
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

function intToRoman(int $n): string {
    $vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
    $syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
    $r = '';
    foreach ($vals as $i => $v) { while ($n >= $v) { $r .= $syms[$i]; $n -= $v; } }
    return $r;
}
foreach ([1, 4, 8, 12] as $h) echo intToRoman($h) . "\n";

__vybe_check(ob_get_clean(), "I\nIV\nVIII\nXII");
