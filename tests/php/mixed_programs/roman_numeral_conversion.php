<?php
// vybe-test: php/mixed_programs/roman_numeral_conversion
// origin: languages/php/tests/php/test_mixed_programs.rs

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

function toRoman(int $n): string {
    $map = [1000=>'M',900=>'CM',500=>'D',400=>'CD',100=>'C',90=>'XC',50=>'L',40=>'XL',10=>'X',9=>'IX',5=>'V',4=>'IV',1=>'I'];
    $result = '';
    foreach ($map as $val => $sym) { while ($n >= $val) { $result .= $sym; $n -= $val; } }
    return $result;
}
echo toRoman(2024) . ',' . toRoman(42) . ',' . toRoman(1999);

__vybe_check(ob_get_clean(), "MMXXIV,XLII,MCMXCIX");
