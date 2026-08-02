<?php
// vybe-test: php/array_udiff_uassoc_callback/array_udiff_uassoc_empty_first_array
// origin: languages/php/tests/php/test_array_udiff_uassoc_callback.rs

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

$a = [];
$b = ["a" => 1];
$r = array_udiff_uassoc($a, $b, fn($v1, $v2) => $v1 <=> $v2, fn($k1, $k2) => strcmp($k1, $k2));
echo is_array($r) ? 'array' : 'not-array';
echo '|';
echo count($r);

__vybe_check(ob_get_clean(), "array|0");
