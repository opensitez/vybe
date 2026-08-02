<?php
// vybe-test: php/array_uintersect_uassoc_callback/array_uintersect_uassoc_key_as_numeric_string
// origin: languages/php/tests/php/test_array_uintersect_uassoc_callback.rs

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

$a = ["1" => "A", "2" => "B", "x" => "X"];
$b = [1 => "a", "2" => "b"];
$r = array_uintersect_uassoc($a, $b, fn($v1,$v2)=>strcmp($v1,$v2), "strcasecmp");
ksort($r);
echo implode('|', array_keys($r)) . ":" . $r["1"];

__vybe_check(ob_get_clean(), "1|2:A");
