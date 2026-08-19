<?php
// vybe-test: php/array_uintersect_uassoc_callback/array_uintersect_uassoc_numeric_key_casting
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

$a = ["1" => "A", 1 => "B"];
$b = ["01" => "a", "1" => "a"];
$r = array_uintersect_uassoc($a, $b, fn($v1,$v2)=>strcasecmp((string)$v1, (string)$v2), function($k1, $k2){ return (string)$k1 <=> (string)$k2; });
ksort($r);
echo implode('|', array_keys($r));

__vybe_check(ob_get_clean(), "");
