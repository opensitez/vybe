<?php
// vybe-test: php/array_uintersect_uassoc_callback/array_uintersect_uassoc_key_numeric_compare
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

$a = ["01" => "x", 1 => "y", "2" => "z"];
$b = [1 => "Y", "2" => "Z"];
$r = array_uintersect_uassoc($a, $b, "strcasecmp", function($k1, $k2) { return (string)$k1 <=> (string)$k2; });
ksort($r);
echo implode('|', array_keys($r));

__vybe_check(ob_get_clean(), "1|2");
