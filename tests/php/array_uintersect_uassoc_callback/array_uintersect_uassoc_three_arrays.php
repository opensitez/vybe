<?php
// vybe-test: php/array_uintersect_uassoc_callback/array_uintersect_uassoc_three_arrays
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

$a = ["a" => "X", "b" => "Y"];
$b = ["a" => "x", "c" => "z"];
$c = ["a" => "x", "b" => "y"];
$r = array_uintersect_uassoc($a, $b, "strcasecmp", "strcasecmp");
$r = array_uintersect_uassoc($r, $c, "strcasecmp", "strcasecmp");
echo count($r) . "|" . implode(',', array_keys($r));

__vybe_check(ob_get_clean(), "1|a");
