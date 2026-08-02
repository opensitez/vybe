<?php
// vybe-test: php/array_builtins_extended/array_fill_keys_collision_keeps_last
// origin: languages/php/tests/php/test_array_builtins_extended.rs

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

$keys = ["x", "y", "x"];
$a = array_fill_keys($keys, 0);
$a["x"] = 9;
echo $a["x"] . "|" . count($a);

__vybe_check(ob_get_clean(), "9|2");
