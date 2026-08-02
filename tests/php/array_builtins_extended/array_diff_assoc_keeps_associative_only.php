<?php
// vybe-test: php/array_builtins_extended/array_diff_assoc_keeps_associative_only
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

$a = ["x" => 1, "y" => 2, "z" => 3];
$b = ["x" => 9, "y" => 2];
$d = array_diff_assoc($a, $b);
ksort($d);
echo json_encode(array_keys($d)) . "|" . $d["z"];

__vybe_check(ob_get_clean(), "[\"x\",\"z\"]|3");
