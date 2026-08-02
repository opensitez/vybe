<?php
// vybe-test: php/array_builtins_extended/array_merge_union_and_numeric_index_behavior
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

$a = [0 => "a", 1 => "b", "x" => 1];
$b = [0 => "z", 2 => "c", "y" => 2];
$m = array_merge($a, $b);
$u = $a + $b;
echo $m[0] . $m[1] . $m[2] . $m[3];
echo "|";
echo $u[0] . $u[1] . $u["x"] . (array_key_exists("y", $u) ? "y" : "n");

__vybe_check(ob_get_clean(), "abzc|ab1y");
