<?php
// vybe-test: php/arrays/array_diff_and_intersect_runtime
// origin: languages/php/tests/php/test_arrays.rs

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

$a = [1, 2, 3, 4];
$b = [2, 4];
$d = array_diff($a, $b);
$i = array_intersect($a, $b);
$items = [];
foreach ($d as $v) { $items[] = $v; }
foreach ($i as $v) { $items[] = "i$v"; }
echo implode('-', $items);

__vybe_check(ob_get_clean(), "1-3-i2-i4");
