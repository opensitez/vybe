<?php
// vybe-test: php/modern_php_deep/spread_in_various_contexts
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

function sum(int ...$nums): int {
    return array_sum($nums);
}
echo sum(1, 2, 3);
echo sum(...[4, 5, 6]);

$a = [1, 2, 3];
$b = [4, 5, 6];
$merged = [...$a, ...$b];
echo implode(",", $merged);

__vybe_check(ob_get_clean(), "6151,2,3,4,5,6");
