<?php
// vybe-test: php/closures_runtime/closure_recursive_via_reference_use
// origin: languages/php/tests/php/test_closures_runtime.rs

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

$fact = null;
$fact = function (int $n) use (&$fact): int {
    return $n <= 1 ? 1 : $n * $fact($n - 1);
};
echo $fact(5);

__vybe_check(ob_get_clean(), "120");
