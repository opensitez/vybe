<?php
// vybe-test: php/operators_runtime/precedence_between_addition_and_null_coalesce_runtime
// origin: languages/php/tests/php/test_operators_runtime.rs

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

echo 1 + 2 + 3 ?? 'fallback';
echo '|';
echo 0 + (null ?? 7);
echo '|';
echo 4 + (null ?? 1) . '';
echo '|';
echo (null ?? 1) + 4;

__vybe_check(ob_get_clean(), "6|7|5|5");
