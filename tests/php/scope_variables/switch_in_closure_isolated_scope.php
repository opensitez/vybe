<?php
// vybe-test: php/scope_variables/switch_in_closure_isolated_scope
// origin: languages/php/tests/php/test_scope_variables.rs

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

$flag = 2;
$fn = function (int $v): int {
    switch ($v) {
        case 1: return 10;
        case 2: return 20;
        default: return 30;
    }
};
echo $fn($flag);

__vybe_check(ob_get_clean(), "20");
