<?php
// vybe-test: php/enums_advanced/enum_backed_cases_values
// origin: languages/php/tests/php/test_enums_advanced.rs

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

enum Priority: int { case Low = 1; case Mid = 2; case High = 3; }
echo implode(',', array_map(fn($c) => $c->value, Priority::cases()));

__vybe_check(ob_get_clean(), "1,2,3");
