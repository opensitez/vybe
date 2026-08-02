<?php
// vybe-test: php/control_flow_advanced/switch_return_from_function
// origin: languages/php/tests/php/test_control_flow_advanced.rs

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

function day(int $n): string {
    switch ($n) {
        case 1: return 'Mon';
        case 2: return 'Tue';
        case 3: return 'Wed';
        default: return 'Other';
    }
}
echo day(2);

__vybe_check(ob_get_clean(), "Tue");
