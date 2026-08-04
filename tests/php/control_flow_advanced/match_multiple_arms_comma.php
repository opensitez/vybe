<?php
// vybe-test: php/control_flow_advanced/match_multiple_arms_comma
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

function classify(int $n): string {
    return match(true) {
        $n < 0 => 'negative',
        $n === 0 => 'zero',
        $n < 10 => 'small',
        $n < 100 => 'medium',
        default => 'large' };
}
echo classify(-5) . ',' . classify(0) . ',' . classify(7) . ',' . classify(50) . ',' . classify(200);

__vybe_check(ob_get_clean(), "negative,zero,small,medium,large");
