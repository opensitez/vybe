<?php
// vybe-test: php/control_flow/for_loop_multiple_update_expressions_runtime
// origin: languages/php/tests/php/test_control_flow.rs

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

$total = 0;
for ($i = 0, $j = 0; $i < 4; $i++, $j++) {
    if ($i === 2) {
        continue;
    }
    $total += $i + $j;
}
echo $total;

__vybe_check(ob_get_clean(), "8");
