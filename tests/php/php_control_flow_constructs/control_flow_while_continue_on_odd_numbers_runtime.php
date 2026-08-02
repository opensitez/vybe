<?php
// vybe-test: php/php_control_flow_constructs/control_flow_while_continue_on_odd_numbers_runtime
// origin: languages/php/tests/php/test_php_control_flow_constructs.rs

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

$i = 0;
$sum = 0;
while ($i < 6) {
    $i++;
    if ($i % 2 === 0) {
        continue;
    }
    $sum += $i;
}
echo $sum;

__vybe_check(ob_get_clean(), "9");
