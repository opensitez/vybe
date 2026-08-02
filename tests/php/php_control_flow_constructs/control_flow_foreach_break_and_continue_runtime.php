<?php
// vybe-test: php/php_control_flow_constructs/control_flow_foreach_break_and_continue_runtime
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

$values = [1, 2, 3, 4];
$sum = 0;
$i = 0;
foreach ($values as $n) {
    $i++;
    if ($n === 2) {
        continue;
    }
    if ($n === 4) {
        break;
    }
    $sum += $n;
}
echo $sum . '|' . $i;

__vybe_check(ob_get_clean(), "4|4");
