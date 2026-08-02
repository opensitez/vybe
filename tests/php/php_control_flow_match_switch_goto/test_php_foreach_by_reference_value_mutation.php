<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_foreach_by_reference_value_mutation
// origin: languages/php/tests/php/test_php_control_flow_match_switch_goto.rs

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

$nums = [1, 2, 3];
foreach ($nums as &$val) {
    $val *= 10;
}
unset($val); // break reference binding
echo implode("-", $nums);

__vybe_check(ob_get_clean(), "10-20-30");
