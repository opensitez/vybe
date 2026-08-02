<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_match_in_loop_with_falsy_subject_edge
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

$sum = 0;
for ($i = 0; $i < 4; $i++) {
    $sum += match ((string)$i) {
        '' => 0,
        '0' => 1,
        '1' => 2,
        '2' => 3,
        default => 4,
    };
}
echo $sum;

__vybe_check(ob_get_clean(), "10");
