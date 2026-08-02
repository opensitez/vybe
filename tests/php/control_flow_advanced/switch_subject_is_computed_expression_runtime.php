<?php
// vybe-test: php/control_flow_advanced/switch_subject_is_computed_expression_runtime
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

$index = 1;
$total = '';
switch ($index + 1) {
    case 1:
        $total = 'one';
        break;
    case 2:
        $total = 'two';
        break;
    default:
        $total = 'other';
}
echo $total;

__vybe_check(ob_get_clean(), "two");
