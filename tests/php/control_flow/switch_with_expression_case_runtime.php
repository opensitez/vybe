<?php
// vybe-test: php/control_flow/switch_with_expression_case_runtime
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

$mode = 4;
$out = '';
switch ($mode) {
    case 1 + 1:
        $out .= '2';
        break;
    case 2 + 2:
        $out .= '4';
        break;
    default:
        $out .= 'd';
}
echo $out;

__vybe_check(ob_get_clean(), "4");
