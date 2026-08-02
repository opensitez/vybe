<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_switch_with_expression_cases_and_parentheses_precedence
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

$input = 4;
$out = '';
switch ($input) {
    case 2 + 2:
        $out .= 'four';
        break;
    case (int) '3':
        $out .= 'three';
        break;
    default:
        $out .= 'other';
}
echo $out;

__vybe_check(ob_get_clean(), "four");
