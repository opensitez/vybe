<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_switch_subject_mutation_in_case_condition
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

$x = 0;
$out = '';
switch (1) {
    case 1:
        $out .= 'first';
        $x = 2;
        // falls through by design:
    case 2:
        $out .= '-second';
        break;
    case 3:
        $out .= '-third';
        break;
}
echo $out;

__vybe_check(ob_get_clean(), "first-second");
