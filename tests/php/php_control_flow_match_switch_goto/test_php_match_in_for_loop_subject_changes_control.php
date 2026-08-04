<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_match_in_for_loop_subject_changes_control
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

$out = '';
for ($i = 0; $i < 4; $i++) {
    $out .= match ($i) {
        0, 1 => 'L',
        2 => 'M',
        default => 'H' };
}
echo $out;

__vybe_check(ob_get_clean(), "LLMH");
