<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_goto_with_safe_local_jump_pattern
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

function guarded() {
    $x = 0;
    if ($x < 1) {
        goto skip;
    }
    $x = 10;
skip:
    return $x === 0 ? 'zero' : 'set';
}
echo guarded();

__vybe_check(ob_get_clean(), "zero");
