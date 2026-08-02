<?php
// vybe-test: php/php_control_flow_match_switch_goto/test_php_goto_after_switch_case
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

$state = 0;
$out = '';
switch ($state) {
    case 0:
        $out .= 'pre';
        goto done;
    case 1:
        $out .= 'ignored';
}
done:
echo $out;

__vybe_check(ob_get_clean(), "pre");
