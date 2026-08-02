<?php
// vybe-test: php/eval/eval_switch_picks_matching_case
// origin: languages/php/tests/php/test_eval.rs

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

$code = 2;
eval('switch ($code) { case 2: $out = "two"; break; default: $out = "other"; }');
echo $out;

__vybe_check(ob_get_clean(), "two");
