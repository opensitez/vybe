<?php
// vybe-test: php/eval/eval_match_expression_returns_value
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

eval('$m = match (3) { 1 => "a", 3 => "c", default => "z" };');
echo $m;

__vybe_check(ob_get_clean(), "c");
