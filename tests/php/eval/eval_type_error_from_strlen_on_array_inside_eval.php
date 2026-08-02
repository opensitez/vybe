<?php
// vybe-test: php/eval/eval_type_error_from_strlen_on_array_inside_eval
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

try { eval('strlen([]);'); echo 'ok'; }
catch (TypeError $e) { echo 'typed'; }

__vybe_check(ob_get_clean(), "typed");
