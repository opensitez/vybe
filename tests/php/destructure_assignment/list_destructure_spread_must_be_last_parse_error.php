<?php
// vybe-test: php/destructure_assignment/list_destructure_spread_must_be_last_parse_error
// origin: languages/php/tests/php/test_destructure_assignment.rs

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

try {
    eval('[$a, ...$b, $c] = [1,2,3];');
    echo 'ok';
} catch (ParseError $e) {
    echo 'parse';
}

__vybe_check(ob_get_clean(), "parse");
