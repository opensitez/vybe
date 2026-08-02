<?php
// vybe-test: php/destructure_assignment/list_destructure_from_generator_throws_error
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

function g(): Generator { yield 1; yield 2; }
try { [$a, $b] = g(); echo 'ok'; }
catch (\Error $e) { echo 'gen-src'; }

__vybe_check(ob_get_clean(), "gen-src");
