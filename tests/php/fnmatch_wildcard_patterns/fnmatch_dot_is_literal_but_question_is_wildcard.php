<?php
// vybe-test: php/fnmatch_wildcard_patterns/fnmatch_dot_is_literal_but_question_is_wildcard
// origin: languages/php/tests/php/test_fnmatch_wildcard_patterns.rs

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

echo fnmatch("a.b", "aXb") ? "1" : "0";
echo fnmatch("a.b", "a.b") ? "1" : "0";
echo fnmatch("a?b", "aXb") ? "1" : "0";

__vybe_check(ob_get_clean(), "011");
