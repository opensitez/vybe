<?php
// vybe-test: php/string_functions_extended/rawurlencode_space_and_question_runtime
// origin: languages/php/tests/php/test_string_functions_extended.rs

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

echo rawurlencode('a b?c');
echo '|';
echo rawurldecode('a%20b%3Fc');

__vybe_check(ob_get_clean(), "a%20b%3Fc|a b?c");
