<?php
// vybe-test: php/php_string_manipulation/php_string_list_and_join_chain
// origin: languages/php/tests/php/test_php_string_manipulation.rs

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

$parts = ['  a ', ' b ', 'c '];
$normalized = array_map('trim', $parts);
$joined = implode(':', $normalized);
$tokens = explode(':', $joined);
echo $tokens[0] . '|' . $tokens[2];

__vybe_check(ob_get_clean(), "a|c");
