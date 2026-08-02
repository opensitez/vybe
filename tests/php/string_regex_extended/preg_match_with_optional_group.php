<?php
// vybe-test: php/string_regex_extended/preg_match_with_optional_group
// origin: languages/php/tests/php/test_string_regex_extended.rs

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

preg_match('/^(?<x>\w+)(?:\s+(\d+))?$/', 'name 123', $m);
echo isset($m[2]) ? $m[2] : 'none';
echo '|';
preg_match('/^(?<x>\w+)(?:\s+(\d+))?$/', 'name', $m2);
echo isset($m2[2]) ? $m2[2] : 'none';

__vybe_check(ob_get_clean(), "123|none");
