<?php
// vybe-test: php/string_regex_extended/regex_offset_capture_starts_after_offset
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

$matches = [];
$result = preg_match('/a/', 'abcab', $matches, PREG_OFFSET_CAPTURE, 2);
echo $result;
echo '|';
echo $matches[0][0];
echo '|';
echo $matches[0][1];

__vybe_check(ob_get_clean(), "1|a|3");
