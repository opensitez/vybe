<?php
// vybe-test: php/string_advanced/strcmp_strcasecmp
// origin: languages/php/tests/php/test_string_advanced.rs

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

echo strcmp("abc", "abc") === 0 ? "equal" : "not equal";
echo "\n";
echo strcmp("abc", "abd") < 0 ? "less" : "not less";
echo "\n";
echo strcasecmp("Hello", "hello") === 0 ? "equal" : "not equal";
echo "\n";
echo strcasecmp("ABC", "xyz") < 0 ? "less" : "not less";
echo "\n";

__vybe_check(ob_get_clean(), "equal\nless\nequal\nless");
