<?php
// vybe-test: php/php_string_manipulation_formatting/test_php_string_sprintf_format_specifiers
// origin: languages/php/tests/php/test_php_string_manipulation_formatting.rs

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

$formatted = sprintf("User %s ID %04d Price $%.2f", "Alice", 42, 19.95);
echo $formatted;

__vybe_check(ob_get_clean(), "User Alice ID 0042 Price \$19.95");
