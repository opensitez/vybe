<?php
// vybe-test: php/string_advanced/sprintf_with_precision_and_sign_runtime
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

echo sprintf("%+d", 12);
echo "\n";
echo sprintf("% 06d", 12);
echo "\n";
echo sprintf("%'_'9d", 42);

__vybe_check(ob_get_clean(), "+12\n000012\n42");
