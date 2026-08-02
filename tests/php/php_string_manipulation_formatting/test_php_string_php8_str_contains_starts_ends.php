<?php
// vybe-test: php/php_string_manipulation_formatting/test_php_string_php8_str_contains_starts_ends
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

$haystack = "https://laravel.com/docs";
echo str_starts_with($haystack, "https://") ? "YES" : "NO";
echo " ";
echo str_ends_with($haystack, "/docs") ? "YES" : "NO";
echo " ";
echo str_contains($haystack, "laravel") ? "YES" : "NO";

__vybe_check(ob_get_clean(), "YES YES YES");
