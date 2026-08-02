<?php
// vybe-test: php/php_string_searching_substring_positions/test_php_strstr_before_and_after_needle
// origin: languages/php/tests/php/test_php_string_searching_substring_positions.rs

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

$email = "user@example.com";
$domain = strstr($email, "@");
$user = strstr($email, "@", before_needle: true);
echo "$user from $domain";

__vybe_check(ob_get_clean(), "user from @example.com");
