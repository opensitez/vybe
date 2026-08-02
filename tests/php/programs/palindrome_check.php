<?php
// vybe-test: php/programs/palindrome_check
// origin: languages/php/tests/php/test_programs.rs

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

function isPalindrome(string $s): bool {
    $s = strtolower(preg_replace('/[^a-zA-Z0-9]/', '', $s));
    return $s === strrev($s);
}
echo isPalindrome('racecar') ? 'true' : 'false';
echo "\n";
echo isPalindrome('hello') ? 'true' : 'false';
echo "\n";
echo isPalindrome('A man a plan a canal Panama') ? 'true' : 'false';
echo "\n";

__vybe_check(ob_get_clean(), "true\nfalse\ntrue");
