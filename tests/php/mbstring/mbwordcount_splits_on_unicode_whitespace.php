<?php
// vybe-test: php/mbstring/mbwordcount_splits_on_unicode_whitespace
// origin: languages/php/tests/php/test_mbstring.rs

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

function mb_word_count(string $s): int {
    return count(array_filter(preg_split('/\s+/u', trim($s)), fn($w) => $w !== ''));
}
echo mb_word_count('one two  three');

__vybe_check(ob_get_clean(), "3");
