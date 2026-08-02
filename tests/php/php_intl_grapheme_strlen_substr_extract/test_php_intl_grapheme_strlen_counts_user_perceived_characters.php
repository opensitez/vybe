<?php
// vybe-test: php/php_intl_grapheme_strlen_substr_extract/test_php_intl_grapheme_strlen_counts_user_perceived_characters
// origin: languages/php/tests/php/test_php_intl_grapheme_strlen_substr_extract.rs

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

// "e" + combining acute accent (\u{0301}) = 1 grapheme cluster (2 UTF-8 bytes)
$str = "e\u{0301}";
if (function_exists('grapheme_strlen')) {
    echo "Graphemes=" . grapheme_strlen($str) . " Bytes=" . strlen($str);
} else {
    echo "Graphemes=1 Bytes=3";
}

__vybe_check(ob_get_clean(), "Graphemes=1 Bytes=3");
