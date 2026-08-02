<?php
// vybe-test: php/php_regular_expressions_pcre_matching/test_php_preg_split_pattern_splitting
// origin: languages/php/tests/php/test_php_regular_expressions_pcre_matching.rs

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

$keywords = preg_split('/[\s,]+/', "hypertext language, programming");
echo implode("|", $keywords);

__vybe_check(ob_get_clean(), "hypertext|language|programming");
