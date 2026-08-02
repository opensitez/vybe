<?php
// vybe-test: php/php_regular_expressions_pcre_matching/test_php_preg_match_all_global_extraction
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

$text = "Emails: alice@domain.com, bob@example.org";
preg_match_all('/[\w.-]+@[\w.-]+/', $text, $matches);
echo implode(", ", $matches[0]);

__vybe_check(ob_get_clean(), "alice@domain.com, bob@example.org");
