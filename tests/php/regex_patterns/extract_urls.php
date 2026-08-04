<?php
// vybe-test: php/regex_patterns/extract_urls
// origin: languages/php/tests/php/test_regex_patterns.rs

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

$text = "Visit https://example.com and http://test.org/page for more";
preg_match_all('/https?:\/\/[\w.\/]+/', $text, $matches);
echo implode("\n", $matches[0]);

__vybe_check(ob_get_clean(), "https://example.com\nhttp://test.org/page");
