<?php
// vybe-test: php/url_functions/parse_url_php_url_query_component
// origin: languages/php/tests/php/test_url_functions.rs

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

echo parse_url('https://example.com/?q=hello&p=2', PHP_URL_QUERY);

__vybe_check(ob_get_clean(), "q=hello&p=2");
