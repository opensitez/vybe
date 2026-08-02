<?php
// vybe-test: php/pcre_named_groups/preg_replace_callback_array_with_captures
// origin: languages/php/tests/php/test_pcre_named_groups.rs

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

$result = preg_replace_callback_array([
    '/(\d+) USD/' => fn($m) => $m[1] * 2 . ' USD',
    '/(\d+) EUR/' => fn($m) => $m[1] * 3 . ' EUR',
], '10 USD and 5 EUR');
echo $result;

__vybe_check(ob_get_clean(), "20 USD and 15 EUR");
