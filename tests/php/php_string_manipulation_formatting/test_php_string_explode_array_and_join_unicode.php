<?php
// vybe-test: php/php_string_manipulation_formatting/test_php_string_explode_array_and_join_unicode
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

$text = "a,b, c, ";
$items = explode(',', $text);
$joined = implode('|', $items);
echo count($items) . '|' . substr($joined, 0, 12);

__vybe_check(ob_get_clean(), "6|a|b||c||");
