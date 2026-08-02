<?php
// vybe-test: php/php_mbstring_unicode/php_mbstring_unicode_edge_cases_runtime
// origin: languages/php/tests/php/test_php_mbstring_unicode.rs

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

echo mb_check_encoding("Hello", ["ASCII", "UTF-8"]) !== false ? "ok" : "bad";
echo "|";
echo mb_stripos("Café", "CAFÉ", 0, "UTF-8");
echo "|";
echo mb_str_split("😀", 3, "UTF-8")[0] === "😀" ? "one" : "no";

__vybe_check(ob_get_clean(), "ok|0|one");
