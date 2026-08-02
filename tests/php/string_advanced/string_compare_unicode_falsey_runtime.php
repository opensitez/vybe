<?php
// vybe-test: php/string_advanced/string_compare_unicode_falsey_runtime
// origin: languages/php/tests/php/test_string_advanced.rs

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

echo strcmp("", "") === 0 ? "both_empty" : "diff";
echo "\n";
echo strcasecmp("ABC", "abc") === 0 ? "ci" : "noc";
echo "\n";
echo substr_compare("abcdef", "ab", 0, 0);

__vybe_check(ob_get_clean(), "both_empty\nci\n0");
