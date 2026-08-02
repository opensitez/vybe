<?php
// vybe-test: php/string_regex_extended/preg_replace_callback_transform
// origin: languages/php/tests/php/test_string_regex_extended.rs

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

$r = preg_replace_callback('/\d+/', fn($m) => $m[0] * 2, 'I have 3 apples and 5 bananas');
echo $r;
echo "\n";

__vybe_check(ob_get_clean(), "I have 6 apples and 10 bananas");
