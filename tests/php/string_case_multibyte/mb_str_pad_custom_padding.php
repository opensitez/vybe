<?php
// vybe-test: php/string_case_multibyte/mb_str_pad_custom_padding
// origin: languages/php/tests/php/test_string_case_multibyte.rs

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

if (function_exists('mb_str_pad')) {
    echo mb_str_pad('é', 3, '0', STR_PAD_LEFT);
} else {
    echo str_pad('é', 3, '0', STR_PAD_LEFT);
}
echo '|';
if (function_exists('mb_str_pad')) {
    echo mb_str_pad('é', 2, '*', STR_PAD_BOTH);
} else {
    echo str_pad('é', 2, '*', STR_PAD_BOTH);
}

__vybe_check(ob_get_clean(), "00é|é*");
