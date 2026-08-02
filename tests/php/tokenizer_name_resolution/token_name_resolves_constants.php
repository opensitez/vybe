<?php
// vybe-test: php/tokenizer_name_resolution/token_name_resolves_constants
// origin: languages/php/tests/php/test_tokenizer_name_resolution.rs

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

echo token_name(T_CLASS) . '|' . token_name(T_FUNCTION) . '|' . token_name(T_STRING);

__vybe_check(ob_get_clean(), "T_CLASS|T_FUNCTION|T_STRING");
