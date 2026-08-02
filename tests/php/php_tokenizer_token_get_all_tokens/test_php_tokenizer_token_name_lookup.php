<?php
// vybe-test: php/php_tokenizer_token_get_all_tokens/test_php_tokenizer_token_name_lookup
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_tokens.rs

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

echo token_name(T_ECHO) . " " . token_name(T_VARIABLE) . " " . token_name(T_FUNCTION);

__vybe_check(ob_get_clean(), "T_ECHO T_VARIABLE T_FUNCTION");
