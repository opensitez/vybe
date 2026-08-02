<?php
// vybe-test: php/tokenizer_parse_error_recovery/token_get_all_token_parse_flag
// origin: languages/php/tests/php/test_tokenizer_parse_error_recovery.rs

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

// In PHP 8+, TOKEN_PARSE flag can throw ParseError on invalid syntax
$source = '<?php class { public }';
try {
    token_get_all($source, TOKEN_PARSE);
    echo "success";
} catch (\ParseError $e) {
    echo "error";
}

__vybe_check(ob_get_clean(), "error");
