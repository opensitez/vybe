<?php
// vybe-test: php/php_tokenizer_token_get_all_tokens/test_php_tokenizer_token_get_all_token_parse_flag
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

$code = "<?php \$x = 10;";
$tokens = token_get_all($code, TOKEN_PARSE);
echo is_array($tokens) && count($tokens) > 0 ? "PARSE_TOKENS_OK" : "FAIL";

__vybe_check(ob_get_clean(), "PARSE_TOKENS_OK");
