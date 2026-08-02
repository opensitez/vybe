<?php
// vybe-test: php/php_tokenizer_token_get_all_tokens/test_php_tokenizer_token_line_number
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_tokens.rs
// vybe-test-mode: compile

$code = "<?php\n\n\$var = 1;";
$tokens = token_get_all($code);
$varToken = null;
foreach ($tokens as $t) {
    if (is_array($t) && $t[0] === T_VARIABLE) { $varToken = $t; break; }
}
echo $varToken[2] === 3 ? "LINE_3_OK" : "FAIL";
