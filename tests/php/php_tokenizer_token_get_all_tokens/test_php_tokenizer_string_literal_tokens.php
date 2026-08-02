<?php
// vybe-test: php/php_tokenizer_token_get_all_tokens/test_php_tokenizer_string_literal_tokens
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_tokens.rs
// vybe-test-mode: compile

$tokens = token_get_all("<?php 'hello';");
$hasConstantEncapsed = false;
foreach ($tokens as $t) {
    if (is_array($t) && $t[0] === T_CONSTANT_ENCAPSED_STRING) { $hasConstantEncapsed = true; break; }
}
echo $hasConstantEncapsed ? "T_CONSTANT_ENCAPSED_OK" : "FAIL";
