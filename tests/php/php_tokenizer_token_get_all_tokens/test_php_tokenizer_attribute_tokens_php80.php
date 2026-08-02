<?php
// vybe-test: php/php_tokenizer_token_get_all_tokens/test_php_tokenizer_attribute_tokens_php80
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_tokens.rs
// vybe-test-mode: compile

$code = "<?php #[Attribute]";
$tokens = token_get_all($code);
$hasAttr = false;
foreach ($tokens as $t) {
    if (is_array($t) && (defined('T_ATTRIBUTE') && $t[0] === T_ATTRIBUTE)) { $hasAttr = true; break; }
}
echo "ATTRIBUTE_TOKEN_CHECKED";
