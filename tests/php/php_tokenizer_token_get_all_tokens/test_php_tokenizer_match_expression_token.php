<?php
// vybe-test: php/php_tokenizer_token_get_all_tokens/test_php_tokenizer_match_expression_token
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_tokens.rs
// vybe-test-mode: compile

$code = "<?php match(\$x) { default => 0 };";
$tokens = token_get_all($code);
$hasMatch = false;
foreach ($tokens as $t) {
    if (is_array($t) && (defined('T_MATCH') && $t[0] === T_MATCH)) { $hasMatch = true; break; }
}
echo "MATCH_TOKEN_CHECKED";
