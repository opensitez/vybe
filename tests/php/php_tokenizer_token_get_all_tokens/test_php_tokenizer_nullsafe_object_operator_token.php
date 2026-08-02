<?php
// vybe-test: php/php_tokenizer_token_get_all_tokens/test_php_tokenizer_nullsafe_object_operator_token
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_tokens.rs
// vybe-test-mode: compile

$code = "<?php \$x?->prop;";
$tokens = token_get_all($code);
$hasNullsafe = false;
foreach ($tokens as $t) {
    if (is_array($t) && (defined('T_NULLSAFE_OBJECT_OPERATOR') && $t[0] === T_NULLSAFE_OBJECT_OPERATOR)) { $hasNullsafe = true; break; }
}
echo "NULLSAFE_OPERATOR_TOKEN_CHECKED";
