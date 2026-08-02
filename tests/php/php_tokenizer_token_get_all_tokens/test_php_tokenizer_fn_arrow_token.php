<?php
// vybe-test: php/php_tokenizer_token_get_all_tokens/test_php_tokenizer_fn_arrow_token
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_tokens.rs
// vybe-test-mode: compile

$code = "<?php \$f = fn() => 42;";
$tokens = token_get_all($code);
$hasFn = false;
foreach ($tokens as $t) {
    if (is_array($t) && (defined('T_FN') && $t[0] === T_FN)) { $hasFn = true; break; }
}
echo "FN_TOKEN_CHECKED";
