<?php
// vybe-test: php/php_tokenizer_token_get_all_tokens/test_php_tokenizer_comment_tokens
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_tokens.rs
// vybe-test-mode: compile

$code = "<?php // Single line comment\n/* Block comment */";
$tokens = token_get_all($code);
$commentCount = 0;
foreach ($tokens as $t) {
    if (is_array($t) && ($t[0] === T_COMMENT || $t[0] === T_DOC_COMMENT)) { $commentCount++; }
}
echo $commentCount === 2 ? "COMMENT_TOKENS_OK" : "FAIL";
