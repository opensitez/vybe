<?php
// vybe-test: php/php_tokenizer_token_get_all_ast/test_php80_php_token_is_given_kind
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_ast.rs
// vybe-test-mode: compile

$tokens = PhpToken::tokenize('<?php function test() {}');
$fnTok = $tokens[1]; // T_FUNCTION or whitespace
echo $fnTok->is(T_FUNCTION) ? "IS_FUNCTION" : "NOT_FUNCTION";
