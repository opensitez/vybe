<?php
// vybe-test: php/php_tokenizer_token_get_all_ast/test_php_tokenizer_match_token_php80
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_ast.rs
// vybe-test-mode: compile

$code = '<?php $x = match($a) { 1 => "one" };';
$tokens = token_get_all($code);
$hasMatch = false;
foreach ($tokens as $t) {
    if (is_array($t) && defined('T_MATCH') && $t[0] === T_MATCH) {
        $hasMatch = true;
    }
}
echo $hasMatch ? "MATCH_TOKEN_OK" : "NO_MATCH_TOKEN";
