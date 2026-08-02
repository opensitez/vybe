<?php
// vybe-test: php/php_tokenizer_token_get_all_ast/test_php_token_get_all_enum_token_php81
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_ast.rs
// vybe-test-mode: compile

$code = '<?php enum Status { case Active; }';
$tokens = token_get_all($code);
$hasEnum = false;
foreach ($tokens as $t) {
    if (is_array($t) && defined('T_ENUM') && $t[0] === T_ENUM) {
        $hasEnum = true;
    }
}
echo $hasEnum ? "ENUM_TOKEN_OK" : "NO_ENUM_TOKEN";
