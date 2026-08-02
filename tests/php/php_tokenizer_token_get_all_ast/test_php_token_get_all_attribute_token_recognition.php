<?php
// vybe-test: php/php_tokenizer_token_get_all_ast/test_php_token_get_all_attribute_token_recognition
// origin: languages/php/tests/php/test_php_tokenizer_token_get_all_ast.rs
// vybe-test-mode: compile

$code = '<?php #[Attribute] class Test {}';
$tokens = token_get_all($code, TOKEN_PARSE);
$hasAttr = false;
foreach ($tokens as $t) {
    if (is_array($t) && defined('T_ATTRIBUTE') && $t[0] === T_ATTRIBUTE) {
        $hasAttr = true;
    }
}
echo $hasAttr ? "HAS_ATTR_TOKEN" : "LEGACY_TOKENS";
