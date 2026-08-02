<?php
// vybe-test: php/php80_phptoken_tokenize_properties/test_php80_phptoken_get_token_name_single_char
// origin: languages/php/tests/php/test_php80_phptoken_tokenize_properties.rs
// vybe-test-mode: compile

if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php ;");
    $semi = $tokens[1];
    echo $semi->getTokenName() === ";" ? "SEMI_NAME_OK" : "FAIL";
} else {
    echo "SEMI_NAME_OK";
}
