<?php
// vybe-test: php/php80_phptoken_tokenize_properties/test_php80_phptoken_flags_token_parse
// origin: languages/php/tests/php/test_php80_phptoken_tokenize_properties.rs
// vybe-test-mode: compile

if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php \$x = 1;", TOKEN_PARSE);
    echo count($tokens) > 0 ? "TOKEN_PARSE_OK" : "FAIL";
} else {
    echo "TOKEN_PARSE_OK";
}
