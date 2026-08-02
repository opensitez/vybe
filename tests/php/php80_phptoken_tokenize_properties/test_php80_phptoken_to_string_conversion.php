<?php
// vybe-test: php/php80_phptoken_tokenize_properties/test_php80_phptoken_to_string_conversion
// origin: languages/php/tests/php/test_php80_phptoken_tokenize_properties.rs
// vybe-test-mode: compile

if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php \$var");
    $varTok = $tokens[1];
    echo (string)$varTok === "\$var" ? "TO_STRING_OK" : "FAIL";
} else {
    echo "TO_STRING_OK";
}
