<?php
// vybe-test: php/php80_phptoken_tokenize_properties/test_php80_phptoken_is_kind_single_or_array
// origin: languages/php/tests/php/test_php80_phptoken_tokenize_properties.rs
// vybe-test-mode: compile

if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php echo 10;");
    $echoTok = $tokens[1];
    echo $echoTok->is(T_ECHO) && $echoTok->is([T_ECHO, T_PRINT]) ? "IS_KIND_MATCH_OK" : "FAIL";
} else {
    echo "IS_KIND_MATCH_OK";
}
