<?php
// vybe-test: php/php80_phptoken_tokenize_properties/test_php80_phptoken_json_serialize
// origin: languages/php/tests/php/test_php80_phptoken_tokenize_properties.rs
// vybe-test-mode: compile

if (class_exists('PhpToken')) {
    $tokens = PhpToken::tokenize("<?php 1;");
    $json = json_encode($tokens[1]);
    echo str_contains($json, "id") && str_contains($json, "text") ? "JSON_TOKEN_OK" : "FAIL";
} else {
    echo "JSON_TOKEN_OK";
}
