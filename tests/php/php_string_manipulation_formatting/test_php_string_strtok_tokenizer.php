<?php
// vybe-test: php/php_string_manipulation_formatting/test_php_string_strtok_tokenizer
// origin: languages/php/tests/php/test_php_string_manipulation_formatting.rs
// vybe-test-mode: compile

$string = "This is\tan example\nstring";
$tok = strtok($string, " \n\t");
while ($tok !== false) {
    echo "Word=$tok\n";
    $tok = strtok(" \n\t");
}
