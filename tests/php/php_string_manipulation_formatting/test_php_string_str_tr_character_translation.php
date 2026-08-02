<?php
// vybe-test: php/php_string_manipulation_formatting/test_php_string_str_tr_character_translation
// origin: languages/php/tests/php/test_php_string_manipulation_formatting.rs
// vybe-test-mode: compile

$trans = ["h" => "hello", "hello" => "hi"];
echo strtr("hi all, I said hello", $trans);
