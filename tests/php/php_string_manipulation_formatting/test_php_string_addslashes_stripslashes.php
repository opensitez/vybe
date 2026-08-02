<?php
// vybe-test: php/php_string_manipulation_formatting/test_php_string_addslashes_stripslashes
// origin: languages/php/tests/php/test_php_string_manipulation_formatting.rs
// vybe-test-mode: compile

$str = "Is your name O'Reilly?";
$escaped = addslashes($str);
echo stripslashes($escaped);
