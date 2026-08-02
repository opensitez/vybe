<?php
// vybe-test: php/php_string_manipulation_formatting/test_php_string_wordwrap_breaking
// origin: languages/php/tests/php/test_php_string_manipulation_formatting.rs
// vybe-test-mode: compile

$text = "A very long words sentence that needs wrapping";
$newtext = wordwrap($text, 10, "\n", true);
echo $newtext;
