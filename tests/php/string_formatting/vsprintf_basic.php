<?php
// vybe-test: php/string_formatting/vsprintf_basic
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

$args = ['PHP', '8.3'];
$result = vsprintf("%s version %s", $args);
echo $result;
echo "\n";
