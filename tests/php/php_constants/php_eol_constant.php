<?php
// vybe-test: php/php_constants/php_eol_constant
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$line = 'hello' . PHP_EOL;
echo strlen($line) > 5 ? 'has_eol' : 'no_eol';
