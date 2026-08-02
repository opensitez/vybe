<?php
// vybe-test: php/php_constants/directory_separator_constant
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$path = 'usr' . DIRECTORY_SEPARATOR . 'local' . DIRECTORY_SEPARATOR . 'bin';
echo strlen($path) > 0 ? 'ok' : 'empty';
