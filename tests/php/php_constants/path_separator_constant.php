<?php
// vybe-test: php/php_constants/path_separator_constant
// origin: languages/php/tests/php/test_php_constants.rs
// vybe-test-mode: compile

$env = '/usr/bin' . PATH_SEPARATOR . '/usr/local/bin';
echo strlen($env) > 0 ? 'ok' : 'empty';
