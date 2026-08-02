<?php
// vybe-test: php/variable_functions/isset_chained_array_key
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

$config = ['db' => ['host' => 'localhost', 'port' => 3306]];
echo isset($config['db']['host']) ? 'set' : 'missing';
echo isset($config['db']['password']) ? 'set' : 'missing';
echo isset($config['cache']['host']) ? 'set' : 'missing';
