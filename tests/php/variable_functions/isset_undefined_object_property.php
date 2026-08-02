<?php
// vybe-test: php/variable_functions/isset_undefined_object_property
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

class Config {
    public string $host = 'localhost';
}
$c = new Config();
echo isset($c->host)    ? 'set' : 'unset';
echo isset($c->missing) ? 'set' : 'unset';
