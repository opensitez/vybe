<?php
// vybe-test: php/class_inspection/get_class_vars_default_properties
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class Config {
    public string $host = 'localhost';
    public int $port = 8080;
}
$vars = get_class_vars('Config');
echo isset($vars['host']) ? 'yes' : 'no';
echo isset($vars['port']) ? 'yes' : 'no';
