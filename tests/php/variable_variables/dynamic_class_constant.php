<?php
// vybe-test: php/variable_variables/dynamic_class_constant
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Status {
    const OK    = 200;
    const ERROR = 500;
}
$const = 'OK';
echo constant("Status::$const");
