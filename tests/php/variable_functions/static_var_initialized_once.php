<?php
// vybe-test: php/variable_functions/static_var_initialized_once
// origin: languages/php/tests/php/test_variable_functions.rs
// vybe-test-mode: compile

function makeId(): string {
    static $id = 'ID-000';
    return $id;
}
echo makeId();
echo makeId();
