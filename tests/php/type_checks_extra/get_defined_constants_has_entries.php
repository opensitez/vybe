<?php
// vybe-test: php/type_checks_extra/get_defined_constants_has_entries
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

define('MY_CONST', 99);
$consts = get_defined_constants(true);
echo isset($consts['user']['MY_CONST']) ? 'yes' : 'no';
