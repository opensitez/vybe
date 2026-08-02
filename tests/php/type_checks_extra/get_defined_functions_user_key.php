<?php
// vybe-test: php/type_checks_extra/get_defined_functions_user_key
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

function myFunc() { return 1; }
$fns = get_defined_functions();
echo isset($fns['user']) ? 'yes' : 'no';
