<?php
// vybe-test: php/type_checks_extra/get_defined_vars_local_scope
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

function checkVars() {
    $x = 10;
    $y = 20;
    $vars = get_defined_vars();
    echo isset($vars['x']) ? 'yes' : 'no';
    echo isset($vars['y']) ? 'yes' : 'no';
}
checkVars();
