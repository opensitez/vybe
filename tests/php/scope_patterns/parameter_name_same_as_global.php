<?php
// vybe-test: php/scope_patterns/parameter_name_same_as_global
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$x = 'global';
function echo_param(string $x): void {
    echo $x;
}
echo_param('local');
echo $x;
