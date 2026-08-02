<?php
// vybe-test: php/scope_patterns/closure_local_shadows_outer_by_value
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$name = 'outer';
$fn = function() use ($name): string {
    $name = 'inner';
    return $name;
};
echo $fn();
echo $name;
