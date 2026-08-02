<?php
// vybe-test: php/php_reflection_fiber_generator_inspection/test_php_reflection_generator_closed_generator_error
// origin: languages/php/tests/php/test_php_reflection_fiber_generator_inspection.rs
// vybe-test-mode: compile

function simpleGen() { return 42; }
$g = simpleGen();
foreach ($g as $v) {} // exhaust generator

try {
    $rg = new ReflectionGenerator($g);
} catch (Error $e) {
    echo "Closed generator reflection error caught";
}
