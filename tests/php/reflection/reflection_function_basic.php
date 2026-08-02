<?php
// vybe-test: php/reflection/reflection_function_basic
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

function greet(string $name, string $prefix = 'Hello'): string {
    return "$prefix, $name!";
}
$rf = new ReflectionFunction('greet');
echo $rf->getName();
echo ':' . $rf->getNumberOfParameters();
