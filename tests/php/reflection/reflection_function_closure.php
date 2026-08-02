<?php
// vybe-test: php/reflection/reflection_function_closure
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

$fn = fn(int $a, int $b) => $a + $b;
$rf = new ReflectionFunction($fn);
echo $rf->getNumberOfParameters();
echo $rf->isClosure() ? ':closure' : ':function';
