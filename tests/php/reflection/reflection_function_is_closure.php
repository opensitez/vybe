<?php
// vybe-test: php/reflection/reflection_function_is_closure
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

$f = function() {};
$rf = new ReflectionFunction($f);
echo $rf->isClosure() ? 'closure' : 'function';
