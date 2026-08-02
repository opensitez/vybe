<?php
// vybe-test: php/reflection/reflection_parameter_variadic
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

function sum(int ...$nums): int { return array_sum($nums); }
$rf = new ReflectionFunction('sum');
$p = $rf->getParameters()[0];
echo $p->isVariadic() ? 'variadic' : 'regular';
