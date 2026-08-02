<?php
// vybe-test: php/php_reflection_fiber_generator_inspection/test_php_reflection_generator_get_executing_generator
// origin: languages/php/tests/php/test_php_reflection_fiber_generator_inspection.rs
// vybe-test-mode: compile

function innerGen() { yield 100; }
function outerGen() { yield from innerGen(); }

$g = outerGen();
$g->current();
$rg = new ReflectionGenerator($g);
$execGen = $rg->getExecutingGenerator();
echo $execGen instanceof Generator ? "EXEC_GEN_OK" : "FAIL";
