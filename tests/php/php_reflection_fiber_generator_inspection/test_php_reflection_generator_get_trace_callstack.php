<?php
// vybe-test: php/php_reflection_fiber_generator_inspection/test_php_reflection_generator_get_trace_callstack
// origin: languages/php/tests/php/test_php_reflection_fiber_generator_inspection.rs
// vybe-test-mode: compile

function levelB() { yield "B"; }
function levelA() { yield from levelB(); }

$gen = levelA();
$gen->current();

$rg = new ReflectionGenerator($gen);
$trace = $rg->getTrace();
echo is_array($trace) ? "TRACE_ARRAY_OK" : "FAIL";
