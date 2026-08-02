<?php
// vybe-test: php/php_spl_stack_lifo_push_pop/test_php_spl_stack_push_multiple_types
// origin: languages/php/tests/php/test_php_spl_stack_lifo_push_pop.rs
// vybe-test-mode: compile

$s = new SplStack();
$s->push(123);
$s->push(["key" => "value"]);
$s->push(new stdClass());
echo count($s) === 3 ? "PUSH_MIXED_OK" : "FAIL";
