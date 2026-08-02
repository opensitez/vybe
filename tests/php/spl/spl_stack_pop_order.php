<?php
// vybe-test: php/spl/spl_stack_pop_order
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$stack = new SplStack();
foreach ([10, 20, 30] as $v) { $stack->push($v); }
$result = [];
while (!$stack->isEmpty()) { $result[] = $stack->pop(); }
echo implode(',', $result);
