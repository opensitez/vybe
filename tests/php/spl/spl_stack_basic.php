<?php
// vybe-test: php/spl/spl_stack_basic
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$stack = new SplStack();
$stack->push(1);
$stack->push(2);
$stack->push(3);
echo $stack->top();
echo $stack->count();
