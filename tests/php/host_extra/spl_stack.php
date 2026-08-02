<?php
// vybe-test: php/host_extra/spl_stack
// origin: languages/php/tests/php/test_host_extra.rs
// vybe-test-mode: compile

$stack = new SplStack();
$stack->push('a');
$stack->push('b');
$stack->push('c');
$top = $stack->pop();
echo $top;
