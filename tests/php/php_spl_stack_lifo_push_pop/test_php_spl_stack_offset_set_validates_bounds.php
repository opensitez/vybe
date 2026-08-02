<?php
// vybe-test: php/php_spl_stack_lifo_push_pop/test_php_spl_stack_offset_set_validates_bounds
// origin: languages/php/tests/php/test_php_spl_stack_lifo_push_pop.rs
// vybe-test-mode: compile

$stack = new SplStack();
$stack->push("val1");
try {
    $stack[5] = "out_of_bounds";
} catch (OutOfRangeException $e) {
    echo "OutOfRangeException caught";
}
