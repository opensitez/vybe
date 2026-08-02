<?php
// vybe-test: php/php_spl_stack_lifo_push_pop/test_php_spl_stack_empty_check_with_is_empty
// origin: languages/php/tests/php/test_php_spl_stack_lifo_push_pop.rs
// vybe-test-mode: compile

$s = new SplStack();
echo $s->isEmpty() ? "EMPTY" : "NOT_EMPTY";
