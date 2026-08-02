<?php
// vybe-test: php/php_spl_stack_lifo_push_pop/test_php_spl_stack_default_iterator_mode_is_lifo_keep
// origin: languages/php/tests/php/test_php_spl_stack_lifo_push_pop.rs
// vybe-test-mode: compile

$stack = new SplStack();
$stack->push("a");
$stack->push("b");
$mode = $stack->getIteratorMode();
echo ($mode & SplDoublyLinkedList::IT_MODE_LIFO) ? "MODE_LIFO" : "FAIL";
