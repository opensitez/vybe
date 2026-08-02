<?php
// vybe-test: php/php_spl_stack_lifo_push_pop/test_php_spl_stack_top_peek_without_popping
// origin: languages/php/tests/php/test_php_spl_stack_lifo_push_pop.rs
// vybe-test-mode: compile

$s = new SplStack();
$s->push("element");
echo $s->top() === "element" && count($s) === 1 ? "PEEK_OK" : "FAIL";
