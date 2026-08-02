<?php
// vybe-test: php/php_spl_stack_lifo_push_pop/test_php_spl_stack_array_access_indexing
// origin: languages/php/tests/php/test_php_spl_stack_lifo_push_pop.rs
// vybe-test-mode: compile

$s = new SplStack();
$s->push("first_pushed");
$s->push("second_pushed");
// Index 0 in LIFO stack corresponds to top (second_pushed)
echo $s[0] === "second_pushed" ? "LIFO_INDEX0_OK" : "FAIL";
