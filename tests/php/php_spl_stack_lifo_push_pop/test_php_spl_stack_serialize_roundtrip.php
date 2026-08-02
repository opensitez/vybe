<?php
// vybe-test: php/php_spl_stack_lifo_push_pop/test_php_spl_stack_serialize_roundtrip
// origin: languages/php/tests/php/test_php_spl_stack_lifo_push_pop.rs
// vybe-test-mode: compile

$s = new SplStack();
$s->push("data1");
$s->push("data2");
$serialized = serialize($s);
$restored = unserialize($serialized);
echo $restored->pop() === "data2" ? "RESTORED_LIFO_OK" : "FAIL";
