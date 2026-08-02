<?php
// vybe-test: php/spl_extra/spl_stack_lifo_iteration
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$s = new SplStack();
$s->push(1); $s->push(2); $s->push(3);
$result = [];
foreach ($s as $v) { $result[] = $v; }
echo implode(',', $result); // 3,2,1
