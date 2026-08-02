<?php
// vybe-test: php/spl_extra/spl_stack_push_pop_top
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$s = new SplStack();
$s->push('first');
$s->push('second');
$s->push('third');
echo $s->top();
echo $s->pop();
echo $s->count();
