<?php
// vybe-test: php/spl_extra/spl_stack_is_empty_check
// origin: languages/php/tests/php/test_spl_extra.rs
// vybe-test-mode: compile

$s = new SplStack();
echo $s->isEmpty() ? 'empty' : 'not empty';
$s->push(42);
echo $s->isEmpty() ? 'empty' : 'not empty';
$s->pop();
echo $s->isEmpty() ? 'empty' : 'not empty';
