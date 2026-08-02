<?php
// vybe-test: php/spl/spl_stack_is_empty
// origin: languages/php/tests/php/test_spl.rs
// vybe-test-mode: compile

$s = new SplStack();
echo $s->isEmpty() ? 'empty' : 'not empty';
$s->push('a');
echo $s->isEmpty() ? 'empty' : 'not empty';
