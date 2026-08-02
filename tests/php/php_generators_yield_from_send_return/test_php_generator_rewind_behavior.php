<?php
// vybe-test: php/php_generators_yield_from_send_return/test_php_generator_rewind_behavior
// origin: languages/php/tests/php/test_php_generators_yield_from_send_return.rs
// vybe-test-mode: compile

function gen() {
    yield 1;
    yield 2;
}

$g = gen();
$g->rewind();
echo $g->current();
$g->next();
echo $g->current();
