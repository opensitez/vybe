<?php
// vybe-test: php/php_generators_yield_from_send_return/test_php_generator_throw_exception_into_generator
// origin: languages/php/tests/php/test_php_generators_yield_from_send_return.rs
// vybe-test-mode: compile

function exceptionGen() {
    try {
        yield "start";
    } catch (RuntimeException $e) {
        yield "handled: " . $e->getMessage();
    }
}

$g = exceptionGen();
echo $g->current();
echo $g->throw(new RuntimeException("Injected Error"));
