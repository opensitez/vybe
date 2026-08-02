<?php
// vybe-test: php/php_generators_yield_from_send_return/test_php_generator_closed_state_inspection
// origin: languages/php/tests/php/test_php_generators_yield_from_send_return.rs
// vybe-test-mode: compile

function simpleGen() {
    yield 1;
}

$g = simpleGen();
foreach ($g as $v) {}
echo $g->valid() ? "VALID" : "EXHAUSTED";
