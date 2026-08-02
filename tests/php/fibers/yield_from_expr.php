<?php
// vybe-test: php/fibers/yield_from_expr
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

function inner() {
    yield 1;
    yield 2;
}
function outer() {
    yield from inner();
    yield 3;
}
