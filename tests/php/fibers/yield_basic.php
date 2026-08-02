<?php
// vybe-test: php/fibers/yield_basic
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

function gen() {
    yield 1;
    yield 2;
    yield 3;
}
