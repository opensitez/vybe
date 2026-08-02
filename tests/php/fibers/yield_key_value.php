<?php
// vybe-test: php/fibers/yield_key_value
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

function pairs() {
    yield 'a';
    yield 'b';
    yield 'c';
}
