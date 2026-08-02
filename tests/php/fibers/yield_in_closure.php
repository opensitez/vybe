<?php
// vybe-test: php/fibers/yield_in_closure
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

$gen = function() {
    yield 1;
    yield 2;
};
