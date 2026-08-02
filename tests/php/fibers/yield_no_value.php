<?php
// vybe-test: php/fibers/yield_no_value
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

function signals() {
    yield;
    yield;
}
