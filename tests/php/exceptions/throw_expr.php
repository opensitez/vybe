<?php
// vybe-test: php/exceptions/throw_expr
// origin: languages/php/tests/php/test_exceptions.rs
// vybe-test-mode: compile

function fail() { throw new Exception('fail'); }
