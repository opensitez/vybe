<?php
// vybe-test: php/php5_legacy/callable_hint
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

function apply(callable $fn, $val) { return $fn($val); }
