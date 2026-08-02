<?php
// vybe-test: php/closures/arrow_nested_capture
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$a = 1; $outer = fn($b) => fn($c) => $a + $b + $c;
