<?php
// vybe-test: php/closures/arrow_auto_capture
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$x = 10; $fn = fn($y) => $x + $y; echo $fn(5);
