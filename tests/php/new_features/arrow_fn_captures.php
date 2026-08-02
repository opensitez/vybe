<?php
// vybe-test: php/new_features/arrow_fn_captures
// origin: languages/php/tests/php/test_new_features.rs
// vybe-test-mode: compile

$x = 5; $fn = fn($y) => $x + $y; echo $fn(3);
