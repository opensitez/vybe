<?php
// vybe-test: php/phase2/arrow_fn_nested_capture
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

$x = 10;
$outer = fn($a) => fn($b) => $a + $b + $x;
