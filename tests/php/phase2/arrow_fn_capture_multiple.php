<?php
// vybe-test: php/phase2/arrow_fn_capture_multiple
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

$base = 100;
$tax = 0.2;
$calc = fn($price) => ($price + $base) * (1 + $tax);
echo $calc(50);
