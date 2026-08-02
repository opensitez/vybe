<?php
// vybe-test: php/phase2/arrow_fn_auto_capture
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

$multiplier = 3;
$fn = fn($x) => $x * $multiplier;
echo $fn(5);
