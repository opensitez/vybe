<?php
// vybe-test: php/scoping/closure_by_reference_capture
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

$value = 1; $inc = function() use (&$value) { $value++; }; $inc(); $inc(); echo $value;
