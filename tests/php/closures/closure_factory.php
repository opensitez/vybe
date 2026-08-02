<?php
// vybe-test: php/closures/closure_factory
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

function multiplier($factor) { return function($x) use ($factor) { return $x * $factor; }; } $double = multiplier(2); echo $double(5);
