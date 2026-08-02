<?php
// vybe-test: php/scoping/nested_closure_capture_from_arrow_function
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

$factor = 2; $double = fn(int $n) => $n * $factor; $twice = function(int $n) use ($double): int { return $double($n); }; echo $twice(4);
