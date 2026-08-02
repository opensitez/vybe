<?php
// vybe-test: php/references/closure_reference_accumulator
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

$total = 0;
$add = function(int $n) use (&$total) { $total += $n; };
array_walk([1, 2, 3, 4, 5], $add);
echo $total;
