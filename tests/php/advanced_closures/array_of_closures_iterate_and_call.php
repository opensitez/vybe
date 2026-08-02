<?php
// vybe-test: php/advanced_closures/array_of_closures_iterate_and_call
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$fns = [
    fn(int $x): int => $x + 1,
    fn(int $x): int => $x * 2,
    fn(int $x): int => $x - 3,
];
$val = 10;
foreach ($fns as $fn) {
    $val = $fn($val);
}
echo $val;
