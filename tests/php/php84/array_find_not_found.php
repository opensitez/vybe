<?php
// vybe-test: php/php84/array_find_not_found
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

$numbers = [1, 3, 5, 7, 9];
$even = array_find($numbers, fn($n) => $n % 2 === 0);
echo $even === null ? 'not found' : 'found';
