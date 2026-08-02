<?php
// vybe-test: php/php84/array_any_false
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

$numbers = [1, 3, 5, 7, 9];
echo array_any($numbers, fn($n) => $n % 2 === 0) ? 'has even' : 'all odd';
