<?php
// vybe-test: php/php84/array_all_false
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

$mixed = [1, 5, -3, 8];
echo array_all($mixed, fn($n) => $n > 0) ? 'all positive' : 'some negative';
