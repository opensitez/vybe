<?php
// vybe-test: php/php84/array_all_true
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

$positives = [1, 5, 8, 3, 12];
echo array_all($positives, fn($n) => $n > 0) ? 'all positive' : 'some negative';
