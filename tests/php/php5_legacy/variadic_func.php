<?php
// vybe-test: php/php5_legacy/variadic_func
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

function sum(...$nums) { return array_sum($nums); } echo sum(1, 2, 3);
