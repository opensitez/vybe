<?php
// vybe-test: php/edge_cases/variadic
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

function sum(...$nums) { return array_sum($nums); } echo sum(1,2,3,4,5);
