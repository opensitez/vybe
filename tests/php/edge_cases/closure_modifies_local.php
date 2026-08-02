<?php
// vybe-test: php/edge_cases/closure_modifies_local
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

$arr = []; $fn = function($v) use ($arr) { array_push($arr, $v); return $arr; }; $fn(1);
