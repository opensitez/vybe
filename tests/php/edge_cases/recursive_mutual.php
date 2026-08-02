<?php
// vybe-test: php/edge_cases/recursive_mutual
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

function isEven($n) { if ($n == 0) return true; return isOdd($n - 1); } function isOdd($n) { if ($n == 0) return false; return isEven($n - 1); } echo isEven(4);
