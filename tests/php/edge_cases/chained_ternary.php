<?php
// vybe-test: php/edge_cases/chained_ternary
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

$x = $a ? 'a' : ($b ? 'b' : 'c');
