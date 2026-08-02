<?php
// vybe-test: php/edge_cases/assign_in_while
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

$i = 0; while (($i = $i + 1) < 10) { echo $i; }
