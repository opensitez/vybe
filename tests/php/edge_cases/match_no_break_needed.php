<?php
// vybe-test: php/edge_cases/match_no_break_needed
// origin: languages/php/tests/php/test_edge_cases.rs
// vybe-test-mode: compile

$x = match(2) { 1 => 'one', 2 => 'two', 3 => 'three', default => '?' };
