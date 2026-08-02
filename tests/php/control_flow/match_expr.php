<?php
// vybe-test: php/control_flow/match_expr
// origin: languages/php/tests/php/test_control_flow.rs
// vybe-test-mode: compile

$x = match($v) { 1 => 'one', 2 => 'two', default => 'other' };
