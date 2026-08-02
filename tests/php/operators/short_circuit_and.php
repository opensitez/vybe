<?php
// vybe-test: php/operators/short_circuit_and
// origin: languages/php/tests/php/test_operators.rs
// vybe-test-mode: compile

$x = false && expensive();
