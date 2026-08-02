<?php
// vybe-test: php/operators/short_circuit_or
// origin: languages/php/tests/php/test_operators.rs
// vybe-test-mode: compile

$x = true || expensive();
