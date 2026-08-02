<?php
// vybe-test: php/error_handling_deep/arithmetic_error
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

try { $r = intdiv(1, 0); }
catch (\DivisionByZeroError $e) { echo 'div by zero'; }
