<?php
// vybe-test: php/functions/nested_functions
// origin: languages/php/tests/php/test_functions.rs
// vybe-test-mode: compile

function outer() { function inner() { return 42; } return inner(); } echo outer();
