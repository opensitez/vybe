<?php
// vybe-test: php/functions/default_params
// origin: languages/php/tests/php/test_functions.rs
// vybe-test-mode: compile

function add($a, $b = 10) { return $a + $b; } echo add(5);
