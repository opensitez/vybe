<?php
// vybe-test: php/php5_legacy/arg_unpack
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

function add($a, $b) { return $a + $b; } echo add(...[3, 4]);
