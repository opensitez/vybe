<?php
// vybe-test: php/php7/argument_unpack
// origin: languages/php/tests/php/test_php7.rs
// vybe-test-mode: compile

function add($a, $b) { return $a + $b; } echo add(...[1, 2]);
