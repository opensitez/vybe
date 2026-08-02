<?php
// vybe-test: php/scoping/swap_vars
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

$a = 1; $b = 2; $tmp = $a; $a = $b; $b = $tmp; echo $a . $b;
