<?php
// vybe-test: php/cross_lang/random_obj
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

$rng = new Random();
$n = $rng->nextInt(1, 100);
$f = $rng->nextFloat();
