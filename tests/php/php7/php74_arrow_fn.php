<?php
// vybe-test: php/php7/php74_arrow_fn
// origin: languages/php/tests/php/test_php7.rs
// vybe-test-mode: compile

$fn = fn($x) => $x * 2; echo $fn(5);
