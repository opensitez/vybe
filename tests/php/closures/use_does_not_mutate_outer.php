<?php
// vybe-test: php/closures/use_does_not_mutate_outer
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$x = 'original'; $fn = function() use ($x) { $x = 'modified'; return $x; }; echo $fn(); echo $x;
