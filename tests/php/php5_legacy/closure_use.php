<?php
// vybe-test: php/php5_legacy/closure_use
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

$prefix = 'Mr.'; $fn = function($name) use ($prefix) { return $prefix . ' ' . $name; }; echo $fn('Smith');
