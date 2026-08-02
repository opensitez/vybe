<?php
// vybe-test: php/php5_legacy/closure_basic
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

$greet = function($name) { return 'Hello ' . $name; }; echo $greet('World');
