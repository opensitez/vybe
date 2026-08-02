<?php
// vybe-test: php/advanced_closures/call_user_func_with_closure
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$greet = function(string $name): string { return 'hi ' . $name; };
echo call_user_func($greet, 'world');
