<?php
// vybe-test: php/scoping/closure_use_scope
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

$x = 'hello'; $fn = function() use ($x) { return $x; }; $x = 'changed'; echo $fn();
