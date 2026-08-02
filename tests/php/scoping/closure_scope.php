<?php
// vybe-test: php/scoping/closure_scope
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

$x = 'outer'; $fn = function() { $x = 'inner'; return $x; }; echo $fn(); echo $x;
