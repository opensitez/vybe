<?php
// vybe-test: php/scoping/arrow_fn_scope
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

$x = 5; $fn = fn() => $x; echo $fn();
