<?php
// vybe-test: php/closures/arrow_factory
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

function adder($n) { return fn($x) => $x + $n; } $add5 = adder(5); echo $add5(10);
