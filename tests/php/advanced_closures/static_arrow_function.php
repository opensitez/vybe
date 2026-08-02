<?php
// vybe-test: php/advanced_closures/static_arrow_function
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

$double = static fn(int $x): int => $x * 2;
echo $double(21);
