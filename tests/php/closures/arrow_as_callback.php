<?php
// vybe-test: php/closures/arrow_as_callback
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

$doubled = array_map(fn($n) => $n * 2, [1, 2, 3]);
