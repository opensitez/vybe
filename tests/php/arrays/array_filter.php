<?php
// vybe-test: php/arrays/array_filter
// origin: languages/php/tests/php/test_arrays.rs
// vybe-test-mode: compile

$x = array_filter([0, 1, '', 'a', null], fn($v) => $v);
