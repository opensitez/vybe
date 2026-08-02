<?php
// vybe-test: php/arrays/array_push
// origin: languages/php/tests/php/test_arrays.rs
// vybe-test-mode: compile

$a = [1]; array_push($a, 2); echo count($a);
