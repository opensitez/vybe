<?php
// vybe-test: php/closures/closure_in_map
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

array_map(function($x) { return $x * $x; }, [1,2,3]);
