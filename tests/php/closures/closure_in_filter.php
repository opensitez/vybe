<?php
// vybe-test: php/closures/closure_in_filter
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

array_filter([1,2,3,4,5], function($x) { return $x > 2; });
