<?php
// vybe-test: php/closures/closure_in_reduce
// origin: languages/php/tests/php/test_closures.rs
// vybe-test-mode: compile

array_reduce([1,2,3], function($carry, $item) { return $carry + $item; }, 0);
