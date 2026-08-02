<?php
// vybe-test: php/arrays/spread_in_call
// origin: languages/php/tests/php/test_arrays.rs
// vybe-test-mode: compile

function sum(...$nums) { return 0; } sum(...[1,2,3]);
