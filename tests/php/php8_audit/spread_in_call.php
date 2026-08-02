<?php
// vybe-test: php/php8_audit/spread_in_call
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

function sum(...$nums) { return 0; } sum(...[1,2,3]);
