<?php
// vybe-test: php/php7/variadic
// origin: languages/php/tests/php/test_php7.rs
// vybe-test-mode: compile

function sum(int ...$nums): int { return 0; } echo sum(1,2,3);
