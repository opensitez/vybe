<?php
// vybe-test: php/builtins/array_walk
// origin: languages/php/tests/php/test_builtins.rs
// vybe-test-mode: compile

$a = [1,2,3]; array_walk($a, fn($v,$k) => $v);
