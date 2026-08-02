<?php
// vybe-test: php/binary_data/bitwise_flags_pattern
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

const READ    = 1;
const WRITE   = 2;
const EXECUTE = 4;
$perms = READ | WRITE;
echo ($perms & READ)    ? 'can read '    : '';
echo ($perms & WRITE)   ? 'can write '   : '';
echo ($perms & EXECUTE) ? 'can execute ' : 'no execute';
