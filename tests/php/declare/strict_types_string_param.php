<?php
// vybe-test: php/declare/strict_types_string_param
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function shout(string $s): string { return strtoupper($s) . '!'; }
echo shout("hello");
