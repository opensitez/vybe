<?php
// vybe-test: php/declare/strict_types_bool_param
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function toggle(bool $flag): bool { return !$flag; }
var_dump(toggle(true));
