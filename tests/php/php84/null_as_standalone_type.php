<?php
// vybe-test: php/php84/null_as_standalone_type
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

function alwaysNull(): null { return null; }
$v = alwaysNull();
var_dump($v);
