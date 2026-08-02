<?php
// vybe-test: php/match_advanced/match_null_value
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$v = null;
$result = match($v) {
    null  => 'null',
    false => 'false',
    0     => 'zero',
    ''    => 'empty string',
    default => 'something',
};
echo $result;
