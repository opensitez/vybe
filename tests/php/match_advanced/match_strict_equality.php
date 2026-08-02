<?php
// vybe-test: php/match_advanced/match_strict_equality
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$val = "1";
// match uses === not ==
$result = match(true) {
    $val === 1   => 'int 1',
    $val === "1" => 'string 1',
    default      => 'other',
};
echo $result;
