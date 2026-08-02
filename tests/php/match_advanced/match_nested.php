<?php
// vybe-test: php/match_advanced/match_nested
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$type = 'user';
$role = 'admin';
$result = match($type) {
    'user' => match($role) {
        'admin'  => 'admin user',
        'editor' => 'editor user',
        default  => 'regular user',
    },
    'bot' => 'bot',
    default => 'unknown',
};
echo $result;
