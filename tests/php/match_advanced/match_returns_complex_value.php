<?php
// vybe-test: php/match_advanced/match_returns_complex_value
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

$key = 'users';
$config = match($key) {
    'users'  => ['table' => 'users',  'pk' => 'id'],
    'orders' => ['table' => 'orders', 'pk' => 'order_id'],
    default  => ['table' => 'unknown', 'pk' => 'id'],
};
echo $config['table'] . ':' . $config['pk'];
