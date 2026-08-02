<?php
// vybe-test: php/php5_legacy/common_string_ops
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

$email = '  USER@EXAMPLE.COM  ';
$clean = strtolower(trim($email));
$parts = explode('@', $clean);
$user = $parts[0];
$domain = $parts[1];
echo $user . ' at ' . $domain;
