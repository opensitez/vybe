<?php
// vybe-test: php/pcre_advanced/preg_quote_special_chars
// origin: languages/php/tests/php/test_pcre_advanced.rs
// vybe-test-mode: compile

$user_input = 'price is $10.00 (USD) + tax?';
$escaped = preg_quote($user_input, '/');
echo preg_match('/' . $escaped . '/', $user_input) ? 'found' : 'not found';
$special = '\.+*?[^]$(){}=!<>|:-#';
$q = preg_quote($special);
echo is_string($q) ? 'quoted' : 'fail';
