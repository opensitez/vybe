<?php
// vybe-test: php/cross_lang/fopen_fgets
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

$fp = fopen('test.txt', 'r');
$line = fgets($fp);
fclose($fp);
