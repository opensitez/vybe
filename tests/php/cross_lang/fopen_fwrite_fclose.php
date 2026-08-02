<?php
// vybe-test: php/cross_lang/fopen_fwrite_fclose
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

$fp = fopen('test.txt', 'w');
fwrite($fp, 'Hello World');
fclose($fp);
