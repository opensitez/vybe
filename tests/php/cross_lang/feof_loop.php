<?php
// vybe-test: php/cross_lang/feof_loop
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

$fp = fopen('data.csv', 'r');
while (!feof($fp)) {
    $line = fgets($fp);
    echo $line;
}
fclose($fp);
