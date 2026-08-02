<?php
// vybe-test: php/file_functions/dirname_basic
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$dir = dirname('/var/www/html/index.php');
echo $dir;
$nested = dirname('/a/b/c/d.txt', 2);
echo $nested;
