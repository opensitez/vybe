<?php
// vybe-test: php/file_functions/basename_basic
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

echo basename('/var/www/html/index.php');
echo basename('/var/www/html/index.php', '.php');
