<?php
// vybe-test: php/file_functions/pathinfo_components
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$info = pathinfo('/var/www/html/index.php');
echo $info['dirname'];
echo $info['basename'];
echo $info['extension'];
echo $info['filename'];
echo pathinfo('/var/www/index.php', PATHINFO_EXTENSION);
