<?php
// vybe-test: php/file_functions/file_put_contents_basic
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$bytes = file_put_contents('/tmp/test_vybe.txt', 'hello world');
echo $bytes !== false ? 'wrote' : 'failed';
