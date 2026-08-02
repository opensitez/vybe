<?php
// vybe-test: php/magic_constants/magic_file_in_function
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

function getFile(): string { return __FILE__; }
echo is_string(getFile()) ? 'ok' : 'fail';
