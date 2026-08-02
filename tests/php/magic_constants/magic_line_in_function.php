<?php
// vybe-test: php/magic_constants/magic_line_in_function
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

function getLine(): int { return __LINE__; }
$l = getLine();
echo $l > 0 ? 'ok' : 'fail';
