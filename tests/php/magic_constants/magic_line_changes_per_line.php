<?php
// vybe-test: php/magic_constants/magic_line_changes_per_line
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

$a = __LINE__;
$b = __LINE__;
$c = __LINE__;
echo ($b > $a && $c > $b) ? 'increasing' : 'fail';
