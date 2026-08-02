<?php
// vybe-test: php/type_functions_extended/is_object_check
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

class Point { public int $x; public int $y; }
$p = new Point();
echo is_object($p) ? 'yes' : 'no';
echo is_object([]) ? 'yes' : 'no';
