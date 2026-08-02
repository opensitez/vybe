<?php
// vybe-test: php/variable_variables/dynamic_property_get
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Point { public int $x = 1; public int $y = 2; }
$p = new Point();
$prop = 'x';
echo $p->$prop;
