<?php
// vybe-test: php/variable_variables/dynamic_property_set
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Box { public int $width = 0; public int $height = 0; }
$b = new Box();
foreach (['width' => 10, 'height' => 5] as $prop => $val) {
    $b->$prop = $val;
}
echo $b->width . 'x' . $b->height;
