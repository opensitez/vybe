<?php
// vybe-test: php/type_juggling/cast_object
// origin: languages/php/tests/php/test_type_juggling.rs
// vybe-test-mode: compile

$arr = ['name' => 'Alice', 'age' => 30];
$obj = (object) $arr;
echo $obj->name . ':' . $obj->age;
