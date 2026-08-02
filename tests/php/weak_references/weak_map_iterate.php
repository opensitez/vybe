<?php
// vybe-test: php/weak_references/weak_map_iterate
// origin: languages/php/tests/php/test_weak_references.rs
// vybe-test-mode: compile

$map = new WeakMap();
$objs = [];
for ($i = 0; $i < 3; $i++) {
    $obj = new stdClass();
    $obj->n = $i;
    $objs[] = $obj;
    $map[$obj] = "value_$i";
}
$vals = [];
foreach ($map as $k => $v) { $vals[] = $v; }
sort($vals);
echo implode(',', $vals);
