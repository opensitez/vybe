<?php
// vybe-test: php/weak_references/weak_map_count
// origin: languages/php/tests/php/test_weak_references.rs
// vybe-test-mode: compile

$map = new WeakMap();
$a = new stdClass();
$b = new stdClass();
$c = new stdClass();
$map[$a] = 1;
$map[$b] = 2;
$map[$c] = 3;
echo count($map);
