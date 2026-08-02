<?php
// vybe-test: php/serialization_advanced/serialize_object_basic
// origin: languages/php/tests/php/test_serialization_advanced.rs
// vybe-test-mode: compile

class Point { public function __construct(public int $x, public int $y) {} }
$p = new Point(3, 7);
$s = serialize($p);
$p2 = unserialize($s);
echo $p2->x . ',' . $p2->y;
