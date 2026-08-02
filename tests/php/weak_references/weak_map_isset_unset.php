<?php
// vybe-test: php/weak_references/weak_map_isset_unset
// origin: languages/php/tests/php/test_weak_references.rs
// vybe-test-mode: compile

$map = new WeakMap();
$obj = new stdClass();
echo isset($map[$obj]) ? 'set' : 'not set';
$map[$obj] = 'value';
echo isset($map[$obj]) ? 'set' : 'not set';
unset($map[$obj]);
echo isset($map[$obj]) ? 'set' : 'not set';
