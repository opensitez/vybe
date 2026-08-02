<?php
// vybe-test: php/weak_references/weak_map_type_checking
// origin: languages/php/tests/php/test_weak_references.rs
// vybe-test-mode: compile

$map = new WeakMap();
echo ($map instanceof WeakMap) ? 'is WeakMap' : 'not WeakMap';
