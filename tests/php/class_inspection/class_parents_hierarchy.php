<?php
// vybe-test: php/class_inspection/class_parents_hierarchy
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class A {}
class B extends A {}
class C extends B {}
$parents = class_parents('C');
echo isset($parents['B']) ? 'yes' : 'no';
echo isset($parents['A']) ? 'yes' : 'no';
