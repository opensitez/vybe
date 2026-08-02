<?php
// vybe-test: php/class_inspection/get_parent_class_with_inheritance
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class Base {}
class Child extends Base {}
echo get_parent_class(new Child());
echo get_parent_class('Child');
