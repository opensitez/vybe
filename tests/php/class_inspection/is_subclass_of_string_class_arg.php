<?php
// vybe-test: php/class_inspection/is_subclass_of_string_class_arg
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class Shape {}
class Circle extends Shape {}
echo is_subclass_of('Circle', 'Shape') ? 'yes' : 'no';
echo is_subclass_of('Shape', 'Circle') ? 'yes' : 'no';
