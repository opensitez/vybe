<?php
// vybe-test: php/class_inspection/class_exists_no_autoload
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class RegisteredClass {}
echo class_exists('RegisteredClass', false) ? 'yes' : 'no';
echo class_exists('UnregisteredClass', false) ? 'yes' : 'no';
