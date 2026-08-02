<?php
// vybe-test: php/magic_constants/class_constant_on_object
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

class Widget { public string $type = 'button'; }
$w = new Widget();
echo $w::class;
