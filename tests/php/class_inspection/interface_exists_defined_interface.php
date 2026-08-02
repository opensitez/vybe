<?php
// vybe-test: php/class_inspection/interface_exists_defined_interface
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

interface Drawable { public function draw(): void; }
echo interface_exists('Drawable') ? 'yes' : 'no';
echo interface_exists('Nonexistent') ? 'yes' : 'no';
