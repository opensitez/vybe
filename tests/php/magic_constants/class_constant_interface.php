<?php
// vybe-test: php/magic_constants/class_constant_interface
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

interface Drawable { public function draw(): void; }
echo Drawable::class;
