<?php
// vybe-test: php/magic_constants/class_constant_enum
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

enum Color { case Red; case Blue; }
echo Color::class;
