<?php
// vybe-test: php/php8_audit/php81_enum_backed
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

enum Color: string { case Red = 'red'; case Blue = 'blue'; } echo Color::Red->value;
