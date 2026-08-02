<?php
// vybe-test: php/php8_audit/php81_first_class_callable
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

$fn = strlen(...); echo $fn('hello');
