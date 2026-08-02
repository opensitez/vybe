<?php
// vybe-test: php/php8_audit/php81_enum_basic
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

enum Status { case Active; case Inactive; } $s = Status::Active;
