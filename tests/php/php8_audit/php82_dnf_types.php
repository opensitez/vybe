<?php
// vybe-test: php/php8_audit/php82_dnf_types
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

function foo((A&B)|C $x) {}
