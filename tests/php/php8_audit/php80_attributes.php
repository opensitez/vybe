<?php
// vybe-test: php/php8_audit/php80_attributes
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

#[Attr] #[Attr2('arg')] function foo() {}
