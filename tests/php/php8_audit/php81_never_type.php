<?php
// vybe-test: php/php8_audit/php81_never_type
// origin: languages/php/tests/php/test_php8_audit.rs
// vybe-test-mode: compile

function fail(): never { throw new Exception('x'); }
