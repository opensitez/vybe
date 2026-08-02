<?php
// vybe-test: php/gaps_audit/nullable_type_parse
// origin: languages/php/tests/php/test_gaps_audit.rs
// vybe-test-mode: compile

function foo(?int $x): ?string { return null; }
