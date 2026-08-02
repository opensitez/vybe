<?php
// vybe-test: php/php7/php71_nullable_type
// origin: languages/php/tests/php/test_php7.rs
// vybe-test-mode: compile

function foo(?int $x): ?string { return null; }
