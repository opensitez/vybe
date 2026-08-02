<?php
// vybe-test: php/scoping/static_vars_in_functions_isolated
// origin: languages/php/tests/php/test_scoping.rs
// vybe-test-mode: compile

function next_id(): int { static $id = 0; return ++$id; } echo next_id(); echo '-'; echo next_id();
