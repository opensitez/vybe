<?php
// vybe-test: php/scope_patterns/function_cannot_see_outer_variable
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$secret = 42;
function noAccess(): mixed {
    return isset($secret) ? $secret : null;
}
echo noAccess() === null ? 'hidden' : 'leaked';
