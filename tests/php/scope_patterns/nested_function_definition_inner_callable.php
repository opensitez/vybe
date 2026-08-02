<?php
// vybe-test: php/scope_patterns/nested_function_definition_inner_callable
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

function outer(): void {
    function inner(): string { return 'inside'; }
}
outer();
echo inner();
