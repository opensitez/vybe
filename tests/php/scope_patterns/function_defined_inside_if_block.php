<?php
// vybe-test: php/scope_patterns/function_defined_inside_if_block
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$flag = true;
if ($flag) {
    function conditionalFn(): int { return 99; }
}
echo conditionalFn();
