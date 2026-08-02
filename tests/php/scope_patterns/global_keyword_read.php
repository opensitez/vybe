<?php
// vybe-test: php/scope_patterns/global_keyword_read
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

$counter = 10;
function readGlobal(): int {
    global $counter;
    return $counter;
}
echo readGlobal();
