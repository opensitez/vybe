<?php
// vybe-test: php/references/static_variable_via_reference
// origin: languages/php/tests/php/test_references.rs
// vybe-test-mode: compile

function counter(): int {
    static $n = 0;
    return ++$n;
}
echo counter() . ',' . counter() . ',' . counter();
