<?php
// vybe-test: php/cross_lang/runtime_exception_canonical
// origin: languages/php/tests/php/test_cross_lang.rs
// vybe-test-mode: compile

try {
    throw new RuntimeException('runtime error');
} catch (RuntimeException $e) {
    // Maps to canonical "RuntimeError" — catchable in Python as RuntimeError
    echo $e;
}
