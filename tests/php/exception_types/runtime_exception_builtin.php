<?php
// vybe-test: php/exception_types/runtime_exception_builtin
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

try {
    throw new RuntimeException('runtime issue');
} catch (RuntimeException $e) {
    echo $e->getMessage();
}
