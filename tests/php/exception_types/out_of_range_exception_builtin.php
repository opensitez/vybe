<?php
// vybe-test: php/exception_types/out_of_range_exception_builtin
// origin: languages/php/tests/php/test_exception_types.rs
// vybe-test-mode: compile

try {
    throw new OutOfRangeException('index out of range');
} catch (OutOfRangeException $e) {
    echo $e->getMessage();
}
