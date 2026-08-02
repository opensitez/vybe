<?php
// vybe-test: php/covariant_return_types/never_return_type_function_exits
// origin: languages/php/tests/php/test_covariant_return_types.rs
// vybe-test-mode: compile

function abort(int $code): never {
    throw new RuntimeException("abort: $code");
}
