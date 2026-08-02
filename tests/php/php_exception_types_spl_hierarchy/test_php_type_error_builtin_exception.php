<?php
// vybe-test: php/php_exception_types_spl_hierarchy/test_php_type_error_builtin_exception
// origin: languages/php/tests/php/test_php_exception_types_spl_hierarchy.rs
// vybe-test-mode: compile

try {
    throw new TypeError("Type mismatch error");
} catch (Error $e) {
    echo "Caught Error: " . $e->getMessage();
}
