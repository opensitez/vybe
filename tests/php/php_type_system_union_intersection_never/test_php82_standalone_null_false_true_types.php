<?php
// vybe-test: php/php_type_system_union_intersection_never/test_php82_standalone_null_false_true_types
// origin: languages/php/tests/php/test_php_type_system_union_intersection_never.rs
// vybe-test-mode: compile

function getNull(): null {
    return null;
}

function alwaysFalse(): false {
    return false;
}

function alwaysTrue(): true {
    return true;
}

echo is_null(getNull()) && !alwaysFalse() && alwaysTrue() ? "TYPES_OK" : "FAIL";
