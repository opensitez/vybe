<?php
// vybe-test: php/php_type_system_union_intersection_never/test_php_void_return_type_enforcement
// origin: languages/php/tests/php/test_php_type_system_union_intersection_never.rs
// vybe-test-mode: compile

function doWork(): void {
    // Return with no value allowed
    return;
}

doWork();
