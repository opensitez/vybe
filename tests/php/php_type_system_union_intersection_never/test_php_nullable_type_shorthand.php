<?php
// vybe-test: php/php_type_system_union_intersection_never/test_php_nullable_type_shorthand
// origin: languages/php/tests/php/test_php_type_system_union_intersection_never.rs
// vybe-test-mode: compile

function findName(?int $id): ?string {
    if ($id === 1) return "Alice";
    return null;
}

echo findName(1) . " " . (findName(2) ?? "NULL");
