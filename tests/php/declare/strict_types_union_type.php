<?php
// vybe-test: php/declare/strict_types_union_type
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function formatId(int|string $id): string { return "ID:$id"; }
echo formatId(42);
echo formatId("uuid-abc");
