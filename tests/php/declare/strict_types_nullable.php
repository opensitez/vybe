<?php
// vybe-test: php/declare/strict_types_nullable
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
function findUser(?int $id): ?string {
    if ($id === null) return null;
    return "user_$id";
}
echo findUser(5) ?? 'none';
echo findUser(null) ?? 'none';
