<?php
// vybe-test: php/string_formatting/sprintf_sql_like_pattern
// origin: languages/php/tests/php/test_string_formatting.rs
// vybe-test-mode: compile

// Template-based string building (illustrative, not real SQL)
function buildQuery(string $table, int $id): string {
    return sprintf("SELECT * FROM `%s` WHERE id = %d LIMIT 1", $table, $id);
}
echo buildQuery('users', 42);
echo "\n";
