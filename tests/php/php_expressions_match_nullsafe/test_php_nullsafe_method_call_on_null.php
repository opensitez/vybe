<?php
// vybe-test: php/php_expressions_match_nullsafe/test_php_nullsafe_method_call_on_null
// origin: languages/php/tests/php/test_php_expressions_match_nullsafe.rs
// vybe-test-mode: compile

class Repository {
    public function find(int $id): ?object {
        return null;
    }
}

$repo = new Repository();
$name = $repo->find(10)?->getName();
echo $name ?? "no_object";
