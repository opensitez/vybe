<?php
// vybe-test: php/php_type_system_union_intersection_never/test_php81_type_variance_in_methods
// origin: languages/php/tests/php/test_php_type_system_union_intersection_never.rs
// vybe-test-mode: compile

class ParentService {
    public function handle(int|float $num): int|float|string { return $num; }
}

class ChildService extends ParentService {
    // Parameter type widening & Return type narrowing
    public function handle(int|float|string $num): int|float { return 42; }
}

$cs = new ChildService();
echo $cs->handle("123");
