<?php
// vybe-test: php/php_type_system_union_intersection_never/test_php82_dnf_disjunctive_normal_form_types
// origin: languages/php/tests/php/test_php_type_system_union_intersection_never.rs
// vybe-test-mode: compile

interface HasId { public function getId(): int; }
interface HasName { public function getName(): string; }

class Entity implements HasId, HasName {
    public function getId(): int { return 1; }
    public function getName(): string { return "e1"; }
}

function processEntity((HasId&HasName)|string $target): string {
    if (is_string($target)) return $target;
    return $target->getName();
}

echo processEntity(new Entity()) . " " . processEntity("literal");
