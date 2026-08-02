<?php
// vybe-test: php/trait_conflict_resolution/trait_same_property_compatible_redeclaration
// origin: languages/php/tests/php/test_trait_conflict_resolution.rs
// vybe-test-mode: compile

trait HasId { public int $id = 0; }
class Entity {
    use HasId;
    public int $id = 0;
}
