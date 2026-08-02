<?php
// vybe-test: php/traits_deep/trait_property_multiple_classes
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait HasId {
    private static int $nextId = 0;
    private int $id;
    public function initId(): void { $this->id = ++self::$nextId; }
    public function getId(): int   { return $this->id; }
}
class UserA { use HasId; }
class UserB { use HasId; }
$a = new UserA(); $a->initId();
$b = new UserB(); $b->initId();
echo $a->getId() . ',' . $b->getId();
