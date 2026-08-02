<?php
// vybe-test: php/traits_deep/trait_satisfies_interface
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

interface Identifiable { public function getId(): int; }
trait HasId {
    private int $id;
    public function __construct(int $id) { $this->id = $id; }
    public function getId(): int { return $this->id; }
}
class User implements Identifiable { use HasId; }
$u = new User(42);
echo $u->getId();
