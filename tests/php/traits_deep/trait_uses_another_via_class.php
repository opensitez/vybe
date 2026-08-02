<?php
// vybe-test: php/traits_deep/trait_uses_another_via_class
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait HasTimestamps {
    private int $createdAt = 0;
    private int $updatedAt = 0;
    public function touch(): void { $this->updatedAt = time(); }
    public function getUpdatedAt(): int { return $this->updatedAt; }
}
trait HasSoftDelete {
    private ?int $deletedAt = null;
    public function delete(): void { $this->deletedAt = time(); }
    public function isDeleted(): bool { return $this->deletedAt !== null; }
}
class Post {
    use HasTimestamps, HasSoftDelete;
    public function __construct(public string $title) {}
}
$p = new Post('Hello World');
$p->delete();
echo $p->isDeleted() ? 'deleted' : 'active';
