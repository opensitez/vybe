<?php
// vybe-test: php/namespaces/namespace_trait
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Concerns;
trait Timestampable {
    private int $createdAt = 0;
    public function setCreatedAt(int $ts): void { $this->createdAt = $ts; }
    public function getCreatedAt(): int { return $this->createdAt; }
}

namespace Models;
use Concerns\Timestampable;
class Post {
    use Timestampable;
    public function __construct(public string $title) {}
}
$p = new Post('Hello');
$p->setCreatedAt(1000);
echo $p->getCreatedAt();
