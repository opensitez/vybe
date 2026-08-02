<?php
// vybe-test: php/traits_deep/trait_anonymous_class
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait Taggable {
    private array $tags = [];
    public function addTag(string $tag): void { $this->tags[] = $tag; }
    public function getTags(): array { return $this->tags; }
}
$obj = new class { use Taggable; };
$obj->addTag('php');
$obj->addTag('oop');
echo implode(',', $obj->getTags());
