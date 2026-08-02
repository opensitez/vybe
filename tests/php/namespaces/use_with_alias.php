<?php
// vybe-test: php/namespaces/use_with_alias
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Library;
class Collection {
    private array $items = [];
    public function add(mixed $v): void { $this->items[] = $v; }
    public function count(): int { return count($this->items); }
}

namespace App;
use Library\Collection as List_;
$list = new List_();
$list->add(1); $list->add(2);
echo $list->count();
