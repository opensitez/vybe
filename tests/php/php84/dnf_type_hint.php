<?php
// vybe-test: php/php84/dnf_type_hint
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

interface Countable2 { public function count2(): int; }
interface Serializable2 { public function serialize2(): string; }
class Set implements Countable2 {
    private array $items = [];
    public function count2(): int { return count($this->items); }
    public function add(mixed $v): void { $this->items[] = $v; }
}
function describe((Countable2&Serializable2)|null $obj): string {
    if ($obj === null) return 'null';
    return 'count=' . $obj->count2();
}
// Pass null (valid for (C&S)|null)
echo describe(null);
