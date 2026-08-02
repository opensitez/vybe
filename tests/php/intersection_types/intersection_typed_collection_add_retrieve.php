<?php
// vybe-test: php/intersection_types/intersection_typed_collection_add_retrieve
// origin: languages/php/tests/php/test_intersection_types.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

interface Hashable { public function hash(): string; }
interface Comparable2 { public function compareTo(mixed $other): int; }
class SortedKey implements Hashable, Comparable2 {
    public function __construct(private string $key) {}
    public function hash(): string { return md5($this->key); }
    public function compareTo(mixed $other): int { return strcmp($this->key, $other->key); }
}
class TypedSet {
    private array $items = [];
    public function add(Hashable&Comparable2 $item): void { $this->items[$item->hash()] = $item; }
    public function count(): int { return count($this->items); }
}
$set = new TypedSet();
$set->add(new SortedKey('foo'));
$set->add(new SortedKey('bar'));
echo $set->count();

__vybe_check(ob_get_clean(), "2");
