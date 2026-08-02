<?php
// vybe-test: php/covariant_return_types/intersection_return_type_both_interfaces_satisfied
// origin: languages/php/tests/php/test_covariant_return_types.rs

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

interface Countable2 { public function count(): int; }
interface Listable { public function toList(): array; }
class Collection implements Countable2, Listable {
    private array $items;
    public function __construct(array $items) { $this->items = $items; }
    public function count(): int { return count($this->items); }
    public function toList(): array { return $this->items; }
}
function getCollection(): Countable2&Listable { return new Collection([1,2,3]); }
$c = getCollection();
echo $c->count();

__vybe_check(ob_get_clean(), "3");
