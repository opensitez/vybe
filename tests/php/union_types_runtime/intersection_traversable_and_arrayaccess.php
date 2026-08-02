<?php
// vybe-test: php/union_types_runtime/intersection_traversable_and_arrayaccess
// origin: languages/php/tests/php/test_union_types_runtime.rs

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

class C implements Traversable, ArrayAccess, Countable, IteratorAggregate {
    public function __construct(private array $d) {}
    public function getIterator(): Traversable { yield from $this->d; }
    public function offsetExists($k): bool { return isset($this->d[$k]); }
    public function offsetGet($k): mixed { return $this->d[$k]; }
    public function offsetSet($k, $v): void { $this->d[$k] = $v; }
    public function offsetUnset($k): void { unset($this->d[$k]); }
    public function count(): int { return count($this->d); }
}
function read(ArrayAccess&Countable $x): int { return count($x); }
echo read(new C([1]));

__vybe_check(ob_get_clean(), "1");
