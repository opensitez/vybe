<?php
// vybe-test: php/intersection_types/intersection_return_type_both_methods_callable
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

interface Countable3 { public function count(): int; }
interface Iterable2 { public function toArray(): array; }
class Collection implements Countable3, Iterable2 {
    private array $data;
    public function __construct(array $data) { $this->data = $data; }
    public function count(): int { return count($this->data); }
    public function toArray(): array { return $this->data; }
}
function getCollection(): Countable3&Iterable2 { return new Collection([1,2,3]); }
$c = getCollection();
echo $c->count() . ',' . implode(',', $c->toArray());

__vybe_check(ob_get_clean(), "3,1,2,3");
