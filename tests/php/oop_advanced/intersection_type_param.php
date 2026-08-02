<?php
// vybe-test: php/oop_advanced/intersection_type_param
// origin: languages/php/tests/php/test_oop_advanced.rs

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

interface Countable2 {
    public function count(): int;
}
interface Stringable2 {
    public function __toString(): string;
}
class Items implements Countable2, Stringable2 {
    private array $data;
    public function __construct(array $data) { $this->data = $data; }
    public function count(): int { return count($this->data); }
    public function __toString(): string { return implode(",", $this->data); }
}
function describe(Countable2&Stringable2 $obj): void {
    echo $obj->count(), "\n";
    echo $obj, "\n";
}
describe(new Items(["a", "b", "c"]));

__vybe_check(ob_get_clean(), "3\na,b,c");
