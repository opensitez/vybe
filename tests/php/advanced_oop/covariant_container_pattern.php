<?php
// vybe-test: php/advanced_oop/covariant_container_pattern
// origin: languages/php/tests/php/test_advanced_oop.rs

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

class Box {
    public function __construct(private mixed $value) {}
    public function get(): mixed { return $this->value; }
    public function map(callable $fn): static { return new static($fn($this->value)); }
}
$result = (new Box(5))->map(fn($n) => $n * 2)->map(fn($n) => "value:$n")->get();
echo $result;

__vybe_check(ob_get_clean(), "value:10");
