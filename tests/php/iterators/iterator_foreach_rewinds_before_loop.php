<?php
// vybe-test: php/iterators/iterator_foreach_rewinds_before_loop
// origin: languages/php/tests/php/test_iterators.rs

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

class Letters implements Iterator {
    private int $p = 0;
    private array $v = ['a', 'b'];
    public function current(): string { return $this->v[$this->p]; }
    public function key(): int { return $this->p; }
    public function next(): void { $this->p++; }
    public function rewind(): void { $this->p = 0; }
    public function valid(): bool { return $this->p < count($this->v); }
}
$it = new Letters();
foreach ($it as $ch) { echo $ch; }

__vybe_check(ob_get_clean(), "ab");
