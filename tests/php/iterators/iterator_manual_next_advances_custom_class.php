<?php
// vybe-test: php/iterators/iterator_manual_next_advances_custom_class
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

class Three implements Iterator {
    private int $i = 0;
    public function current(): int { return $this->i; }
    public function key(): int { return $this->i; }
    public function next(): void { $this->i++; }
    public function rewind(): void { $this->i = 0; }
    public function valid(): bool { return $this->i < 3; }
}
$c = new Three();
$out = [];
while ($c->valid()) { $out[] = $c->current(); $c->next(); }
echo implode(',', $out);

__vybe_check(ob_get_clean(), "0,1,2");
