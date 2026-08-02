<?php
// vybe-test: php/arrayaccess_countable/iterator_interface_manual
// origin: languages/php/tests/php/test_arrayaccess_countable.rs

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

class Countdown implements Iterator {
    private int $cur;
    public function __construct(private int $start) { $this->cur = $start; }
    public function current(): int { return $this->cur; }
    public function key(): int { return $this->start - $this->cur; }
    public function next(): void { $this->cur--; }
    public function rewind(): void { $this->cur = $this->start; }
    public function valid(): bool { return $this->cur > 0; }
}
foreach (new Countdown(3) as $n) echo $n;

__vybe_check(ob_get_clean(), "321");
