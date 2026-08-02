<?php
// vybe-test: php/oop_advanced/object_implements_iterator
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

class NumberRange implements Iterator {
    private int $current;
    public function __construct(
        private int $start,
        private int $end,
    ) {
        $this->current = $start;
    }
    public function current(): int  { return $this->current; }
    public function key(): int      { return $this->current - $this->start; }
    public function next(): void    { $this->current++; }
    public function rewind(): void  { $this->current = $this->start; }
    public function valid(): bool   { return $this->current <= $this->end; }
}
$range = new NumberRange(1, 5);
$vals = [];
foreach ($range as $k => $v) {
    $vals[] = "$k:$v";
}
echo implode(",", $vals), "\n";

__vybe_check(ob_get_clean(), "0:1,1:2,2:3,3:4,4:5");
