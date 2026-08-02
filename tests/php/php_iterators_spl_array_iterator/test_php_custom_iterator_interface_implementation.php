<?php
// vybe-test: php/php_iterators_spl_array_iterator/test_php_custom_iterator_interface_implementation
// origin: languages/php/tests/php/test_php_iterators_spl_array_iterator.rs

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
    private int $position = 0;
    public function __construct(private int $start, private int $end) {
        $this->position = $start;
    }
    public function current(): mixed { return $this->position; }
    public function key(): mixed { return $this->position - $this->start; }
    public function next(): void { $this->position++; }
    public function rewind(): void { $this->position = $this->start; }
    public function valid(): bool { return $this->position <= $this->end; }
}

$range = new NumberRange(10, 12);
$out = [];
foreach ($range as $k => $v) {
    $out[] = "$k:$v";
}
echo implode(", ", $out);

__vybe_check(ob_get_clean(), "0:10, 1:11, 2:12");
