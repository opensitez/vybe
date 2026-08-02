<?php
// vybe-test: php/arrayaccess_countable/iterator_aggregate_foreach
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

class NumberRange implements IteratorAggregate {
    public function __construct(private int $from, private int $to) {}
    public function getIterator(): ArrayIterator {
        return new ArrayIterator(range($this->from, $this->to));
    }
}
$r = new NumberRange(1, 4);
foreach ($r as $n) echo $n;

__vybe_check(ob_get_clean(), "1234");
