<?php
// vybe-test: php/php_iterators_spl_array_iterator/test_php_iterator_aggregate_get_iterator
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

class Collection implements IteratorAggregate {
    private array $items = ["a", "b", "c"];
    public function getIterator(): Traversable {
        return new ArrayIterator($this->items);
    }
}

$c = new Collection();
echo implode("-", iterator_to_array($c));

__vybe_check(ob_get_clean(), "a-b-c");
