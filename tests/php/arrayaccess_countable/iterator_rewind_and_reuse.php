<?php
// vybe-test: php/arrayaccess_countable/iterator_rewind_and_reuse
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

class Words implements IteratorAggregate {
    public function getIterator(): ArrayIterator {
        return new ArrayIterator(['foo','bar','baz']);
    }
}
$w = new Words;
foreach ($w as $v) echo $v[0];
foreach ($w as $v) echo strtoupper($v[0]);

__vybe_check(ob_get_clean(), "fbbFBB");
