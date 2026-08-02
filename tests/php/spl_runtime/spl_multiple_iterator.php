<?php
// vybe-test: php/spl_runtime/spl_multiple_iterator
// origin: languages/php/tests/php/test_spl_runtime.rs

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

$m = new MultipleIterator();
$m->attachIterator(new ArrayIterator([1]));
$m->attachIterator(new ArrayIterator([2]));
$m->rewind();
echo $m->valid() ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "yes");
