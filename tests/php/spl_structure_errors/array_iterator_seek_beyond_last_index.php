<?php
// vybe-test: php/spl_structure_errors/array_iterator_seek_beyond_last_index
// origin: languages/php/tests/php/test_spl_structure_errors.rs

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

$it = new ArrayIterator([10, 20]);
try { $it->seek(5); echo $it->current(); }
catch (OutOfBoundsException $e) { echo 'seek-oob'; }

__vybe_check(ob_get_clean(), "seek-oob");
