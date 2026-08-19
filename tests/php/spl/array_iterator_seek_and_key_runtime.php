<?php
// vybe-test: php/spl/array_iterator_seek_and_key_runtime
// origin: languages/php/tests/php/test_spl.rs

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

$it = new ArrayIterator(['a' => 1, 'b' => 2, 'c' => 3]);
$it->seek(2);
echo $it->key();
echo '|';
echo $it->current();

__vybe_check(ob_get_clean(), "c|3");
