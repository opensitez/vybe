<?php
// vybe-test: php/php_iterators_spl_array_iterator/test_php_limit_iterator_offset_and_count
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

$array = [10, 20, 30, 40, 50];
$it = new LimitIterator(new ArrayIterator($array), 1, 3); // offset 1, count 3
echo implode(",", iterator_to_array($it, false));

__vybe_check(ob_get_clean(), "20,30,40");
